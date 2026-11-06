<#
.SYNOPSIS
    Meeting Cleanup - Microsoft Graph: token, requests and the $batch scheduler (dot-sourced by MeetingCleanup.psm1).

.DESCRIPTION
    Application permissions (app-only), as for every mailbox of the tenant:
      Calendars.ReadWrite   read the calendars, remove a copy, cancel a meeting (Calendars.Read: report only)
      User.Read.All         the addresses of the organizer, the list of every mailbox
      Place.Read.All        the room mailboxes (places API)
      GroupMember.Read.All  the members of a group invited to a meeting
    Only Calendars.ReadWrite (or Calendars.Read for a report) is required; Calendars.ReadWrite.All (Calendars.Read.All)
    counts as well (measured: Exchange Online accepts it for the events). Without the others the tool says
    what it cannot do and goes on.

    Token: certificate (client assertion built here, no module needed, recommended) or client secret
    (environment variable or typed, never written). It is renewed 5 minutes before it expires.

    Requests go through one transport (Start-MclGraphSend / Complete-MclGraphSend, replaced by a simulated
    tenant in the tests). Invoke-MclGraphBatch sends many requests in $batch calls of 20, several calls at
    once (Graph.MaxConcurrency), never more than 4 requests at a time for the same mailbox (limit of
    Exchange Online), retries 429 / 5xx after the Retry-After delay, and can follow the pages of a list.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.3.1
#>

$script:GraphRoot = 'https://graph.microsoft.com/v1.0'
$script:LoginHost = 'https://login.microsoftonline.com'
$script:Http = $null
# Exchange Online processes at most 4 requests at a time for one mailbox and one application.
$script:MailboxConcurrency = 4
$script:BatchSize = 20

function ConvertTo-MclBase64Url { param([byte[]]$Bytes) [Convert]::ToBase64String($Bytes).TrimEnd('=').Replace('+', '-').Replace('/', '_') }

function Get-MclTokenClaims {
    <# Payload of a JWT (no signature check: only used to read tid, roles and the application name). #>
    param([Parameter(Mandatory = $true)][string]$Token)
    $parts = $Token.Split('.')
    if ($parts.Count -lt 2) { throw 'The access token is not a JWT.' }
    $p = $parts[1].Replace('-', '+').Replace('_', '/')
    while ($p.Length % 4) { $p += '=' }
    return [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($p)) | ConvertFrom-Json
}

function Get-MclCertificate {
    <# Certificate with its private key, from Cert:\CurrentUser\My or Cert:\LocalMachine\My. #>
    param([Parameter(Mandatory = $true)][string]$Thumbprint)
    $thumb = $Thumbprint.Trim().ToUpperInvariant()
    foreach ($store in 'Cert:\CurrentUser\My', 'Cert:\LocalMachine\My') {
        $cert = Get-Item -LiteralPath (Join-Path $store $thumb) -ErrorAction SilentlyContinue
        if ($cert) {
            if (-not $cert.HasPrivateKey) { throw "Certificate $thumb found in $store without its private key: import the .pfx (not the .cer) for the account that runs the tool." }
            if ($cert.NotAfter -lt (Get-Date)) { throw "Certificate $thumb expired on $($cert.NotAfter.ToString('yyyy-MM-dd')). Upload a new certificate to the application and update Authentication.CertificateThumbprint." }
            return $cert
        }
    }
    throw "Certificate $thumb not found in Cert:\CurrentUser\My nor Cert:\LocalMachine\My (account $([Environment]::UserName)). Developer guide, chapter 5 'Application'."
}

function New-MclClientAssertion {
    <# Client assertion (RFC 7523) signed with the certificate: RS256, header x5t = SHA-1 thumbprint, valid 10 minutes. #>
    param(
        [Parameter(Mandatory = $true)][Security.Cryptography.X509Certificates.X509Certificate2]$Certificate,
        [Parameter(Mandatory = $true)][string]$TenantId,
        [Parameter(Mandatory = $true)][string]$AppId
    )
    $now = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
    $header = [ordered]@{ alg = 'RS256'; typ = 'JWT'; x5t = (ConvertTo-MclBase64Url $Certificate.GetCertHash()) } | ConvertTo-Json -Compress
    $claims = [ordered]@{ aud = "$($script:LoginHost)/$TenantId/oauth2/v2.0/token"; iss = $AppId; sub = $AppId; jti = [guid]::NewGuid().ToString(); nbf = $now - 60; iat = $now; exp = $now + 600 } | ConvertTo-Json -Compress
    $unsigned = (ConvertTo-MclBase64Url ([Text.Encoding]::UTF8.GetBytes($header))) + '.' + (ConvertTo-MclBase64Url ([Text.Encoding]::UTF8.GetBytes($claims)))
    $rsa = [Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($Certificate)
    if (-not $rsa) { throw "Certificate $($Certificate.Thumbprint): the private key is not an RSA key, or this account cannot use it." }
    try { $signature = $rsa.SignData([Text.Encoding]::ASCII.GetBytes($unsigned), [Security.Cryptography.HashAlgorithmName]::SHA256, [Security.Cryptography.RSASignaturePadding]::Pkcs1) }
    finally { $rsa.Dispose() }
    return "$unsigned.$(ConvertTo-MclBase64Url $signature)"
}

function Get-MclEntraErrorHint {
    <# A sentence for the most frequent Microsoft Entra sign-in errors. #>
    param([string]$Message)
    $hints = [ordered]@{
        'AADSTS700016'  = 'the application ID is not found in this tenant (Authentication.AppId, Tenant.TenantId).'
        'AADSTS90002'   = 'the tenant is not found (Tenant.TenantId).'
        'AADSTS700027'  = 'the certificate is not registered on the application, or not the right one (thumbprint).'
        'AADSTS7000215' = 'the client secret is not valid for this application.'
        'AADSTS7000222' = 'the client secret has expired: create a new one, or move to a certificate.'
        'AADSTS700024'  = 'the clock of this computer is not on time (the assertion is outside its validity).'
        'AADSTS53003'   = 'blocked by Conditional Access for workload identities.'
    }
    foreach ($code in $hints.Keys) { if ($Message -match $code) { return "$code - $($hints[$code])" } }
    return $null
}

function Get-MclAppToken {
    <# App-only token (client credentials): certificate or secret, for Graph or another resource (Scope). Returns @{ Token; ExpiresUtc }. #>
    param([Parameter(Mandatory = $true)][hashtable]$Settings, $Certificate, [Security.SecureString]$Secret, [string]$Scope = 'https://graph.microsoft.com/.default')
    $body = @{ client_id = $Settings.AppId; scope = $Scope; grant_type = 'client_credentials' }
    if ($Certificate) {
        $body['client_assertion_type'] = 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'
        $body['client_assertion'] = New-MclClientAssertion -Certificate $Certificate -TenantId $Settings.TenantId -AppId $Settings.AppId
    }
    else { $body['client_secret'] = [Net.NetworkCredential]::new('', $Secret).Password }
    try { $r = Invoke-RestMethod -Method Post -Uri "$($script:LoginHost)/$($Settings.TenantId)/oauth2/v2.0/token" -Body $body -ErrorAction Stop }
    catch {
        $text = try { ($_.ErrorDetails.Message | ConvertFrom-Json).error_description } catch { $null }
        if (-not $text) { $text = $_.Exception.Message }
        $hint = Get-MclEntraErrorHint $text
        throw ('Microsoft Entra sign-in of the application failed{0}: {1}' -f $(if ($hint) { " ($hint)" } else { '' }), ($text -split "`r?`n")[0])
    }
    finally { $body.Clear() }
    return @{ Token = $r.access_token; ExpiresUtc = [datetime]::UtcNow.AddSeconds([int]$r.expires_in) }
}

function Connect-MclGraph {
    <#
    .SYNOPSIS
        Obtains the first token, checks the tenant and the permissions, and keeps the connection for the run.
    .PARAMETER Secret
        ClientSecret mode: the secret (window). Otherwise the environment variable, else a prompt.
    .PARAMETER Action
        Report needs Calendars.Read or Calendars.ReadWrite; Remove and Cancel need Calendars.ReadWrite.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][hashtable]$Settings, [Security.SecureString]$Secret, [string]$Action = 'Report')

    $check = Test-MclConfiguration -Configuration $Settings -ForConnection
    if (-not $check.IsValid) { throw ("Connection settings:`n - " + ($check.Problems -join "`n - ")) }
    $connection = @{ TenantId = $Settings.TenantId; AppId = $Settings.AppId; Mode = $Settings.AuthMode; Settings = $Settings; Certificate = $null; Secret = $null; Token = $null; ExpiresUtc = [datetime]::MinValue }
    if ($Settings.AuthMode -eq 'Certificate') {
        $connection.Certificate = Get-MclCertificate $Settings.CertificateThumbprint
        if ($connection.Certificate.NotAfter -lt (Get-Date).AddDays(30)) { Write-MclItem Warn ('Certificate {0} expires on {1}: renew it.' -f $connection.Certificate.Thumbprint, $connection.Certificate.NotAfter.ToString('yyyy-MM-dd')) }
    }
    else {
        if (-not $Secret) {
            $value = [Environment]::GetEnvironmentVariable($Settings.ClientSecretVariable)
            if ($value) { $Secret = ConvertTo-SecureString $value -AsPlainText -Force; $value = $null }
            elseif (-not $script:Ui -and [Environment]::UserInteractive -and -not [Console]::IsInputRedirected) { $Secret = Read-Host -AsSecureString "      Client secret of application $($Settings.AppId)" }
            else { throw "ClientSecret mode: the environment variable $($Settings.ClientSecretVariable) is empty and no secret was typed." }
        }
        if (-not $Secret -or $Secret.Length -eq 0) { throw 'No client secret given.' }
        $connection.Secret = $Secret
    }
    $script:Graph = $connection
    Update-MclToken -Force

    $claims = Get-MclTokenClaims $connection.Token
    if ([string]$Settings.TenantId -match $script:GuidPattern -and $claims.tid -ne $Settings.TenantId) {
        $script:Graph = $null
        throw "Connected to tenant $($claims.tid), but Tenant.TenantId is $($Settings.TenantId). Nothing was read."
    }
    $roles = @($claims.PSObject.Properties['roles'] | ForEach-Object { $_.Value })
    $connection.Roles = $roles
    $connection.TenantGuid = [string]$claims.tid
    $connection.AppName = [string](Get-MclProperty $claims 'app_displayname')
    # Calendars.ReadWrite.All and Calendars.Read.All are documented for the work hours and locations of the users, but
    # Exchange Online accepts them for the events too (lab, 2026-10-07: .ReadWrite.All reads, creates and deletes
    # events; .Read.All reads them, a creation is refused 403): they count as Calendars.ReadWrite / Calendars.Read.
    $connection.CanWrite = @($roles | Where-Object { $_ -in 'Calendars.ReadWrite', 'Calendars.ReadWrite.All' }).Count -gt 0
    $connection.CanRead = $connection.CanWrite -or @($roles | Where-Object { $_ -in 'Calendars.Read', 'Calendars.Read.All' }).Count -gt 0
    $connection.CanReadUsers = @($roles | Where-Object { $_ -in 'User.Read.All', 'User.ReadWrite.All', 'Directory.Read.All', 'Directory.ReadWrite.All' }).Count -gt 0
    $connection.CanReadPlaces = @($roles | Where-Object { $_ -in 'Place.Read.All', 'Place.ReadWrite.All' }).Count -gt 0
    $connection.CanReadGroups = @($roles | Where-Object { $_ -in 'GroupMember.Read.All', 'Group.Read.All', 'Group.ReadWrite.All', 'Directory.Read.All', 'Directory.ReadWrite.All' }).Count -gt 0
    $grant = "Entra admin center > App registrations > $(if ($connection.AppName) { $connection.AppName } else { $Settings.AppId }) > API permissions > Microsoft Graph > Application permissions, then 'Grant admin consent'"
    if (-not $connection.CanRead) {
        $script:Graph = $null
        throw "The application has no application permission Calendars.ReadWrite (or Calendars.ReadWrite.All) with admin consent (roles in the token: $(if ($roles.Count) { $roles -join ', ' } else { 'none' })). $grant."
    }
    if ($Action -ne 'Report' -and -not $connection.CanWrite) {
        $script:Graph = $null
        throw "The application has Calendars.Read only: it can report but not remove, cancel or transfer. Add Calendars.ReadWrite ($grant)."
    }
    return [pscustomobject]$connection
}

function Update-MclToken {
    <# Renews the token when it expires within 5 minutes (or at once with -Force, after a 401). #>
    param([switch]$Force)
    $g = $script:Graph
    if (-not $g) { throw 'Not connected to Microsoft Graph (Connect-MclGraph).' }
    if (-not $Force -and $g.Token -and $g.ExpiresUtc -gt [datetime]::UtcNow.AddMinutes(5)) { return }
    if ($g.ContainsKey('Renew') -and $g.Renew) { $t = & $g.Renew }
    else { $t = Get-MclAppToken -Settings $g.Settings -Certificate $g.Certificate -Secret $g.Secret }
    $g.Token = $t.Token
    $g.ExpiresUtc = $t.ExpiresUtc
    Write-MclLog 'INFO' ('Access token obtained, valid until {0:HH:mm:ss} UTC.' -f $t.ExpiresUtc)
}

function Get-MclHttpClient {
    if (-not $script:Http) {
        $handler = [Net.Http.SocketsHttpHandler]::new()
        $handler.MaxConnectionsPerServer = 32
        $handler.AutomaticDecompression = [Net.DecompressionMethods]::GZip -bor [Net.DecompressionMethods]::Deflate
        $script:Http = [Net.Http.HttpClient]::new($handler)
        $script:Http.Timeout = [TimeSpan]::FromSeconds($(if ($script:Graph) { [int]$script:Graph.Settings.TimeoutSeconds } else { 120 }))
        $script:Http.DefaultRequestHeaders.UserAgent.ParseAdd("MeetingCleanup/$($script:ToolVersion)")
    }
    return $script:Http
}

function Start-MclGraphSend {
    <#
        Sends one HTTP request to Graph without waiting. Returns a handle for Complete-MclGraphSend.
        The only function that touches the network for Graph: the tests replace it with a simulated tenant.
    #>
    param([Parameter(Mandatory = $true)][string]$Method, [Parameter(Mandatory = $true)][string]$Url, [string]$Body)
    $request = [Net.Http.HttpRequestMessage]::new([Net.Http.HttpMethod]::new($Method), $Url)
    $request.Headers.Authorization = [Net.Http.Headers.AuthenticationHeaderValue]::new('Bearer', $script:Graph.Token)
    [void]$request.Headers.TryAddWithoutValidation('client-request-id', [guid]::NewGuid().ToString())
    if ($Body) { $request.Content = [Net.Http.StringContent]::new($Body, [Text.Encoding]::UTF8, 'application/json') }
    [pscustomobject]@{ Task = (Get-MclHttpClient).SendAsync($request); Request = $request; Response = $null }
}

function Test-MclGraphSendDone { param($Handle) return ($null -ne $Handle.Response) -or $Handle.Task.IsCompleted }

function Complete-MclGraphSend {
    <# Result of a sent request: @{ Status; RetryAfter; Content }. Status 0 = no answer (network, timeout). #>
    param([Parameter(Mandatory = $true)]$Handle)
    if ($null -ne $Handle.Response) { return $Handle.Response }
    $response = $null
    try {
        $response = $Handle.Task.GetAwaiter().GetResult()
        $content = $response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
        $retry = if ($response.Headers.RetryAfter -and $response.Headers.RetryAfter.Delta) { $response.Headers.RetryAfter.Delta.Value.TotalSeconds } else { 0 }
        return @{ Status = [int]$response.StatusCode; RetryAfter = $retry; Content = $content }
    }
    catch {
        $e = $_.Exception
        while ($e.InnerException) { $e = $e.InnerException }
        return @{ Status = 0; RetryAfter = 0; Content = ''; Error = $e.Message }
    }
    finally {
        if ($response) { $response.Dispose() }
        if ($Handle.Request) { $Handle.Request.Dispose() }
    }
}

function Wait-MclUi {
    <# Waits a little while keeping the window responsive; throws when the user stopped the run. #>
    param([int]$Milliseconds = 50, [switch]$NoCancel)
    if ($script:Ui -and $script:Ui.Pump) {
        $until = [datetime]::UtcNow.AddMilliseconds($Milliseconds)
        do { & $script:Ui.Pump; Start-Sleep -Milliseconds 15 } while ([datetime]::UtcNow -lt $until)
    }
    else { Start-Sleep -Milliseconds $Milliseconds }
    if (-not $NoCancel) { Assert-MclNotCancelled }
}

function Assert-MclNotCancelled {
    # Ui.Hold: a step that must not stop half-way (a meeting re-created and not yet sent): Stop waits for its end.
    if ($script:Ui -and $script:Ui.Cancel -and -not $script:Ui.Hold) { throw [OperationCanceledException]::new('Stopped by the user.') }
}

function Get-MclGraphError {
    <# Code and message of a Graph error body. #>
    param($Body)
    $code = ''; $message = ''
    if ($Body -and $Body.PSObject.Properties['error']) {
        $code = [string]$Body.error.code
        $message = [string]$Body.error.message
    }
    return @{ Code = $code; Message = $message }
}

function ConvertFrom-MclJson {
    param([AllowEmptyString()][AllowNull()][string]$Content)
    if ([string]::IsNullOrWhiteSpace($Content)) { return $null }
    try { return $Content | ConvertFrom-Json -Depth 64 } catch { return [pscustomobject]@{ error = [pscustomobject]@{ code = 'InvalidJson'; message = ($Content.Substring(0, [Math]::Min(300, $Content.Length))) } } }
}

function Get-MclMailboxKey {
    <# The mailbox a request reads or changes (/users/<address>/...), to keep at most 4 requests at a time per mailbox. #>
    param([string]$Url)
    $m = [regex]::Match($Url, '^/?users/([^/?]+)', 'IgnoreCase')
    if ($m.Success) { return [Uri]::UnescapeDataString($m.Groups[1].Value).ToLowerInvariant() }
    return '(directory)'
}

function Invoke-MclGraph {
    <#
        One request (not batched), with the same retries as the batches. Path relative to /v1.0 or an
        absolute URL (nextLink). Returns @{ Status; Body; ErrorCode; ErrorMessage }.
    #>
    param([string]$Method = 'GET', [Parameter(Mandatory = $true)][string]$Path, $Body)

    $url = if ($Path -match '^https://') { $Path } else { "$($script:GraphRoot)/$($Path.TrimStart('/'))" }
    $json = if ($null -eq $Body) { $null } elseif ($Body -is [string]) { $Body } else { $Body | ConvertTo-Json -Depth 20 -Compress }
    $max = [int]$script:Graph.Settings.MaxRetries
    for ($attempt = 0; ; $attempt++) {
        Update-MclToken
        $handle = Start-MclGraphSend -Method $Method -Url $url -Body $json
        while (-not (Test-MclGraphSendDone $handle)) { Wait-MclUi 30 -NoCancel }
        $r = Complete-MclGraphSend $handle
        $parsed = ConvertFrom-MclJson $r.Content
        if ($r.Status -eq 401 -and $attempt -lt 1) { Update-MclToken -Force; continue }
        if ($r.Status -in 0, 429, 500, 502, 503, 504 -and $attempt -lt $max) {
            $delay = if ($r.RetryAfter -gt 0) { $r.RetryAfter } else { [Math]::Min(60, [Math]::Pow(2, $attempt + 1)) }
            Write-MclLog 'WARN' ("Graph $Method $Path -> $($r.Status)$(if ($r['Error']) { " ($($r['Error']))" }), retry in $delay s")
            Wait-MclUi ([int]($delay * 1000))
            continue
        }
        $err = Get-MclGraphError $parsed
        if ($r.Status -eq 0) { $err = @{ Code = 'NoResponse'; Message = [string]$r['Error'] } }
        return [pscustomobject]@{ Status = $r.Status; Body = $parsed; ErrorCode = $err.Code; ErrorMessage = $err.Message }
    }
}

function Get-MclGraphAll {
    <# Every item of a list (follows @odata.nextLink). Throws on an error, with the Graph message. #>
    param([Parameter(Mandatory = $true)][string]$Path)
    $items = [Collections.Generic.List[object]]::new()
    $next = $Path
    while ($next) {
        $r = Invoke-MclGraph -Path $next
        if ($r.Status -ne 200) { throw "Graph GET $Path -> $($r.Status) $($r.ErrorCode): $($r.ErrorMessage)" }
        foreach ($v in @($r.Body.value)) { if ($null -ne $v) { $items.Add($v) } }
        $next = if ($r.Body.PSObject.Properties['@odata.nextLink']) { [string]$r.Body.'@odata.nextLink' } else { $null }
    }
    return , $items.ToArray()
}

function New-MclGraphRequest {
    <# One request for Invoke-MclGraphBatch. Url relative to /v1.0 (/users/...). Headers: Prefer (immutable IDs)... #>
    param([Parameter(Mandatory = $true)][string]$Id, [Parameter(Mandatory = $true)][string]$Url, [string]$Method = 'GET', $Body, [hashtable]$Headers)
    [pscustomobject]@{ Id = $Id; Method = $Method; Url = '/' + $Url.TrimStart('/'); Body = $Body; Headers = $Headers; Mailbox = (Get-MclMailboxKey $Url.TrimStart('/')) }
}

function Invoke-MclGraphBatch {
    <#
    .SYNOPSIS
        Sends many requests through $batch and returns their results by Id.
    .PARAMETER Requests
        Objects of New-MclGraphRequest (Id unique).
    .PARAMETER FollowPages
        A list answer with @odata.nextLink is followed: Values holds every item of every page.
    .PARAMETER OnProgress
        Called after each $batch call with (requests done, requests in total).
    .OUTPUTS
        Hashtable Id -> [pscustomobject]@{ Id; Status; Body; Values; ErrorCode; ErrorMessage; Done; DoneUtc }.
        Stopped by the user: OperationCanceledException, with the results so far in Exception.Data['Results'].
    #>
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Requests,
        [switch]$FollowPages,
        [scriptblock]$OnProgress
    )

    $results = @{}
    if (-not $Requests.Count) { return $results }
    $settings = $script:Graph.Settings
    $maxRetries = [int]$settings.MaxRetries
    $concurrency = [Math]::Max(1, [int]$settings.MaxConcurrency)
    $queue = [Collections.Generic.List[object]]::new()
    foreach ($r in $Requests) {
        $queue.Add(@{ Request = $r; Attempt = 0; NotBefore = [datetime]::MinValue })
        $results[$r.Id] = [pscustomobject]@{ Id = $r.Id; Status = 0; Body = $null; Values = [Collections.Generic.List[object]]::new(); ErrorCode = ''; ErrorMessage = ''; Done = $false; DoneUtc = [datetime]::MinValue }
    }
    $inflight = [Collections.Generic.List[object]]::new()
    $busy = @{}
    $total = $Requests.Count; $done = 0; $calls = 0; $retries = 0
    $started = [datetime]::UtcNow
    $requeue = {
        param($item, [double]$Seconds)
        $item.Attempt++
        $item.NotBefore = [datetime]::UtcNow.AddSeconds($Seconds)
        $queue.Add($item)
    }
    $finish = {
        param($item, [int]$Status, $Body)
        $res = $results[$item.Request.Id]
        $res.Status = $Status
        $err = Get-MclGraphError $Body
        $res.ErrorCode = $err.Code; $res.ErrorMessage = $err.Message
        if ($Status -eq 200 -and $Body -and $Body.PSObject.Properties['value']) {
            foreach ($v in @($Body.value)) { if ($null -ne $v) { $res.Values.Add($v) } }
            $next = if ($Body.PSObject.Properties['@odata.nextLink']) { [string]$Body.'@odata.nextLink' } else { '' }
            if ($FollowPages -and $next) {
                $nextUrl = $next -replace '^https://graph\.microsoft\.com/v1\.0', ''
                $queue.Add(@{ Request = [pscustomobject]@{ Id = $item.Request.Id; Method = 'GET'; Url = $nextUrl; Body = $null; Headers = $item.Request.Headers; Mailbox = $item.Request.Mailbox }; Attempt = 0; NotBefore = [datetime]::MinValue })
                return
            }
        }
        $res.Body = $Body
        $res.Done = $true
        $res.DoneUtc = [datetime]::UtcNow
        $script:MclBatchDone++
    }

    $script:MclBatchDone = 0
    $cancelled = $false
    while ($queue.Count -or $inflight.Count) {
        if (-not $cancelled -and $script:Ui -and $script:Ui.Cancel -and -not $script:Ui.Hold) { $cancelled = $true }
        if ($cancelled -and -not $inflight.Count) {
            # The requests already done are given to the caller (Data.Results): a change done must be reported.
            $stop = [OperationCanceledException]::new('Stopped by the user.')
            $stop.Data['Results'] = $results
            throw $stop
        }

        # ---- schedule new $batch calls --------------------------------------------------------
        while (-not $cancelled -and $inflight.Count -lt $concurrency -and $queue.Count) {
            $now = [datetime]::UtcNow
            $picked = [Collections.Generic.List[object]]::new()
            $inBatch = @{}
            foreach ($item in $queue) {
                if ($picked.Count -ge $script:BatchSize) { break }
                if ($item.NotBefore -gt $now) { continue }
                $key = $item.Request.Mailbox
                $used = [int]$busy[$key] + [int]$inBatch[$key]
                # The limit of 4 applies to a mailbox; directory requests (users, groups, places) are not limited.
                if ($used -ge $script:MailboxConcurrency -and -not $key.StartsWith('(')) { continue }
                $inBatch[$key] = [int]$inBatch[$key] + 1
                $picked.Add($item)
            }
            if (-not $picked.Count) { break }
            foreach ($item in $picked) { [void]$queue.Remove($item); $busy[$item.Request.Mailbox] = [int]$busy[$item.Request.Mailbox] + 1 }
            $subs = for ($i = 0; $i -lt $picked.Count; $i++) {
                $req = $picked[$i].Request
                $sub = [ordered]@{ id = [string]$i; method = $req.Method; url = $req.Url }
                $headers = @{}
                if ($req.PSObject.Properties['Headers'] -and $req.Headers) { foreach ($h in $req.Headers.Keys) { $headers[$h] = $req.Headers[$h] } }
                if ($null -ne $req.Body) { $headers['Content-Type'] = 'application/json'; $sub.body = $req.Body }
                if ($headers.Count) { $sub.headers = $headers }
                $sub
            }
            Update-MclToken
            $payload = @{ requests = @($subs) } | ConvertTo-Json -Depth 20 -Compress
            $inflight.Add(@{ Handle = (Start-MclGraphSend -Method 'POST' -Url "$($script:GraphRoot)/`$batch" -Body $payload); Items = $picked })
            $calls++
        }

        if (-not $inflight.Count) {
            # Only requests waiting for their Retry-After.
            $wait = ($queue | ForEach-Object { $_.NotBefore } | Measure-Object -Minimum).Minimum
            $ms = [int][Math]::Max(50, [Math]::Min(5000, ($wait - [datetime]::UtcNow).TotalMilliseconds))
            Wait-MclUi $ms -NoCancel
            continue
        }

        # ---- wait for a $batch call to finish ------------------------------------------------
        $finished = @($inflight | Where-Object { Test-MclGraphSendDone $_.Handle })
        if (-not $finished.Count) { Wait-MclUi 25 -NoCancel; continue }
        foreach ($call in $finished) {
            [void]$inflight.Remove($call)
            foreach ($item in $call.Items) { $busy[$item.Request.Mailbox] = [int]$busy[$item.Request.Mailbox] - 1 }
            $response = Complete-MclGraphSend $call.Handle
            if ($response.Status -eq 200) {
                $parsed = ConvertFrom-MclJson $response.Content
                $byId = @{}
                foreach ($sub in @(Get-MclProperty $parsed 'responses')) { if ($sub) { $byId[[string]$sub.id] = $sub } }
                for ($i = 0; $i -lt $call.Items.Count; $i++) {
                    $item = $call.Items[$i]
                    $sub = $byId[[string]$i]
                    if (-not $sub) { & $requeue $item 2; $retries++; continue }
                    $status = [int]$sub.status
                    $body = if ($sub.PSObject.Properties['body']) { $sub.body } else { $null }
                    if ($status -in 429, 500, 502, 503, 504 -and $item.Attempt -lt $maxRetries) {
                        $after = 0.0
                        if ($sub.PSObject.Properties['headers'] -and $sub.headers -and $sub.headers.PSObject.Properties['Retry-After']) { [void][double]::TryParse([string]$sub.headers.'Retry-After', [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$after) }
                        if ($after -le 0) { $after = [Math]::Min(60, [Math]::Pow(2, $item.Attempt + 1)) }
                        & $requeue $item $after; $retries++
                        continue
                    }
                    if ($status -eq 401 -and $item.Attempt -lt 1) { Update-MclToken -Force; & $requeue $item 0; $retries++; continue }
                    & $finish $item $status $body
                }
            }
            elseif ($response.Status -eq 401 -and @($call.Items | Where-Object { $_.Attempt -lt 1 }).Count) {
                Update-MclToken -Force
                foreach ($item in $call.Items) { & $requeue $item 0 }
                $retries += $call.Items.Count
            }
            elseif ($response.Status -in 0, 429, 500, 502, 503, 504 -and @($call.Items | Where-Object { $_.Attempt -lt $maxRetries }).Count) {
                $after = if ($response.RetryAfter -gt 0) { $response.RetryAfter } else { [Math]::Min(60, [Math]::Pow(2, $call.Items[0].Attempt + 1)) }
                Write-MclLog 'WARN' ("Graph `$batch -> $($response.Status)$(if ($response['Error']) { " ($($response['Error']))" }), $($call.Items.Count) request(s) retried in $after s")
                foreach ($item in $call.Items) { & $requeue $item $after }
                $retries += $call.Items.Count
            }
            else {
                $parsed = ConvertFrom-MclJson $response.Content
                if ($response.Status -eq 0) { $parsed = [pscustomobject]@{ error = [pscustomobject]@{ code = 'NoResponse'; message = [string]$response['Error'] } } }
                foreach ($item in $call.Items) { & $finish $item $response.Status $parsed }
            }
        }
        if ($OnProgress) { & $OnProgress $script:MclBatchDone $total }
    }
    Write-MclLog 'INFO' ("Graph: {0} request(s) in {1} `$batch call(s), {2} retried, {3}" -f $total, $calls, $retries, (Format-MclDuration ([datetime]::UtcNow - $started).TotalSeconds))
    foreach ($res in $results.Values) { if ($res.Status -eq 200 -and $null -eq $res.Body) { $res.Body = [pscustomobject]@{ value = $res.Values.ToArray() } } }
    return $results
}

# SIG # Begin signature block
# MIImIgYJKoZIhvcNAQcCoIImEzCCJg8CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCEFcpWWcDSSTj9
# 5TpHnpz/ISTi9bmqq7FKSZ1NiNgSaKCCDW8wggZyMIIEWqADAgECAghkM1HTxzif
# CDANBgkqhkiG9w0BAQsFADB8MQswCQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMx
# EDAOBgNVBAcMB0hvdXN0b24xGDAWBgNVBAoMD1NTTCBDb3Jwb3JhdGlvbjExMC8G
# A1UEAwwoU1NMLmNvbSBSb290IENlcnRpZmljYXRpb24gQXV0aG9yaXR5IFJTQTAe
# Fw0xNjA2MjQyMDQ0MzBaFw0zMTA2MjQyMDQ0MzBaMHgxCzAJBgNVBAYTAlVTMQ4w
# DAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91c3RvbjERMA8GA1UECgwIU1NMIENv
# cnAxNDAyBgNVBAMMK1NTTC5jb20gQ29kZSBTaWduaW5nIEludGVybWVkaWF0ZSBD
# QSBSU0EgUjEwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQCfgxNzqrDG
# bSHL24t6h3TQcdyOl3Ka5LuINLTdgAPGL0WkdJq/Hg9Q6p5tePOf+lEmqT2d0bKU
# Vz77OYkbkStW72fL5gvjDjmMxjX0jD3dJekBrBdCfVgWQNz51ShEHZVkMGE6ZPKX
# 13NMfXsjAm3zdetVPW+qLcSvvnSsXf5qtvzqXHnpD0OctVIFD+8+sbGP0EmtpuNC
# GVQ/8y8Ooct8/hP5IznaJRy4PgBKOm8yMDdkHseudQfYVdIYyQ6KvKNc8HwKp4WB
# wg6vj5lc02AlvINaaRwlE81y9eucgJvcLGfE3ckJmNVz68Qho+Uyjj4vUpjGYDdk
# jLJvSlRyGMwnh/rNdaJjIUy1PWT9K6abVa8mTGC0uVz+q0O9rdATZlAfC9KJpv/X
# gAbxwxECMzNhF/dWH44vO2jnFfF3VkopngPawismYTJboFblSSmNNqf1x1KiVgMg
# Lzh4gL32Bq5BNMuURb2bx4kYHwu6/6muakCZE93vUN8BuvIE1tAx3zQ4XldbyDge
# VtSsSKbt//m4wTvtwiS+RGCnd83VPZhZtEPqqmB9zcLlL/Hr9dQg1Zc0bl0EawUR
# 0tOSjAknRO1PNTFGfnQZBWLsiePqI3CY5NEv1IoTGEaTZeVYc9NMPSd6Ij/D+KNV
# t/nmh4LsRR7Fbjp8sU65q2j3m2PVkUG8qQIDAQABo4H7MIH4MA8GA1UdEwEB/wQF
# MAMBAf8wHwYDVR0jBBgwFoAU3QQJB6L1en1SUxKSle44gCUNplkwMAYIKwYBBQUH
# AQEEJDAiMCAGCCsGAQUFBzABhhRodHRwOi8vb2NzcHMuc3NsLmNvbTARBgNVHSAE
# CjAIMAYGBFUdIAAwEwYDVR0lBAwwCgYIKwYBBQUHAwMwOwYDVR0fBDQwMjAwoC6g
# LIYqaHR0cDovL2NybHMuc3NsLmNvbS9zc2wuY29tLXJzYS1Sb290Q0EuY3JsMB0G
# A1UdDgQWBBRUwv4QlQCTzWr158DX2bJLuI8M4zAOBgNVHQ8BAf8EBAMCAYYwDQYJ
# KoZIhvcNAQELBQADggIBAPUPJodwr5miyvXWyfCNZj05gtOII9iCv49UhCe204MH
# 154niU2EjlTRIO5gQ9tXQjzHsJX2vszqoz2OTwbGK1mGf+tzG8rlQCbgPW/M9r1x
# xs19DiBAOdYF0q+UCL9/wlG3K7V7gyHwY9rlnOFpLnUdTsthHvWlM98CnRXZ7WmT
# V7pGRS6AvGW+5xI+3kf/kJwQrfZWsqTU+tb8LryXIbN2g9KR+gZQ0bGAKID+260P
# Z+34fdzZcFt6umi1s0pmF4/n8OdX3Wn+vF7h1YyfE7uVmhX7eSuF1W0+Z0duGwdc
# +1RFDxYRLhHDsLy1bhwzV5Qe/kI0Ro4xUE7bM1eV+jjk5hLbq1guRbfZIsr0WkdJ
# LCjoT4xCPGRo6eZDrBmRqccTgl/8cQo3t51Qezxd96JSgjXktefTCm9r/o35pNfV
# HUvnfWII+NnXrJlJ27WEQRQu9i5gl1NLmv7xiHp0up516eDap8nMLDt7TAp4z5T3
# NmC2gzyKVMtODWgqlBF1JhTqIDfM63kXdlV4cW3iSTgzN9vkbFnHI2LmvM4uVEv9
# XgMqyN0eS3FE0HU+MWJliymm7STheh2ENH+kF3y0rH0/NVjLw78a3Z9UVm1F5VPz
# iIorMaPKPlDRADTsJwjDZ8Zc6Gi/zy4WZbg8Zv87spWrmo2dzJTw7XhQf+xkR6Od
# MIIG9TCCBN2gAwIBAgIQBof4ljXihJpmCqqt2mhAWzANBgkqhkiG9w0BAQsFADB4
# MQswCQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNVBAcMB0hvdXN0b24x
# ETAPBgNVBAoMCFNTTCBDb3JwMTQwMgYDVQQDDCtTU0wuY29tIENvZGUgU2lnbmlu
# ZyBJbnRlcm1lZGlhdGUgQ0EgUlNBIFIxMB4XDTI2MTAxMzAyNDgwMloXDTI3MTAx
# MzAyNDgwMlowXDELMAkGA1UEBhMCRlIxGzAZBgNVBAcMEkNoYW50ZWxvdXAtRW4t
# QnJpZTEXMBUGA1UECgwOTmljb2xhcyBGYWJlcnQxFzAVBgNVBAMMDk5pY29sYXMg
# RmFiZXJ0MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAsg6CT0mFQ4jB
# V2plIS4uWFSoVqRHd82Y7XYZ4CUTe/7nPBRUlXXdKTZMMHzGgiG8dEXNE8K1Xe90
# Dgl/FnorzAFlrNWSwf6t27UAUw9xkqtRDVVpVgpixpWkAwjcQeLG+YMqBkGbFi35
# TGvS23lPFWvLp4DxARrv15MtZ/NjNTCAAGUlMP2hFVBknEmrnYjEGyeJ93sl20RS
# WXQQLNqlenBdkGCrQW4XjHtRjJAEqAC68hN6nWevopSQ8CLUV7A1RFfHFarDrzrL
# zr0/NjIdSyIfM9hShln31NDrktB6Ups2XIyrtPDNB7fjkW4jpLlOqQGfYKsU9ErE
# 4SnFsh0yOOjitegyh1RvZ9wH++3DF4CmJfTli5l/gEIm8xCCryxI+K5EogW3wu1y
# u8NT2hcvHjdh7OpRyPyFAuLZ8MfX7kPAJpnuCeOuktzY/20eVD8bDglj+4zgqzsn
# cR/t6T/7w1/HxpRyrVoKvEyRx+mphYh5jMdEa0BjxOL0vAJpdGZSfXGmPm/yPqRn
# 5STTkMtzkVQXPbcvgKM0luzFcQKPHk3qdHLEXa8gf93ktYDbLE5Wd2UHrXz3SZxN
# URHPdIXCr1qgrqQve7041kMpzrkDI5DnG2lvRG4108gp02bHG0muKyZlDdCOra7h
# Zonu2nO79sFallruzzNi0o994PoWksUCAwEAAaOCAZUwggGRMAwGA1UdEwEB/wQC
# MAAwHwYDVR0jBBgwFoAUVML+EJUAk81q9efA19myS7iPDOMwegYIKwYBBQUHAQEE
# bjBsMEgGCCsGAQUFBzAChjxodHRwOi8vY2VydC5zc2wuY29tL1NTTGNvbS1TdWJD
# QS1Db2RlU2lnbmluZy1SU0EtNDA5Ni1SMS5jZXIwIAYIKwYBBQUHMAGGFGh0dHA6
# Ly9vY3Nwcy5zc2wuY29tMFEGA1UdIARKMEgwCAYGZ4EMAQQBMDwGDCsGAQQBgqkw
# AQMDATAsMCoGCCsGAQUFBwIBFh5odHRwczovL3d3dy5zc2wuY29tL3JlcG9zaXRv
# cnkwEwYDVR0lBAwwCgYIKwYBBQUHAwMwTQYDVR0fBEYwRDBCoECgPoY8aHR0cDov
# L2NybHMuc3NsLmNvbS9TU0xjb20tU3ViQ0EtQ29kZVNpZ25pbmctUlNBLTQwOTYt
# UjEuY3JsMB0GA1UdDgQWBBRoUx3630l9wA1fUpkKkw2GUIxl/jAOBgNVHQ8BAf8E
# BAMCB4AwDQYJKoZIhvcNAQELBQADggIBAE3oj6cNuwlDToi1BwjgsA5djfRN4DIi
# BYDFZIRYaIRe/rEdnItz3aqQQdFuwXnJA/MQEmi2WAGek8GeZF+R41s8FbcDxBUE
# BoSt98TuBE+0xgxQWDiPhhXaX1Aw9V0GF4ZnlfCSw1eu2KiASaKn7Bo0PLbEAXff
# vftRXof/kyqtw8iID6I4Rhrf5DI3WVZQodakUUmWnbnziWvEZOWbNOfzwP+7BwEt
# 7jQANfbmEfb4p75rfy5dQVwniTDSvdo6xkrbe1Cr93GS7cJta18IG5CoIVUmdK+s
# W38B1sLuy8Mi0MstMXV3004yaRuzLdNaAM0f9WAFerxemiQC0kdmA44/lWV3lVkM
# WJSwkHBkkSg8Eh0hxcGzTH2xmtDhLJHNNLRwkI9u2Ur+HbLCt50oBgSV8BcCs/1i
# qi884iC/yuX+HIv9YgTUNZ8c4lLaq4VuliJpXLu2buHIsTkO0v0+5oCM91zjTVZm
# ULeH38N0Vr83psiJVuBSxQyt5nRVrpYz7FACq5gtXELkyViq/rGOlJKaRrgQ6WWq
# eoIXqKhVFAdP/jAEm2M8ezoVYsUgPyKY2OcuecV9coulAcxUIUydx6qYiOfsGj5u
# Dy7l4aYjvB771wkUMzUQwxzcDYXJtMMTaWRe6n9EemR67jc73u5CJyFg3PfA5WR1
# wZislcdCZuS6MYIYCTCCGAUCAQEwgYwweDELMAkGA1UEBhMCVVMxDjAMBgNVBAgM
# BVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMREwDwYDVQQKDAhTU0wgQ29ycDE0MDIG
# A1UEAwwrU1NMLmNvbSBDb2RlIFNpZ25pbmcgSW50ZXJtZWRpYXRlIENBIFJTQSBS
# MQIQBof4ljXihJpmCqqt2mhAWzANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3
# AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisG
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDs4Jrl1lL8
# DwYcp52x4TiikK/et4lqF/kHkM9axTksADANBgkqhkiG9w0BAQEFAASCAgAfsUdJ
# m6nVOWX2BYrNEvesPytlWN96NqrCApia7nrCmh9+gfJ1X+InE9MbqmcafZnxWDHO
# DAbWhJzyumhdO8hcQ9VKmyHP0Vuawhdyz7CzmJXggcqM0N8wIrsuWxJFBg7km1f+
# wblpY4djaDmpcwy1BqcHu/sp6TFUGryCrpuCxSDRhFz9vBlRIkbbOGd0j6jzHmuE
# 7qVLGYE2cmKve+lATltVcgeRbjRdMmw3j91lJw75oYzu9PPUGZ63RIoCMDLrrGoF
# Pih/Lsf6NYa/SdI2L7oC5UoOSE0C5vAnycB0g1rDgY8w47gGrS5weDpL3NiEipQA
# 56OQqKD8dhtMVIA2z3CyWLcV0Dw+3DGOfP7mRIujgT5psUtr26OxNSEbO7n2NZvY
# hZkmLDEJkWulaaBTX6xb5/MveoRpzIG5b3YHz7F3io8QxVuehyYpxxoykTkTMRpI
# YqTvY1SlvinBn3ErWlngUZx3W7gnSm2Lr/n5OBXazIi9qckSCJj5CVUuXouRiN0O
# rNEobRSrDAQ4YU7okrFrGEaHWyAwwKlLYq9CAN8x4qAlAov3jqylEcpCqxK2KQY8
# O1twehQGpMSrXNg6D4HoHzqG2zaUHo4lcXKcyhn15L7ODfdiH/6r7Xqm6Wl1D2AZ
# D0elXV2EdmnlTBdrDAZNqlNjaE4wJfUciSScwqGCFMYwghTCBgorBgEEAYI3AwMB
# MYIUsjCCFK4GCSqGSIb3DQEHAqCCFJ8wghSbAgEDMQ0wCwYJYIZIAWUDBAIBMIGC
# BgsqhkiG9w0BCRABBKBzBHEwbwIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEIM4HzJRL+sIxaKQ3F1vdOnIcDxeXYYNiXnehgjOvGpGqAghfOinPahEn
# zBgPMjAyNjExMDYwODQyMjNaMAMCAQECCQC3cHI75+p4qKCCEacwggTzMIIC26AD
# AgECAhBUQ/g3vVruGAGEcrhc2o7jMA0GCSqGSIb3DQEBCwUAMHIxCzAJBgNVBAYT
# AlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91c3RvbjERMA8GA1UECgwI
# U1NMIENvcnAxLjAsBgNVBAMMJVNTTC5jb20gVGltZXN0YW1waW5nIElzc3Vpbmcg
# UlNBIENBIDIwHhcNMjYwNjEyMTcyMjQ2WhcNMzYwNjA5MTcyMjQ1WjBuMQswCQYD
# VQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNVBAcMB0hvdXN0b24xETAPBgNV
# BAoMCFNTTCBDb3JwMSowKAYDVQQDDCFTU0wuY29tIFRpbWVzdGFtcGluZyBVbml0
# IDIwMjYgRTEwWTATBgcqhkjOPQIBBggqhkjOPQMBBwNCAAToJYu4ZhMuSJI48ws2
# yQ4G2FewiFFIDjp945DC9ItwNo2Hq/FTMEoTHUp0UFEWzs4eucf8rRUDoxsNAgWl
# SDQ9o4IBUjCCAU4wHwYDVR0jBBgwFoAUcJplwUOAo/zpsWSJbvEwnAkhyXwwTQYI
# KwYBBQUHAQEEQTA/MD0GCCsGAQUFBzAChjFodHRwOi8vY2VydC5zc2wuY29tL1NT
# TC5jb20tdGltZVN0YW1waW5nLUktUjIuY2VyMFEGA1UdIARKMEgwPAYMKwYBBAGC
# qTABAwYBMCwwKgYIKwYBBQUHAgEWHmh0dHBzOi8vd3d3LnNzbC5jb20vcmVwb3Np
# dG9yeTAIBgZngQwBBAIwFgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwQgYDVR0fBDsw
# OTA3oDWgM4YxaHR0cDovL2NybHMuc3NsLmNvbS9TU0wuY29tLXRpbWVTdGFtcGlu
# Zy1JLVIyLmNybDAdBgNVHQ4EFgQUJ6wDVrooaroXO4M9QH95qCTswzwwDgYDVR0P
# AQH/BAQDAgeAMA0GCSqGSIb3DQEBCwUAA4ICAQBPdm7O/AOewhlNcGf6GOa4FeDM
# HXZ9tjFP25h7pm+G+ltlYpzINCBpwROQ/DyTrJ6qdgJjIVnTreUw8SoCEAPacwbr
# XhmHs+0db0Nap+yCK/Fs9gF4363tTGCRhPd7GltYx08SoG1ZXIc68twq8zGAlZ8m
# MA3BWTBUN+TL+Ptx/bz6KKY3w8pztL+7z8t6dFGdXZbZh+dxGeTP0J+zmdEPNpmC
# 6IeuunYbHCCc1BIfgYGMc6Xc/MuydVsahkqLhkSOBe6mziYYY4rv4YHYcCVCf45L
# 3jhweqWQfSASW3ZZ7OgI1GkQumH2prG4GeNISpdZZ+ATrIYHA7S7Eo2qlZVMeYDC
# vpK8E+h/G5tci4Sf2K9HgIs64DFRKsnjVkGD+qobSgmMXoaREpvGrtzJZKQRlgmG
# YoKDYeDl9cdqA7GLBuafkk9nW+FYKtuTe/91PHaC9Mrn56sLkf3+tNyWNQjvvrJ3
# at+GN+Sm/ICxu1ezgPueufDyL3Civn5921FIkFXGrMubKQV+2gvRSnZD763IoJR2
# uJt1yesHgcpLez63cCN1xsDToMIvwaJ9CL18vTJEFE2UDijoE/5KvvnkxoI2NRFh
# kO8PLEkcEFPtyqNkjCyJHyz7rzrSegVsEG12B7YjDsI8L/VDfUjbuFWiVLsHTJqj
# e/xfYy1bLeaxmVdNwTCCBd0wggPFoAMCAQICCHssm9MWgDKZMA0GCSqGSIb3DQEB
# CwUAMHwxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91
# c3RvbjEYMBYGA1UECgwPU1NMIENvcnBvcmF0aW9uMTEwLwYDVQQDDChTU0wuY29t
# IFJvb3QgQ2VydGlmaWNhdGlvbiBBdXRob3JpdHkgUlNBMB4XDTE2MDIxMjE3Mzkz
# OVoXDTQxMDIxMjE3MzkzOVowfDELMAkGA1UEBhMCVVMxDjAMBgNVBAgMBVRleGFz
# MRAwDgYDVQQHDAdIb3VzdG9uMRgwFgYDVQQKDA9TU0wgQ29ycG9yYXRpb24xMTAv
# BgNVBAMMKFNTTC5jb20gUm9vdCBDZXJ0aWZpY2F0aW9uIEF1dGhvcml0eSBSU0Ew
# ggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQD5D92jK33L0Cr+7GeFpucu
# G7p34eP1r6Ts+kpdkcRXR2sYd2t28v2T5D0PwhaeC2bDVpSeF4OFzlbv8hb9AGL1
# IglU6GUXTkG54E9Gl6obyLhuYl5psV/bKgJ+/GzK80HY7dDo/D9hSO2wAxQdEA5L
# GeC7TuyGZf82815nAgudhlVh/Xo47f7iGQC3b6FQYnV0PKD6yCWStG56Isf4HqHj
# st2RMasrHQT/pUoEN+mFpDMr/eLWVTR8GaRKaMeyqNO3yqGTiOvBl7yM+R3ZIoQk
# dMcEPWqpKZPM67hb4f5fJao0WMjBI1SdG5gRwzicfj2GbKUPQIZ8AvRcAk8oy65x
# nw86yDP+ESU16vy6xWA92XwY1bKp03V4A3IiyjrDH+8s5S6p+p4stlFG/a8D1upg
# aOqFFjZrhekewLPdxCTcgCqBQW2UPsjgyYFBAJ5ev3/FCJiiGCxCQLP5bzgnS06A
# 9D2BR+CIfOoczrV1XFEuHCt/GnIo5wC10XTG1+SfrQeTtlM1Nfw35MP2XRa+IXPe
# kgr4oGNqvJaSaj74vGVVm971DYkmBPwlGqYlacvCbcp84llfl6zr7y7IvNcbWTwr
# zPIZyJNrJ2MZz/zpJvjKcZt/k/40Z4ROmev8s3gJM3C6ZqZ27Rtz6xqlDcQiEyCU
# VgpOLGxOsf3PnAm6ojPthwIDAQABo2MwYTAdBgNVHQ4EFgQU3QQJB6L1en1SUxKS
# le44gCUNplkwDwYDVR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBTdBAkHovV6fVJT
# EpKV7jiAJQ2mWTAOBgNVHQ8BAf8EBAMCAYYwDQYJKoZIhvcNAQELBQADggIBACAY
# EZQp+yadHB4ecGHxlXKTcSStaJNYjjKvG7NwA/wlK3SFkD14avS5i6WXO7UYkbse
# p/lAW5H5VZmvHhHQXB2nZuOxlAcMMjmm6huwedgdnHBE44rdxPmVH4o4Qz8BhaVH
# pz1GsrzlImj3e5zYLD4KIcgtM6y/xYGZMXTBdXHFvrHwI0X0nWv8GWOdo7wExhgL
# JbtTiQ+zgFDeRe5Ef6uUeGSY0/Yo3YfYcGV0+w65E+unD2GpMpbM3rvtY0wYu6lA
# 96BUbiCIcXUY6nq0NHLgIyd3XLaQ6oYlQKvvMw/Ln4K+oiD79rUtGubChbF0D/vI
# ZQKkUgFH3Ukiwb/Y62usft7sYzMVtyMIj8YPjUFa3Y7FuY/lRT9427rSG0Cx/nFN
# P+CBorpetOwV4JPdCB9+4VWZCyHek54K++ajSb02MP7nd7KgdZe1LYGIF2Ug99qQ
# AJ/JUswyyjV89T0P2CvX9SZsyQY0lhbqcFkaMnl5C7aIfw9SSD2/bNiiRC7RTrdy
# WNOJE5X+RKv414sbbpy8LKBb1WoAr1834dX6EAuYnIbnJo/O8OxuilcLgONOssCg
# Y2GQulVoN3RqtpLbn6GGIrZlJw7stp9CYORnwrXaQQvE04thG7z6H5Er10QHXrop
# rNnF6e9TSFrrgPEoWCHNsAZV+yc/U5BwqQQeVye5MIIGyzCCBLOgAwIBAgIQUp27
# wriJh/WDgOCw66IGNjANBgkqhkiG9w0BAQsFADB8MQswCQYDVQQGEwJVUzEOMAwG
# A1UECAwFVGV4YXMxEDAOBgNVBAcMB0hvdXN0b24xGDAWBgNVBAoMD1NTTCBDb3Jw
# b3JhdGlvbjExMC8GA1UEAwwoU1NMLmNvbSBSb290IENlcnRpZmljYXRpb24gQXV0
# aG9yaXR5IFJTQTAeFw0yNjA2MTIxNjQyMTVaFw00MTAyMTIxNzM5MzlaMHIxCzAJ
# BgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91c3RvbjERMA8G
# A1UECgwIU1NMIENvcnAxLjAsBgNVBAMMJVNTTC5jb20gVGltZXN0YW1waW5nIElz
# c3VpbmcgUlNBIENBIDIwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQCf
# L6vmc02QvDFGmzYHeVpQY1T3Lu9aInnTXfJ/SJfPrltB3Br1uOPyX9KR5eocap9K
# HYWhp77zQTyYMR/3egjCOzLYJfZxeYrGq0XwckZgHh/rJGEoQPnTCRs7KwvTYhd/
# 78RmdrjiN9+GFXPg5T+TnbFD8+ciD+4WFzZaT5Jz8IeoCv6VhWQ/PsvLlGoH1/VE
# oav0kLftbUuBV0cLnR00h2xrdBc+GW1hm5tHeekTgFt+LSOhRbmfk7RrLB43jlH9
# hUvEDRgaWnU4dV9WeyperxuGfDUtziSMnpc6BeU6XD9rzN7ntO9t25mXd6F7BOa8
# 7QCTL+eo5YNL0g8/F3F0bq+msFtK4eXE0A0EBsdOIWCR9mZ/R1liVP8CdkigXECr
# Qcgs+iV/g6BaARPvpKJZKWbZdhEhfBWOl6obSmW9fBEhAHNKIQmPtB5AWcK8fHiB
# Awg283f7Ib/WvGERtqMa97S1M/eVfg884dUATRAmnhnM8jVhgmSFVGuJVXFVFTgW
# yDgE/A/oj5guYdvW1nhEvA4U0uH+AlgOA+UKEKjE6iAXFeDA9lmsPIRFgHRp457F
# bQ4ppw3hFG71yoXYoq78BvkAvt5gMtUZQ+sQTJhsgwClUqql5tzT+MNCeIL7euRq
# 5ECOBVzybWKn8JCAbUPDhn26VVEBj1VHplrbouxNtQIDAQABo4IBUTCCAU0wEgYD
# VR0TAQH/BAgwBgEB/wIBADAfBgNVHSMEGDAWgBTdBAkHovV6fVJTEpKV7jiAJQ2m
# WTBNBggrBgEFBQcBAQRBMD8wPQYIKwYBBQUHMAKGMWh0dHA6Ly9jZXJ0LnNzbC5j
# b20vU1NMY29tLVJvb3RDQS1SU0EtNDA5Ni1SMS5jZXIwPwYDVR0gBDgwNjA0BgRV
# HSAAMCwwKgYIKwYBBQUHAgEWHmh0dHBzOi8vd3d3LnNzbC5jb20vcmVwb3NpdG9y
# eTATBgNVHSUEDDAKBggrBgEFBQcDCDBCBgNVHR8EOzA5MDegNaAzhjFodHRwOi8v
# Y3Jscy5zc2wuY29tL1NTTGNvbS1Sb290Q0EtUlNBLTQwOTYtUjEuY3JsMB0GA1Ud
# DgQWBBRwmmXBQ4Cj/OmxZIlu8TCcCSHJfDAOBgNVHQ8BAf8EBAMCAYYwDQYJKoZI
# hvcNAQELBQADggIBAGUUW80x6YDqIGIFijnogpCNaFa5Bpgq6jiql/Qok0REC3WD
# P4ewDZIvfED9X4QDxJE80/1K9fSKz7mYOymli+krJgapg/fpgVbSZsXUXRm1fG6Y
# JRQ79lXmYE0oN/lDg9jtBvR5YB6lRzSs0mv8y+flmJd8k/ENvbGZyDxpEPrnb+ev
# 46uTZQE/QTjeIl0jgsXtWzb1HdIZEthk7EzXkMjhB5uwpGXQoOV/g6ZgYIpzu6ZP
# GJuRm03ULUpaCd/2Np1T8nCfJJpgQGDHr+/4VULpIY43OJjtNy8j0xX6pjEdQ+U6
# boBpA/Z1LhuNZeebzqV0JtSTG04AK6ZJ9P3vTvsk2JVrs9nUgGR3/Y8TUALyCztT
# 79iTUE31dE8iLXmT+4DXEX2h557koaPKIYqSohPtbxCSBRPnN4KKcUqRwpp9kxxE
# 5JVile2a1CoYZHaDi49ionQrWGazqAlese3/finwz0+tXyg1Z7RmXUSsMy1ed1SD
# eqEkvikDdlSWjvDmUzmaX4DuKOv/caTa/hBGgN4TZKGZNd0eC2qjwv99PXX2ZhwR
# PJwTY1gZFewGXc3SvMOwdGvw/dOumuqFGtfTiJLQ/t5D6UPGBLfKEafc7NYaEO30
# Z/bHWp7T+p2EUBU6RfHMxnPhRq3rGwB/NfT7+bwYQpbk6PD8dj6Qu9XDNI87MYIC
# VTCCAlECAQEwgYYwcjELMAkGA1UEBhMCVVMxDjAMBgNVBAgMBVRleGFzMRAwDgYD
# VQQHDAdIb3VzdG9uMREwDwYDVQQKDAhTU0wgQ29ycDEuMCwGA1UEAwwlU1NMLmNv
# bSBUaW1lc3RhbXBpbmcgSXNzdWluZyBSU0EgQ0EgMgIQVEP4N71a7hgBhHK4XNqO
# 4zALBglghkgBZQMEAgGgggFgMBoGCSqGSIb3DQEJAzENBgsqhkiG9w0BCRABBDAc
# BgkqhkiG9w0BCQUxDxcNMjYxMTA2MDg0MjIzWjAoBgkqhkiG9w0BCTQxGzAZMAsG
# CWCGSAFlAwQCAaEKBggqhkjOPQQDAjAvBgkqhkiG9w0BCQQxIgQg2P6qhpnfe06F
# D3sTj2JE911qSh+HO4ONdRrVUrF+wjAwgcgGCyqGSIb3DQEJEAIvMYG4MIG1MIGy
# MIGvBCBXT+oXqvpefGormwjJmLFzXkULOHPKROaiTKCTDmPzsDCBijB2pHQwcjEL
# MAkGA1UEBhMCVVMxDjAMBgNVBAgMBVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMREw
# DwYDVQQKDAhTU0wgQ29ycDEuMCwGA1UEAwwlU1NMLmNvbSBUaW1lc3RhbXBpbmcg
# SXNzdWluZyBSU0EgQ0EgMgIQVEP4N71a7hgBhHK4XNqO4zAKBggqhkjOPQQDAgRG
# MEQCIFzWcU8kY2H0zbsck2wZQW573Yu/BYxdSfVKRDWRcgJmAiAYpaGrKYQzyMZR
# q/BjtZSXVW5ebkKWqhLLmB+4IFAevg==
# SIG # End signature block
