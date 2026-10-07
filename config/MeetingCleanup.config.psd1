#
#  Meeting Cleanup - configuration file
#  --------------------------------------------------------------------------
#  Author  : Nicolas Fabert
#  Version : 1.2.2
#
#  Read by Invoke-MeetingCleanup.ps1 and by the window (-Gui). It is a PowerShell data file: text between
#  quotes, $true / $false, numbers, @( ) for lists and @{ } for groups of settings. Lines starting with #
#  are comments. Relative paths (.\reports, .\logs) are relative to the tool folder.
#  Every value is checked at start; all the problems are listed at once.
#
#  No secret here: the certificate stays in the Windows certificate store; a client secret is read from
#  an environment variable or typed at each run, and is never written.
#
@{
    # ---------------------------------------------------------------------
    # Tenant. Safety check: the tool stops if the token belongs to another tenant.
    # ---------------------------------------------------------------------
    Tenant = @{
        TenantId     = ''      # Microsoft Entra tenant ID (GUID) or its domain contoso.onmicrosoft.com
        Organization = ''      # optional, shown in the console and the report (contoso.onmicrosoft.com)
    }

    # ---------------------------------------------------------------------
    # Application registered in Microsoft Entra (developer guide, chapter 5). Application permissions of Microsoft
    # Graph, with admin consent:
    #   Calendars.ReadWrite   required (Calendars.Read is enough for a report without action)
    #   User.Read.All         addresses of the organizer, list of every mailbox
    #   Place.Read.All        list of the room mailboxes
    #   GroupMember.Read.All  members of a group invited to a meeting
    #   Certificate  : recommended. The certificate with its private key in Cert:\CurrentUser\My or
    #                  Cert:\LocalMachine\My of the account that runs the tool.
    #   ClientSecret : the secret is read from the environment variable ClientSecretVariable, or typed.
    # ---------------------------------------------------------------------
    Authentication = @{
        Mode                  = 'Certificate'          # Certificate | ClientSecret
        AppId                 = ''                     # application (client) ID
        CertificateThumbprint = ''                     # Certificate: thumbprint (40 hexadecimal characters)
        ClientSecretVariable  = 'MCL_CLIENT_SECRET'    # ClientSecret: name of the environment variable
    }

    # ---------------------------------------------------------------------
    # Where and when the meetings are searched (defaults of -SearchIn, -Start, -End).
    #   SearchIn: one or more of
    #     'Organizer'     the organizer's calendar, when the mailbox still exists
    #     'Rooms'         every room mailbox (places API) plus Rooms and RoomFile below
    #     'Mailboxes'     the mailboxes of -Mailbox or of MailboxFile (one address per line, or a CSV file)
    #     'AllMailboxes'  every mailbox of the tenant: complete but long (a deleted organizer without room)
    #   Each meeting found is then looked up in the calendar of every internal attendee, room and member of
    #   an invited group, wherever it was found.
    #   PastDays / FutureDays: the default period, from today minus PastDays to today plus FutureDays.
    #   A series is kept when one of its occurrences falls in the period, and handled as a whole.
    # ---------------------------------------------------------------------
    Search = @{
        SearchIn    = @('Organizer', 'Rooms')
        PastDays    = 0
        FutureDays  = 365
        Rooms       = @()     # room addresses added to the places API (a new room may take time to appear there)
        RoomFile    = ''      # file of room addresses
        MailboxFile = ''      # default file of the 'Mailboxes' scope
    }

    # ---------------------------------------------------------------------
    # The actions (-Action Remove | Cancel; the default is Report: nothing is changed).
    #   CancelComment: the message of the cancellation sent by the organizer (Cancel), plain text.
    #   Verify: after the action, each copy is read again to check it is gone.
    # ---------------------------------------------------------------------
    Cleanup = @{
        CancelComment = 'This meeting has been cancelled by the IT department.'
        Verify        = $true
    }

    # ---------------------------------------------------------------------
    # Restore (-Action Restore -FromReport <report of a Remove run>, or Restore... in the window).
    # The copies removed are in Recoverable Items of each mailbox for the retention of deleted items
    # (14 days by default). Exchange Online PowerShell (module ExchangeOnlineManagement 3.2+) puts them back:
    # role Mailbox Import Export, in no role group by default (developer guide, chapter 5).
    #   Connection: 'Application'  the same application and certificate: Exchange.ManageAsApp permission and
    #                              its service principal in a role group with Mailbox Import Export;
    #                              Tenant.Organization must be the initial domain (contoso.onmicrosoft.com)
    #               'Interactive'  an administrator signs in (UserPrincipalName), with that role
    #   WindowMinutes: tolerance around the time of each removal, to find it in Recoverable Items.
    #   ReAccept: answer the restored copy again, silently (accepted, or tentative when there was no answer),
    #             so that a room is busy again: Exchange marked it Declined / Free when it was removed.
    # ---------------------------------------------------------------------
    Restore = @{
        Connection        = 'Application'   # Application | Interactive
        UserPrincipalName = ''              # Interactive: the administrator (empty = sign-in window)
        WindowMinutes     = 10
        ReAccept          = $true
    }

    # ---------------------------------------------------------------------
    # Transfer to a new organizer (-Action Transfer -NewOrganizer <address>, or the action of the window).
    #   Method: 'Auto'      Exchange Online moves the meetings whose organizer is still active (account and
    #                       mailbox), the other ones are re-created by the new organizer
    #           'Native'    Exchange Online only (Invoke-ChangeMeetingOrganizer: Exchange Online PowerShell with
    #                       the connection of the Restore section, role 'Meeting Organizer Transfer', developer guide chapter 5)
    #           'Recreate'  Microsoft Graph only: the new organizer sends one invitation, the old copies are removed
    #   Comment: the message of the old organizer when he is still active and his meeting is re-created
    #            ({0} = the new organizer). Plain text.
    # ---------------------------------------------------------------------
    Transfer = @{
        Method  = 'Auto'      # Auto | Native | Recreate
        Comment = 'This meeting is now organized by {0}.'
    }
    # ---------------------------------------------------------------------
    # Requests to Microsoft Graph ($batch of 20 requests, at most 4 at a time per mailbox).
    # ---------------------------------------------------------------------
    Graph = @{
        MaxConcurrency = 16      # $batch calls in flight (1-32): 16 = about 3,000 mailboxes searched per minute
        PageSize       = 500     # calendar items per page (10-1000)
        MaxRetries     = 6       # per request, for 429 / 5xx / no answer (the Retry-After delay is respected)
        TimeoutSeconds = 120
    }

    # ---------------------------------------------------------------------
    # Report files (one sub-folder per run), written locally only.
    # ---------------------------------------------------------------------
    Report = @{
        OutputPath   = '.\reports'
        FilePrefix   = 'MeetingCleanup'
        Formats      = @('Csv', 'Html')   # a Summary.json file is always written as well (-FromReport)
        CsvDelimiter = ';'                # ';' opens directly in Excel with French regional settings
        TimeZone     = ''                 # dates typed and shown: '' = the time zone of Windows, or Europe/Paris...
    }

    # ---------------------------------------------------------------------
    # Log files (one per day, no token, no colour).
    # ---------------------------------------------------------------------
    Logging = @{
        Path          = '.\logs'
        RetentionDays = 30
    }
}
