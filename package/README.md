# Meeting Cleanup

Version: 1.3.1

Meeting Cleanup finds the meetings of organizers who left or stay, or every meeting of some rooms, in every calendar where they are in Exchange Online. Then it removes them silently, has the organizer cancel them, or transfers them to a new organizer, even when the old mailbox is gone.

This folder is the complete release: the tool, its configuration, the report template and both guides. Open `docs\MeetingCleanup-UserGuide.html` in a browser for the full steps.

## What you need

- Exchange Online only.
- PowerShell 7.4 or later (not Windows PowerShell 5.1). Open PowerShell 7 with `pwsh`, then run `$PSVersionTable.PSVersion`: the version must be 7.4 or later.
- Windows 10 / 11 or Windows Server 2016 to 2025.
- An application registered in Microsoft Entra, with a certificate. The user guide, chapter 1, gives every click.

## Get started

**1. Check the download.**

Do: open PowerShell 7 in this folder and run:

```powershell
Get-AuthenticodeSignature .\Invoke-MeetingCleanup.ps1 | Format-List Status, SignerCertificate
```

You should see: `Status : Valid` and the signer `Nicolas Fabert`.
If not: download the ZIP again from the [Releases page](https://github.com/Nico77600/MeetingCleanup/releases/latest). Use the ZIP named `MeetingCleanup-<version>.zip`, not the *Source code* archives.

**2. Unblock the files.** Windows may block files downloaded from the Internet. Run this once, from this folder:

```powershell
Get-ChildItem . -Recurse -File -Force | Unblock-File
```

You should see: nothing, and no error.

**3. Fill in the configuration.** Keep your own files (lists of addresses, certificate exports) outside the tool folder, for example in `C:\Tools\MeetingCleanup-data`, and give their absolute path.

Do: `notepad .\config\MeetingCleanup.config.psd1`, then enter your tenant ID, the application (client) ID and the certificate thumbprint.
You should see: the same file, saved, with your three values.
If not: the user guide, chapter 1, shows how to create the application and the certificate.

**4. First run: a report.** A report changes nothing.

```powershell
.\Invoke-MeetingCleanup.ps1 -Organizer megan.bowen@contoso.com
```

You should see: the steps of the run, the meetings found and the folder `reports\MeetingCleanup_Report_<date>` with `MeetingCleanup.html`. Open that file in a browser.
If not: the user guide, chapter 4, explains the usual messages.

**5. Then the same command with the action.** The tool shows exactly what it will do and asks you to type YES.

```powershell
.\Invoke-MeetingCleanup.ps1 -Organizer megan.bowen@contoso.com -Action Cancel -Comment 'Megan has left the company.'
```

You can also run `.\Invoke-MeetingCleanup.ps1 -Gui` for the window. Remove, Transfer and Restore are in the user guide, chapter 2.

## Content

| Item | Role |
|---|---|
| Invoke-MeetingCleanup.ps1 | Entry script for the window and the command line. |
| MeetingCleanup.psd1, MeetingCleanup.psm1 | The module. |
| PackageIntegrity.psd1 | Signed list of the SHA256 of every file, checked before anything loads. Do not edit. |
| README.md | This file. |
| config | The configuration to fill in. |
| docs | User guide and developer guide (HTML, one file each, open them in a browser). |
| src | PowerShell and C# source files of the tool. |
| templates | HTML report template. |
| LICENSE | MIT license. |
| THIRD-PARTY-NOTICES.md | Third-party notices. |
| CHANGELOG.md | Changes of each version. |

## More

- [User guide](docs/MeetingCleanup-UserGuide.html): prerequisites, setup and one command per everyday question.
- [Developer guide](docs/MeetingCleanup-Guide.html): how it works, rights, settings and troubleshooting.
- [Changes](CHANGELOG.md) · [License](LICENSE) · [Third-party notices](THIRD-PARTY-NOTICES.md)

Project page and releases: https://github.com/Nico77600/MeetingCleanup

## Disclaimer

This Script is a Personal project.
It's provided "AS-IS". It's not an official Microsoft product so no support can be expected from Microsoft.

As any scripts you must read carefully the documentation and test it first in a Test environment before any run in Production.

Removing, cancelling or transferring meetings changes real calendars: always start with a report, and test it in your environment before production use.
