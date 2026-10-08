# Meeting Cleanup

Meeting Cleanup finds the meetings of organizers who left or stay, or every meeting of some rooms, in every calendar where they are in Exchange Online, then removes them silently, has the organizer cancel them, or transfers them to a new organizer, even when the old mailbox is gone.

This folder contains everything needed to run the tool: Invoke-MeetingCleanup.ps1, the module, the configuration, the report template and the guides. Tests and build tools stay outside it, in the repository.

> [!IMPORTANT]
> Files downloaded from the Internet may be blocked by Windows. Unblock them once, from this folder:
>
> ```powershell
> Get-ChildItem . -Recurse -File | Unblock-File
> ```

## Requirements
- Exchange Online only.
- PowerShell 7.4 or later.
- Windows 10 / 11 or Windows Server 2016 to 2025.
- Microsoft Entra application permissions: Calendars.ReadWrite (or Calendars.ReadWrite.All), User.Read.All, Place.Read.All and GroupMember.Read.All.
- Certificate private key in the Windows store of the account that runs the tool.
- ExchangeOnlineManagement 3.2+ and roles are needed for Restore or native Transfer.

## Quick start
```powershell
notepad .\config\MeetingCleanup.config.psd1          # tenant, application, certificate thumbprint

.\Invoke-MeetingCleanup.ps1 -Gui                     # the window: search, untick, act, restore

# Or the command line: always a report first (nothing is changed), then the same command with the action
.\Invoke-MeetingCleanup.ps1 -Organizer megan.bowen@contoso.com
.\Invoke-MeetingCleanup.ps1 -Organizer megan.bowen@contoso.com -Action Cancel -Comment 'Megan has left the company.'
.\Invoke-MeetingCleanup.ps1 -Organizer john.doe@contoso.com -SearchIn Rooms, AllMailboxes -Action Transfer -NewOrganizer jane.roe@contoso.com
.\Invoke-MeetingCleanup.ps1 -Room room-paris-01@contoso.com -Start 2026-11-02 -End 2026-11-13 -Action Cancel -Comment 'Closed for works.'
.\Invoke-MeetingCleanup.ps1 -Organizer megan.bowen@contoso.com -Subject 'Weekly sales review' -SeriesScope Occurrences -Start 2026-11-16 -End 2026-11-16 -Action Cancel -Comment 'Not this Monday.'
.\Invoke-MeetingCleanup.ps1 -Action Restore -FromReport .\reports\MeetingCleanup_Remove_20261105-093000
```

## Content
| Item | Role |
|---|---|
| Invoke-MeetingCleanup.ps1 | Entry script for the window and command line. |
| MeetingCleanup.psd1 | Module manifest. |
| MeetingCleanup.psm1 | Module loader. |
| config | Delivered configuration template. |
| docs | User and developer guides in Markdown and HTML, with images. |
| src | PowerShell source files and native helper source. |
| templates | HTML report template. |
| LICENSE | MIT license. |
| THIRD-PARTY-NOTICES.md | Third-party notices. |

## Documentation
- [User guide](docs/MeetingCleanup-UserGuide.md) - also `docs/MeetingCleanup-UserGuide.html`, a single file to open locally
- [Developer guide](docs/MeetingCleanup-Guide.md) - also `docs/MeetingCleanup-Guide.html`

Project page, releases and change log: https://github.com/Nico77600/MeetingCleanup

License: [MIT](LICENSE).
