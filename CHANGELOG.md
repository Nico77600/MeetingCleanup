# Changelog — Meeting Cleanup

All notable changes are listed here. Versions follow MAJOR.MINOR.PATCH (see the developer guide, Appendix D).
Author: Nicolas Fabert.

## [1.3.1] — 2026-11-05

Signed and verified package, and a Get started that begins with the release ZIP.

### Added
- The PowerShell files carry the owner's timestamped Authenticode signature.
- `PackageIntegrity.psd1`, a signed list of the SHA256 of every shipped file: the tool checks it before anything loads and refuses a changed, added or missing file.
- User guide: a "Check the download" box, and numbered steps for the application registration, its four permissions and the certificate.
- Before it starts, the tool checks the signature and SHA256 of every file of the package and stops if one is changed, missing or added.
- `-DevelopmentMode` starts an unsigned folder that has no manifest (a present manifest is always verified).

### Changed
- Only `config\MeetingCleanup.config.psd1`, `reports\` and `logs\` may change inside the tool folder; keep your own files (lists, certificate export) outside it.
- The C# helper is compiled in memory from the verified source at each start (a few seconds: 3.6 s measured for the whole module import on a busy computer); no compiled file is cached on disk.
- The release ZIP has one root folder, `MeetingCleanup-<version>\`, holds `PackageIntegrity.psd1`, LICENSE, CHANGELOG and THIRD-PARTY-NOTICES, and its README links only files it contains.
- Get started begins with the release ZIP and gives the PowerShell version check, each step with what you should see.
- First run: a report first (nothing is changed), then the same command with the action.
- The run-time files are in the `package` folder of the repository; the source archive leaves out tests, tools and the Markdown guides.

## [1.3.0] — 2026-10-07

### Added
- Series by occurrences: `-SeriesScope Occurrences` limits a series to its occurrences in the period; with a period of one day, one occurrence.
- *Cancel* sends one cancellation per occurrence; *Remove* takes the occurrences out of the attendees' and rooms' calendars without a message.
- **Occurrences...** in the window: tick the occurrences to act on. The report keeps the choice and a replay with `-FromReport` too.

## [1.2.3] — 2026-10-07

### Added
- **Transfers** tab in the HTML report of a *Transfer* run, and `MeetingCleanup-Transfers.csv`: for each meeting, from whom to whom, how, the new meeting and what became of the old one.

### Changed
- `Calendars.ReadWrite.All` is accepted like `Calendars.ReadWrite`.

## [1.2.2] — 2026-10-06

### Added
- Progress bar in the window: the step, the part done and the time left.
- The button of the window in the taskbar shows the same progress.
- The console shows the time left too.
- User guide: the prerequisites and one command per everyday question. The previous guide becomes the developer guide.

## [1.2.1] — 2026-10-06

### Changed
- Searches and actions run in the background: the window keeps answering during a large search and *Stop* acts at once.
- Faster lists, reports and totals on a large search.

## [1.2.0] — 2026-10-06

### Added
- **Rooms mode** (`-Room`, `-RoomFile`): every meeting of the rooms in a period, whatever its organizer. A series loses only its occurrences in the period.
- **Transfer** (`-Action Transfer -NewOrganizer <address>`): moved by Exchange Online, or re-created by the new organizer when the old mailbox is gone.
- Report columns and tiles for rooms mode and transfers.

### Changed
- Guide: the unblock note at the top and `Install-Module ... -Force`.
- Reports of 1.0.0 and 1.1.0 can be read and restored.
- A copy that cannot be read is reported *Not read*; *Cancel* leaves such a meeting as it is.

## [1.1.0] — 2026-10-06

### Added
- List of organizers: `-OrganizerFile`, several lines in the window, *Load a list...*; an *Organizers* tab and file in the report.
- Backup (`Backup.json`) before Remove or Cancel.
- **Restore** (`-Action Restore -FromReport <Remove run>`, *Restore...* in the window): the copies removed come back from Recoverable Items, without a message.
- Configuration section `Restore`.

### Changed
- The console and the window give the exact `-FromReport` command to run next.
- The *Copies* column of the window shows the rooms.
- *Stop* during a removal keeps the copies already removed in the report.
- A meeting of another organizer of the list is not taken as the first one's.

## [1.0.0] — 2026-10-05

First version: one tool for the meetings of an organizer who left, present or deleted mailbox.

### Added
- One search for one meeting, a series or a period, in the organizer's calendar, the rooms, a list of mailboxes or every mailbox.
- Every copy of a meeting is found: attendees, rooms and members of the groups invited.
- **Remove** (silent) and **Cancel** (by the organizer, with your message); a report is the default and an action asks to type YES.
- Microsoft Graph application permissions with a certificate or a client secret.
- A window (`-Gui`), CSV and JSON reports, a self-contained HTML report and a daily log.
- User guide in Markdown and HTML.
