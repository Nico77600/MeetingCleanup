# Changelog — Meeting Cleanup

All notable changes are listed here. Versions follow MAJOR.MINOR.PATCH (see the guide, Appendix D).
Author: Nicolas Fabert.

## [1.2.2] — 2026-10-06

The progress of a run, at a glance.

### Added
- **Progress bar of the window**, the whole width of the *Progress* card: the step and what it counts
  (*Step 4/6 · Search · 1,254/1,861 mailboxes searched*), the part done and the **time left**
  (*67 % · about 20 s left*). While nothing is counted yet (connection, organizers, report), the bar moves and
  the time since the start is shown; it stops during a question (plan of a transfer or a restore).
- **Taskbar**: the button of the window shows the same progress (green), yellow once *Stop* is requested.
- **Time left in the console** too, at the end of the progress line. It is told from the speed of the
  progress in course, once it has run 2 s and 2 %, rounded as a person would say it.
- Window: at most ten updates a second whatever the volume (the queue of the run is read every 100 ms, the last
  part done only); the moving bar costs about 4 % of one processor core.
- **User guide** (`docs\MeetingCleanup-UserGuide.md` and `.html`): the prerequisites and one command per everyday
  question. The guide becomes the **developer guide**; both are in the package, and the links between them work
  on GitHub and in the HTML files.

## [1.2.1] — 2026-10-06

Performance: the window no longer stops responding on a large search, and a long list scrolls freely. Same
behaviour, same files, same configuration.

### Changed
- **Searches and actions run in the background** (window): *Search*, *Remove*, *Cancel*, *Transfer* and
  *Restore* run in a second PowerShell runspace opened with the window; the progress, the list and the buttons
  stay live, and *Stop* answers at once. The plans (Transfer, Restore) are built in the background as well,
  then confirmed in the window. Lab: 1,858 room mailboxes searched, the window never waited more than 163 ms.
- **Compiled engine** (`src\MeetingCleanup.Native.cs`, built by PowerShell when the module loads, no
  dependency): the meeting and copy objects, the filter of each mailbox's events, the totals, the recurrence
  text, the dates and the CSV / JSON files of the report. The loops of the search (copies, attendees, series)
  and of the plan, the backup and the removal no longer use the pipeline.
- **Lists of the window**: ListView rows of compiled objects instead of DataGrids (lighter rows), filled at
  once, sortable by a click on a column header.
- Measured on 600 meetings and 4,200 copies (`tools\Measure-MeetingCleanup.ps1`, both versions on the same
  computer, one after the other):

  | Step | 1.2.0 | 1.2.1 |
  |---|---|---|
  | Objects built from the Graph answers | 3.9 s | 1.1 s |
  | Totals | 1.3 s | 0.05 s |
  | Plan of a Remove | 0.9 s | 0.2 s |
  | Report (CSV, JSON, HTML) | 8.5 s | 1.0 s |
  | List filled (600 meetings) | 0.7 s | 0.25 s |
  | List scrolled top to bottom | 7.0 s | 4.1 to 5.0 s |
  | *Tick all* | 0.23 s | 0.02 to 0.06 s |
  | Module loaded (the C# part is compiled then) | 0.3 s | 1 to 1.6 s |

  On a busy computer the gaps grow (report 13.6 s → 2.3 s, list scrolled 24.7 s → 8.5 s). Above all, the
  window thread does none of this work any more.

### Added
- `tools\Measure-MeetingCleanup.ps1`: measures the steps above on synthetic data, without a tenant
  (`-Search`: a whole search on the simulated tenant of the tests; `-Gui`: the window).

## [1.2.0] — 2026-10-06

Rooms over a period, and the transfer of meetings to a new organizer.

### Added
- **Rooms mode** (`-Room`, `-RoomFile`, *Every meeting of rooms* in the window): every meeting of the rooms given
  in the period, whatever its organizer (present, deleted, the room itself), then every copy as usual. A series
  is limited to its **occurrences in the period**: each copy is replaced by its occurrences (own ID, column
  *Occurrence*); *Cancel* cancels these occurrences only (one cancellation each), *Remove* takes them out of
  the attendees and the rooms; the series goes on outside the period. Only the occurrences the rooms hold
  count (one moved to another room is left); a series whose every occurrence is held by the rooms in the
  period is handled whole. When the occurrences of the organizer or of a room cannot be read, the series is
  left as it is. An action needs the period (`-Start`, `-End`). Measured: a removed occurrence is not kept in
  Recoverable Items — reported *Not restorable* by Restore.
- **Transfer** (`-Action Transfer -NewOrganizer <address>`, `-TransferMethod Auto|Native|Recreate`,
  `-TransferFrom`, section `Transfer`): *Native* runs `Invoke-ChangeMeetingOrganizer` (Exchange Online
  PowerShell, least role *Meeting Organizer Transfer* from *User Options*) for an active organizer; *Recreate*
  creates the meeting again in the new organizer's calendar without attendees, shapes the series (occurrences
  removed or moved), removes the old room copies, then invites everyone at once (one invitation), and removes
  the old attendee copies — an active old organizer cancels his meeting with `Transfer.Comment`. *Auto*: native
  for an active organizer (meeting in his mailbox, account not deleted; not looked up counts as active), re-created
  otherwise (deleted or soft-deleted user). The series is re-created in its own time zone (the copy's, else the
  report's) and every occurrence still to come must find its slot in the new series, else the meeting is not
  transferred. A failure before the invitation removes the new meeting; old room copies already removed are
  restorable. *Stop* (window) waits for the end of the group being re-created (20 meetings). A Transfer report
  replayed skips the meetings it transferred and never removes the new meeting. Backup first.
- Report: columns *Scope*, *Occurrences*, *NewOrganizer*, *NewMeetingId*, *TransferMethod*, *Occurrence*;
  status *Transferred*; tiles and notices of the rooms mode and the transfer.
- Exchange Online connection shared by Restore and Transfer (`Connect-MclExchange -For`).

### Changed
- `Remove-MclCopyWaves`: the removal in waves shared by Remove, Cancel and Transfer.
- Console and window count the copies by mailbox (occurrences once); *Kind* shows `3 occ.` for occurrences.
- Guide: the unblock note at the top, `Install-Module ... -Force`; `tools\New-ReadmeImages.ps1` renders the
  graphics of the GitHub page.

### Fixed
- `Import-MclReport` and the restore read the reports of 1.0.0 and 1.1.0 (properties added since).
- An organizer copy that cannot be read (access denied, throttled) is *Not read*, no longer *Mailbox deleted*:
  *Cancel* leaves such a meeting as it is instead of removing the attendees' copies without a message.

## [1.1.0] — 2026-10-06

A list of organizers, a backup before any change, and the restore of a silent removal.

### Added
- **List of organizers**: `-OrganizerFile` (text file, one address per line, or CSV with `PrimarySmtpAddress`,
  `Mail`, `UserPrincipalName`, `Organizer`, `LegacyExchangeDN`...), several lines in the window and
  *Load a list...*. The organizers are resolved in one `$batch` (addresses, then mailbox); each mailbox is
  read once for all of them. An *Organizer* column in the console, the window and the report when there are
  several; a tab and a file `Organizers.csv` with what was found and done for each.
- **Backup** before Remove or Cancel: `Backup.json` in the folder of the run, with every meeting read in full
  (body, attendees, recurrence) and the state of each copy (response, free/busy).
- **Restore** (`-Action Restore -FromReport <Remove run>`, *Restore...* in the window): the copies removed by
  a run come back from Recoverable Items (Exchange Online PowerShell, `Get-RecoverableItems` /
  `Restore-RecoverableItems -EntryID`: the same item, same ID), are verified by iCalUId, then answered again
  without a message (`sendResponse = false`) so that the rooms are busy again and the organizer's tracking is
  right. Copies already back are left alone. A restore never removes anything: when the item of a copy is not
  certain (same subject, same second, or not the copies the run removed), nothing is restored for that subject
  in that mailbox and the command to do it by hand is given. The same restore again finishes a stopped one
  (copies already back but still Declined / Free are answered again); a copy not verified or not answered
  again makes the run *Warning*. Connection as the application (`Exchange.ManageAsApp` and a role group with
  *Mailbox Import Export*) or as an administrator (`Restore.Connection = 'Interactive'`). Works with the
  reports of 1.0.0. A meeting cancelled by its organizer is reported *Not restorable*.
- Configuration section `Restore` (`Connection`, `UserPrincipalName`, `WindowMinutes`, `ReAccept`).
- Copies: columns *Organizer*, *ShowAs* and *ActionUtc* (time of each removal).

### Changed
- Remove removes the copies with the same subject in the same mailbox one after the other, 3 seconds apart
  (their order is then certain in Recoverable Items); the time of each removal is the end of its own `$batch`
  call. The console and the window give the exact `-FromReport` command to run next.
- Window: the *Copies* column gives the rooms (`4 (1 room)`); the *Organizer* column only with several
  organizers; the questions go through one function.

### Fixed
- A meeting of another organizer of the list, found in the calendar of the first one, was taken as his.
- `Import-MclReport` called from a script without `-MeetingId` kept no meeting (an empty ID was searched).
- *Stop* during a removal: the copies already removed are kept in the report (with their time, so that they
  can be restored), in the folder of the run next to the backup.

## [1.0.0] — 2026-10-05

First version: one tool that replaces the two scripts *purge of an organizer* (rooms, then attendees,
then organizer) and *global purge* (every mailbox searched for a departed organizer).

### Added
- **One search for every case**: one meeting (`-Subject`, `-MeetingId`), a series (handled as a whole) or
  every meeting of a period (`-Start`, `-End`); organizer present or deleted (SMTP, any alias, or X500
  address); searched in the organizer's calendar, every room (places API, `Search.Rooms`, `Search.RoomFile`),
  a list of mailboxes or every mailbox of the tenant (`-SearchIn`).
- **Every copy**: the attendee list of one copy gives every internal attendee, room and member of an
  invited group (transitive members), each asked for its copy by iCalUId. External, deleted and on-premises
  attendees are listed as *Not processed*, with the reason. A copy reached through an alias counts once.
- **Remove** (silent: the copies of the attendees and the rooms, `permanentDelete`; the organizer's meeting
  is kept, removing it there always sends a cancellation) and **Cancel** (the organizer cancels with a
  message, then the copies left are removed). A meeting without an organizer copy is removed silently.
  If a cancellation fails, the copies of that meeting are left untouched.
- **Report by default**; Remove and Cancel show the plan and ask to type YES (`-Force`). Verification of
  every copy removed. `-FromReport` replays a reviewed report by the IDs of its meetings.
- Microsoft Graph app-only: certificate (client assertion built by the tool, no module) or client secret
  (environment variable or typed, never written). Tenant of the token checked; `Calendars.Read` enough for a
  report; the optional permissions (`User.Read.All`, `Place.Read.All`, `GroupMember.Read.All`) reported
  when missing.
- `$batch` scheduler: 20 requests per call, 16 calls in flight, at most 4 requests at a time per mailbox,
  429 / 5xx retried after *Retry-After*, pages followed. About 3,000 to 5,000 mailboxes searched a minute.
- **Window** (`-Gui`), WPF with the Fluent theme of Windows 11 (light or dark, accent #B11F4B): organizer,
  period, subject, where to search, meetings with a box to tick and their copies, progress, Remove or Cancel
  with the message, confirmation, report.
- Reports: `Meetings.csv`, `Copies.csv`, `Summary.json`, self-contained HTML dashboard (light / dark).
  Daily log with every line, the confirmations and the Graph statistics.
- Guide (Markdown and self-contained HTML), 33 Pester tests on a simulated Exchange Online tenant,
  images of the guide rendered from fictitious data, package tool.

### Lab measurements (Graph v1.0, 2026-10-05)
- Removing a meeting from the organizer's calendar, with `DELETE` or `permanentDelete`, sends a cancellation
  to every attendee, external ones included: the former scripts were not silent for the organizer.
- Removing an attendee's or a room's copy sends nothing to the organizer.
- `$filter` on the organizer is refused by Graph (501): the organizer is compared by the tool.
- The iCalUId is the same in every copy of a meeting; a room copy shows the organizer's name as subject.
