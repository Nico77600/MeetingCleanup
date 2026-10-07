// Meeting Cleanup - compiled helpers, loaded once by MeetingCleanup.psm1 (Add-Type).
//
// A PowerShell function call costs tens of microseconds; the steps below run once per calendar item, copy or
// report cell - hundreds of thousands of times on a large tenant. They are compiled here so that a search of
// thousands of copies takes seconds, not minutes. They behave exactly like the PowerShell code they replace:
//
//   Fast      property reads of Graph answers (no exception when a property is missing), the copy object,
//             dates, the report rows, the CSV cells (formula injection neutralised) and the JSON of the report
//   GuiRows   the rows of the window (compiled objects: WPF binds and scrolls them without PowerShell)
//
// Author : Nicolas Fabert
// Version: 1.2.3

using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Management.Automation;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;

namespace MeetingCleanupNative
{
    public static class Fast
    {
        public const string Version = "1.2.3";
        static readonly CultureInfo Inv = CultureInfo.InvariantCulture;
        static readonly string[] CopyRoles = { "Organizer", "Attendee", "Room" };

        // ---- properties ------------------------------------------------------------------------------

        static object Base(object o)
        {
            var p = o as PSObject;
            if (p != null && !(p.BaseObject is PSCustomObject)) { return p.BaseObject; }
            return o;
        }

        /// <summary>A property of an object or a key of a dictionary, or null (Get-MclProperty).</summary>
        public static object Prop(object o, string name)
        {
            if (o == null) { return null; }
            var b = Base(o);
            var d = b as IDictionary;
            if (d != null) { return d.Contains(name) ? d[name] : null; }
            var p = PSObject.AsPSObject(o).Properties[name];
            return p == null ? null : p.Value;
        }

        /// <summary>A property along a path (organizer, emailAddress, address), or null.</summary>
        public static object Path(object o, params string[] names)
        {
            foreach (var n in names) { o = Prop(o, n); if (o == null) { return null; } }
            return o;
        }

        /// <summary>A property as text ([string] of PowerShell); "" when missing.</summary>
        public static string Text(object o, params string[] names)
        {
            return ToText(Path(o, names));
        }

        /// <summary>A property as a boolean (PowerShell truth: [bool]).</summary>
        public static bool Flag(object o, string name)
        {
            return LanguagePrimitives.IsTrue(Prop(o, name));
        }

        public static string ToText(object v)
        {
            if (v == null) { return ""; }
            v = Base(v);
            var s = v as string;
            if (s != null) { return s; }
            return (string)LanguagePrimitives.ConvertTo(v, typeof(string), Inv) ?? "";
        }

        // ---- objects of the tool ---------------------------------------------------------------------

        static readonly TextInfo English = new CultureInfo("en-US").TextInfo;

        /// <summary>A Graph patternedRecurrence in words: Weekly (Tuesday), 6 occurrences from 2026-10-13 (Format-MclRecurrence).</summary>
        public static string Recurrence(object recurrence)
        {
            if (!LanguagePrimitives.IsTrue(recurrence)) { return ""; }
            var p = Prop(recurrence, "pattern"); var r = Prop(recurrence, "range");
            var type = Text(p, "type");
            int interval = ToInt(Prop(p, "interval"));
            var days = new List<string>();
            foreach (var d in Items(Prop(p, "daysOfWeek"))) { var s = ToText(d); if (s.Length > 0) { days.Add(English.ToTitleCase(s)); } }
            var dayList = string.Join(", ", days);
            var every = interval > 1 ? "every " + interval.ToString(Inv) + " " : "";
            string text;
            switch (type)
            {
                case "daily": text = every.Length > 0 ? every + "days" : "Daily"; break;
                case "weekly": text = (every.Length > 0 ? every + "weeks" : "Weekly") + " (" + dayList + ")"; break;
                case "absoluteMonthly": text = (every.Length > 0 ? every + "months" : "Monthly") + " (day " + Text(p, "dayOfMonth") + ")"; break;
                case "relativeMonthly": text = (every.Length > 0 ? every + "months" : "Monthly") + " (" + Text(p, "index") + " " + dayList + ")"; break;
                case "absoluteYearly": text = "Yearly (day " + Text(p, "dayOfMonth") + " of month " + Text(p, "month") + ")"; break;
                case "relativeYearly": text = "Yearly (" + Text(p, "index") + " " + dayList + " of month " + Text(p, "month") + ")"; break;
                default: text = type; break;
            }
            var from = Text(r, "startDate");
            switch (Text(r, "type"))
            {
                case "endDate": text += ", from " + from + " until " + Text(r, "endDate"); break;
                case "numbered": text += ", " + Text(r, "numberOfOccurrences") + " occurrences from " + from; break;
                default: text += ", from " + from + ", no end"; break;
            }
            return text;
        }

        /// <summary>A meeting found in a calendar (New-MclMeeting): the same properties, in the same order.</summary>
        public static PSObject NewMeeting(string key, object ev, TimeZoneInfo zone)
        {
            var start = ToUtc(Prop(ev, "start"));
            var end = ToUtc(Prop(ev, "end"));
            var recurrence = Prop(ev, "recurrence");
            return Row(
                "MeetingId", key ?? "", "Subject", Text(ev, "subject"), "Organizer", Text(ev, "organizer", "emailAddress", "address").ToLowerInvariant(),
                "OrganizerName", Text(ev, "organizer", "emailAddress", "name"), "OrganizerKey", "", "Kind", Text(ev, "type") == "seriesMaster" ? "Series" : "Single",
                "Start", start, "End", end, "StartText", FormatDate(start, zone, false, false), "EndText", FormatDate(end, zone, false, false), "NextInPeriod", "",
                "Recurrence", Recurrence(recurrence), "Location", "", "Cancelled", Flag(ev, "isCancelled"), "OrganizerCopy", "Not checked",
                "Attendees", new object[0], "Copies", new List<object>(), "Selected", true, "Status", "Found", "Notes", new List<string>(), "SubjectFromRoom", false,
                "Scope", "Whole", "Occurrences", 0, "RecurrenceData", null, "TimeZone", "", "NewOrganizer", "", "NewMeetingId", "", "TransferMethod", "");
        }

        /// <summary>
        /// The calendar items of one mailbox that belong to the search: organized by an organizer searched (or, rooms
        /// mode, every meeting of the room), with an iCalUId, among the meeting IDs asked for. One call per mailbox.
        /// </summary>
        public static List<EventMatch> MatchEvents(object values, IDictionary organizers, object mailboxOrganizer, bool roomsMode, string mailbox, ICollection<string> ids)
        {
            var list = new List<EventMatch>();
            foreach (var ev in Items(values))
            {
                var address = Text(ev, "organizer", "emailAddress", "address").ToLowerInvariant();
                bool own; string organizerKey;
                if (roomsMode)
                {
                    own = Flag(ev, "isOrganizer");
                    organizerKey = own ? (mailbox ?? "") : address;
                    if (organizerKey.Length == 0) { continue; }
                }
                else
                {
                    own = mailboxOrganizer != null && Flag(ev, "isOrganizer");
                    var org = own ? mailboxOrganizer : (organizers != null && address.Length > 0 && organizers.Contains(address) ? organizers[address] : null);
                    if (org == null) { continue; }
                    organizerKey = Text(org, "PrimaryAddress");
                }
                var key = Text(ev, "iCalUId").ToUpperInvariant();
                if (key.Length == 0) { continue; }
                if (ids != null && ids.Count > 0 && !ids.Contains(key)) { continue; }
                list.Add(new EventMatch { Event = ev, Key = key, Own = own, OrganizerKey = organizerKey, IsSeriesMaster = Text(ev, "type") == "seriesMaster", Recurrence = Prop(ev, "recurrence") });
            }
            return list;
        }

        /// <summary>The totals of a result (Update-MclResultCounts), in one pass over the copies.</summary>
        public static PSObject Counts(object meetings, int organizers)
        {
            int count = 0, series = 0, selected = 0, transferred = 0, copies = 0, roomCopies = 0, organizerCopies = 0, occurrenceCopies = 0;
            var mailboxes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var byResult = new Dictionary<string, int>(StringComparer.Ordinal);
            foreach (var m in Items(meetings))
            {
                var status = Text(m, "Status");
                if (status == "Transferred") { transferred++; }
                if (status == "Appointment") { continue; }
                count++;
                if (Text(m, "Kind") == "Series") { series++; }
                if (Flag(m, "Selected")) { selected++; }
                foreach (var c in Items(Prop(m, "Copies")))
                {
                    var result = Text(c, "Result");
                    if (result.Length > 0) { int n; byResult.TryGetValue(result, out n); byResult[result] = n + 1; }
                    if (Text(c, "EventId").Length == 0) { continue; }
                    copies++;
                    mailboxes.Add(Text(c, "Mailbox"));
                    var role = Text(c, "Role");
                    if (role == "Room") { roomCopies++; } else if (role == "Organizer") { organizerCopies++; }
                    if (Text(c, "Occurrence").Length > 0) { occurrenceCopies++; }
                }
            }
            Func<string, int> of = k => { int n; return byResult.TryGetValue(k, out n) ? n : 0; };
            return Row(
                "Meetings", count, "Series", series, "Selected", selected, "Copies", copies, "Mailboxes", mailboxes.Count, "RoomCopies", roomCopies, "OrganizerCopies", organizerCopies,
                "NotProcessed", of("Not processed"), "Removed", of("Removed"), "Cancelled", of("Cancelled"), "AlreadyGone", of("Already gone"), "Kept", of("Kept"), "Failed", of("Failed"),
                "Restored", of("Restored"), "NotFound", of("Not found"), "AlreadyPresent", of("Already present"), "NotRestorable", of("Not restorable"),
                "Transferred", transferred, "OccurrenceCopies", occurrenceCopies, "Organizers", organizers);
        }

        /// <summary>A copy of a meeting in one mailbox (New-MclCopy).</summary>
        public static PSObject NewCopy(string key, string mailbox, string role, string via, object ev, string result, string detail)
        {
            var o = new PSObject();
            var p = o.Properties;
            p.Add(new PSNoteProperty("MeetingId", key ?? ""));
            p.Add(new PSNoteProperty("Mailbox", (mailbox ?? "").ToLowerInvariant()));
            p.Add(new PSNoteProperty("Role", role ?? ""));
            p.Add(new PSNoteProperty("Via", via ?? ""));
            p.Add(new PSNoteProperty("EventId", Text(ev, "id")));
            p.Add(new PSNoteProperty("Subject", Text(ev, "subject")));
            p.Add(new PSNoteProperty("Response", Text(ev, "responseStatus", "response")));
            p.Add(new PSNoteProperty("ShowAs", Text(ev, "showAs")));
            p.Add(new PSNoteProperty("Cancelled", Flag(ev, "isCancelled")));
            p.Add(new PSNoteProperty("Action", ""));
            p.Add(new PSNoteProperty("Result", result ?? ""));
            p.Add(new PSNoteProperty("HttpStatus", 0));
            p.Add(new PSNoteProperty("Detail", detail ?? ""));
            p.Add(new PSNoteProperty("Verified", ""));
            p.Add(new PSNoteProperty("ActionUtc", ""));
            p.Add(new PSNoteProperty("Occurrence", ""));
            p.Add(new PSNoteProperty("OccurrenceStart", ""));
            p.Add(new PSNoteProperty("SeriesId", ""));
            return o;
        }

        /// <summary>A Graph dateTimeTimeZone (UTC by default) as a UTC DateTime, or null (ConvertTo-MclDateUtc).</summary>
        public static object ToUtc(object value)
        {
            if (value == null) { return null; }
            var raw = Base(Prop(value, "dateTime"));
            if (raw == null) { return null; }
            DateTime d;
            if (raw is DateTime) { d = (DateTime)raw; if (d.Kind == DateTimeKind.Local) { d = DateTime.SpecifyKind(d, DateTimeKind.Unspecified); } }
            else
            {
                var text = ToText(raw);
                if (text.Length == 0) { return null; }
                d = DateTime.Parse(text, Inv, DateTimeStyles.RoundtripKind);
            }
            if (d.Kind == DateTimeKind.Utc) { return d; }
            var zone = Text(value, "timeZone");
            if (zone.Length == 0 || zone == "UTC") { return DateTime.SpecifyKind(d, DateTimeKind.Utc); }
            try { return TimeZoneInfo.ConvertTimeToUtc(DateTime.SpecifyKind(d, DateTimeKind.Unspecified), TimeZoneInfo.FindSystemTimeZoneById(zone)); }
            catch (Exception) { return DateTime.SpecifyKind(d, DateTimeKind.Utc); }
        }

        /// <summary>A UTC date shown in a time zone: yyyy-MM-dd HH:mm or yyyy-MM-dd (Format-MclDate).</summary>
        public static string FormatDate(object utc, TimeZoneInfo zone, bool dateOnly, bool periodEnd)
        {
            if (utc == null) { return ""; }
            utc = Base(utc);
            DateTime d;
            if (utc is DateTime) { d = (DateTime)utc; }
            else
            {
                var text = ToText(utc);
                if (text.Length == 0) { return ""; }
                d = DateTime.Parse(text, Inv, DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal);
            }
            d = DateTime.SpecifyKind(d.ToUniversalTime(), DateTimeKind.Utc);
            var local = TimeZoneInfo.ConvertTimeFromUtc(d, zone ?? TimeZoneInfo.Local);
            if (periodEnd && local.TimeOfDay == TimeSpan.Zero) { local = local.AddDays(-1); dateOnly = true; }
            return local.ToString(dateOnly ? "yyyy-MM-dd" : "yyyy-MM-dd HH:mm", Inv);
        }

        static IEnumerable Items(object o)
        {
            if (o == null) { yield break; }
            var b = Base(o);
            if (b is string || b is IDictionary) { yield return o; yield break; }
            var e = b as IEnumerable;
            if (e == null) { yield return o; yield break; }
            foreach (var x in e) { if (x != null) { yield return x; } }
        }

        /// <summary>The copies of a meeting found in a calendar, one per mailbox (an occurrence copy counts once for its mailbox).</summary>
        public static List<object> RealCopies(object meeting)
        {
            var list = new List<object>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var c in Items(Prop(meeting, "Copies")))
            {
                if (Text(c, "EventId").Length == 0 || Array.IndexOf(CopyRoles, Text(c, "Role")) < 0) { continue; }
                if (seen.Add(Text(c, "Mailbox"))) { list.Add(c); }
            }
            return list;
        }

        static int CountRole(List<object> copies, string role)
        {
            int n = 0;
            foreach (var c in copies) { if (Text(c, "Role") == role) { n++; } }
            return n;
        }

        // ---- report ----------------------------------------------------------------------------------

        static PSObject Row(params object[] pairs)
        {
            var o = new PSObject();
            for (int i = 0; i < pairs.Length; i += 2) { o.Properties.Add(new PSNoteProperty((string)pairs[i], pairs[i + 1])); }
            return o;
        }

        /// <summary>One row per meeting, for the CSV and the HTML (Get-MclMeetingRows).</summary>
        public static Table MeetingTable(object meetings)
        {
            var t = new Table("MeetingId", "Subject", "Organizer", "OrganizerName", "Kind", "Scope", "Occurrences", "NewOrganizer", "NewMeetingId", "TransferMethod",
                "StartText", "EndText", "NextInPeriod", "Recurrence", "Location", "OrganizerCopy", "Copies", "RoomCopies", "AttendeeCopies", "NotProcessed",
                "Cancelled", "Selected", "Status", "Notes");
            foreach (var m in Items(meetings))
            {
                var copies = RealCopies(m);
                int notProcessed = 0;
                foreach (var c in Items(Prop(m, "Copies"))) { if (Text(c, "Result") == "Not processed") { notProcessed++; } }
                var notes = new List<string>();
                foreach (var n in Items(Prop(m, "Notes"))) { notes.Add(ToText(n)); }
                t.Rows.Add(new object[] {
                    Text(m, "MeetingId"), Text(m, "Subject"), Text(m, "Organizer"), Text(m, "OrganizerName"), Text(m, "Kind"), Text(m, "Scope"), ToInt(Prop(m, "Occurrences")),
                    Text(m, "NewOrganizer"), Text(m, "NewMeetingId"), Text(m, "TransferMethod"),
                    Text(m, "StartText"), Text(m, "EndText"), Text(m, "NextInPeriod"), Text(m, "Recurrence"), Text(m, "Location"), Text(m, "OrganizerCopy"),
                    copies.Count, CountRole(copies, "Room"), CountRole(copies, "Attendee"), notProcessed,
                    Flag(m, "Cancelled"), Flag(m, "Selected"), Text(m, "Status"), notes.ToArray() });
            }
            return t;
        }

        static int ToInt(object v)
        {
            if (v == null) { return 0; }
            try { return (int)LanguagePrimitives.ConvertTo(v, typeof(int), Inv); } catch (Exception) { return 0; }
        }

        /// <summary>One row per copy (Get-MclCopyRows).</summary>
        public static Table CopyTable(object meetings)
        {
            var t = new Table("MeetingId", "MeetingSubject", "Organizer", "Mailbox", "Role", "Via", "Occurrence", "Response", "ShowAs", "Cancelled", "Action", "Result",
                "HttpStatus", "Verified", "ActionUtc", "Detail", "EventId");
            foreach (var m in Items(meetings))
            {
                var id = Text(m, "MeetingId"); var subject = Text(m, "Subject"); var organizer = Text(m, "Organizer");
                foreach (var c in Items(Prop(m, "Copies")))
                {
                    var status = Prop(c, "HttpStatus");
                    t.Rows.Add(new object[] {
                        id, subject, organizer, Text(c, "Mailbox"), Text(c, "Role"), Text(c, "Via"), Text(c, "Occurrence"), Text(c, "Response"), Text(c, "ShowAs"),
                        Flag(c, "Cancelled"), Text(c, "Action"), Text(c, "Result"), LanguagePrimitives.IsTrue(status) ? (object)ToInt(status) : "",
                        Text(c, "Verified"), Text(c, "ActionUtc"), Text(c, "Detail"), Text(c, "EventId") });
                }
            }
            return t;
        }

        /// <summary>One row per organizer of the run, with what was found and done for its meetings (Get-MclOrganizerRows).</summary>
        public static Table OrganizerTable(object organizers, object meetings)
        {
            var t = new Table("Input", "DisplayName", "PrimaryAddress", "State", "Detail", "Meetings", "Series", "Copies", "Removed", "Cancelled", "Restored", "Transferred", "Failed");
            var byKey = new Dictionary<string, List<object>>(StringComparer.OrdinalIgnoreCase);
            foreach (var m in Items(meetings))
            {
                var k = Text(m, "OrganizerKey");
                if (k.Length == 0) { k = Text(m, "Organizer"); }
                List<object> list;
                if (!byKey.TryGetValue(k, out list)) { list = new List<object>(); byKey[k] = list; }
                list.Add(m);
            }
            foreach (var o in Items(organizers))
            {
                var mine = new List<object>();
                var counted = new HashSet<object>(ReferenceEqualityComparer.Instance);
                var keys = new List<string>();
                foreach (var a in Items(Prop(o, "Addresses"))) { keys.Add(ToText(a)); }
                keys.Add(Text(o, "PrimaryAddress"));
                foreach (var k in keys)
                {
                    List<object> list;
                    if (k.Length > 0 && byKey.TryGetValue(k, out list)) { foreach (var m in list) { if (counted.Add(m)) { mine.Add(m); } } }
                }
                int series = 0, copies = 0, removed = 0, cancelled = 0, restored = 0, transferred = 0, failed = 0;
                foreach (var m in mine)
                {
                    if (Text(m, "Kind") == "Series") { series++; }
                    if (Text(m, "Status") == "Transferred") { transferred++; }
                    foreach (var c in Items(Prop(m, "Copies")))
                    {
                        if (Text(c, "EventId").Length > 0) { copies++; }
                        switch (Text(c, "Result"))
                        {
                            case "Removed": removed++; break;
                            case "Cancelled": cancelled++; break;
                            case "Restored": restored++; break;
                            case "Failed": failed++; break;
                        }
                    }
                }
                t.Rows.Add(new object[] { Text(o, "Input"), Text(o, "DisplayName"), Text(o, "PrimaryAddress"), Text(o, "State"), Text(o, "Detail"), mine.Count, series, copies, removed, cancelled, restored, transferred, failed });
            }
            return t;
        }

        /// <summary>
        /// One row per meeting of a transfer (Get-MclTransferRows): the old organizer and its state, the new one, the
        /// method, the new meeting and its invitation, what became of the old organizer's copy and of the old copies.
        /// </summary>
        public static Table TransferTable(object organizers, object meetings)
        {
            var t = new Table("MeetingId", "Subject", "StartText", "Kind", "Recurrence", "OldOrganizer", "OldOrganizerName", "OldOrganizerState", "OldOrganizerDetail",
                "NewOrganizer", "Method", "Status", "NewMeetingId", "NewMeeting", "NewMeetingDetail", "Invited", "Rooms", "OldOrganizerCopy",
                "OldCopiesRemoved", "OldCopiesFailed", "OldCopiesLeft", "Selected", "Notes");
            // The state of each organizer of the run (short, and the detail of the search), by every address it has.
            var state = new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase);
            foreach (var o in Items(organizers))
            {
                var s = Text(o, "State") == "Mailbox" ? (Text(o, "Account") == "Deleted" ? "Account deleted, mailbox present" : "Mailbox present")
                    : Text(o, "State") == "NoMailbox" ? "No mailbox" : "Not in the directory";
                var pair = new[] { s, Text(o, "Detail") };
                foreach (var a in Items(Prop(o, "Addresses"))) { var k = ToText(a); if (k.Length > 0) { state[k] = pair; } }
                var p = Text(o, "PrimaryAddress"); if (p.Length > 0) { state[p] = pair; }
                var i = Text(o, "Input"); if (i.Length > 0 && !state.ContainsKey(i)) { state[i] = pair; }
            }
            foreach (var m in Items(meetings))
            {
                // The meetings of the transfer only: an unticked meeting was not part of it.
                if (!Flag(m, "Selected")) { continue; }
                string newMeeting = "", newDetail = "", oldOrganizerCopy = "";
                int invited = 0, rooms = 0, removed = 0, failed = 0, left = 0;
                var native = Text(m, "TransferMethod") == "Native";
                foreach (var c in Items(Prop(m, "Copies")))
                {
                    var role = Text(c, "Role"); var result = Text(c, "Result");
                    if (role == "New organizer") { newMeeting = result; newDetail = Text(c, "Detail"); continue; }
                    if (Text(c, "EventId").Length == 0) { continue; }
                    if (role == "Organizer") { if (oldOrganizerCopy.Length == 0) { oldOrganizerCopy = result.Length > 0 ? result : "Present"; } }
                    else if (role == "Attendee" || role == "Room")
                    {
                        if (role == "Room") { rooms++; } else { invited++; }
                        if (result == "Removed" || result == "Already gone") { removed++; }
                        else if (result == "Failed" || result == "Not done") { failed++; }
                        // Moved by Exchange Online: the old copies are the meeting itself, updated in place.
                        else if (!native) { left++; }
                    }
                }
                if (oldOrganizerCopy.Length == 0) { oldOrganizerCopy = Text(m, "OrganizerCopy"); }
                var method = Text(m, "TransferMethod");
                var methodText = method == "Native" ? "Exchange Online" : method == "Recreate" ? "Re-created" : "";
                var key = Text(m, "OrganizerKey"); if (key.Length == 0) { key = Text(m, "Organizer"); }
                string[] orgState;
                if (!state.TryGetValue(key, out orgState) && !state.TryGetValue(Text(m, "Organizer"), out orgState)) { orgState = new[] { "", "" }; }
                var notes = new List<string>();
                foreach (var n in Items(Prop(m, "Notes"))) { notes.Add(ToText(n)); }
                t.Rows.Add(new object[] {
                    Text(m, "MeetingId"), Text(m, "Subject"), Text(m, "StartText"), Text(m, "Kind"), Text(m, "Recurrence"),
                    Text(m, "Organizer"), Text(m, "OrganizerName"), orgState[0], orgState[1],
                    Text(m, "NewOrganizer"), methodText, Text(m, "Status"), Text(m, "NewMeetingId"), newMeeting, newDetail,
                    invited, rooms, oldOrganizerCopy, removed, failed, left, Flag(m, "Selected"), notes.ToArray() });
            }
            return t;
        }

        /// <summary>A CSV file of a table: UTF-8 with BOM, the columns given (in that order), one line per row.</summary>
        public static void WriteTableCsv(Table table, string[] columns, string path, string delimiter)
        {
            var index = new int[columns.Length];
            for (int i = 0; i < columns.Length; i++) { index[i] = Array.IndexOf(table.Columns, columns[i]); }
            var sb = new StringBuilder();
            var cells = new string[columns.Length];
            for (int i = 0; i < columns.Length; i++) { cells[i] = CsvCell(columns[i], delimiter); }
            sb.AppendLine(string.Join(delimiter, cells));
            foreach (var r in table.Rows)
            {
                for (int i = 0; i < columns.Length; i++) { cells[i] = index[i] < 0 ? "" : CsvCell(r[index[i]], delimiter); }
                sb.AppendLine(string.Join(delimiter, cells));
            }
            File.WriteAllText(path, sb.ToString(), new UTF8Encoding(true));
        }

        /// <summary>The rows of a table as JSON objects (one property per column), HTML-safe (no &lt; &gt; &amp;).</summary>
        public static string TableJson(Table table)
        {
            var options = new JsonWriterOptions { Encoder = JavaScriptEncoder.Default };
            using (var stream = new MemoryStream())
            {
                using (var w = new Utf8JsonWriter(stream, options))
                {
                    w.WriteStartArray();
                    foreach (var r in table.Rows)
                    {
                        w.WriteStartObject();
                        for (int i = 0; i < table.Columns.Length; i++) { w.WritePropertyName(table.Columns[i]); WriteJson(w, r[i], 1); }
                        w.WriteEndObject();
                    }
                    w.WriteEndArray();
                }
                return Encoding.UTF8.GetString(stream.ToArray());
            }
        }

        /// <summary>A CSV cell: text starting with = + - @ (or tab, CR) prefixed with an apostrophe; quoted when needed (Format-MclCsvCell).</summary>
        public static string CsvCell(object value, string delimiter)
        {
            if (value == null) { return ""; }
            var v = Base(value);
            string text;
            if (v is bool) { text = (bool)v ? "True" : "False"; }
            else if (v is string)
            {
                text = (string)v;
                if (text.Length > 0 && "=+-@\t\r".IndexOf(text[0]) >= 0) { text = "'" + text; }
            }
            else if (v is IEnumerable && !(v is IDictionary))
            {
                var parts = new List<string>();
                foreach (var x in (IEnumerable)v) { parts.Add(ToText(x)); }
                text = string.Join(" | ", parts);
            }
            else { text = ToText(v); }
            if (text.Contains(delimiter) || text.Contains("\"") || text.IndexOf('\r') >= 0 || text.IndexOf('\n') >= 0) { text = "\"" + text.Replace("\"", "\"\"") + "\""; }
            return text;
        }

        /// <summary>A CSV file: UTF-8 with BOM, the columns given, one line per row (Write-MclCsv).</summary>
        public static void WriteCsv(object rows, string[] columns, string path, string delimiter)
        {
            var sb = new StringBuilder();
            var cells = new string[columns.Length];
            for (int i = 0; i < columns.Length; i++) { cells[i] = CsvCell(columns[i], delimiter); }
            sb.AppendLine(string.Join(delimiter, cells));
            foreach (var r in Items(rows))
            {
                for (int i = 0; i < columns.Length; i++) { cells[i] = CsvCell(Prop(r, columns[i]), delimiter); }
                sb.AppendLine(string.Join(delimiter, cells));
            }
            File.WriteAllText(path, sb.ToString(), new UTF8Encoding(true));
        }

        /// <summary>JSON safe inside a script block of the HTML report (no &lt; &gt; &amp;); rows of the report only.</summary>
        public static string ToJson(object value)
        {
            var options = new JsonWriterOptions { Encoder = JavaScriptEncoder.Default };
            using (var stream = new MemoryStream())
            {
                using (var writer = new Utf8JsonWriter(stream, options)) { WriteJson(writer, value, 0); }
                return Encoding.UTF8.GetString(stream.ToArray());
            }
        }

        static void WriteJson(Utf8JsonWriter w, object value, int depth)
        {
            if (value == null || depth > 12) { w.WriteNullValue(); return; }
            var v = Base(value);
            if (v is string) { w.WriteStringValue((string)v); return; }
            if (v is bool) { w.WriteBooleanValue((bool)v); return; }
            if (v is int || v is long || v is short || v is byte) { w.WriteNumberValue(Convert.ToInt64(v, Inv)); return; }
            if (v is double || v is float || v is decimal) { w.WriteNumberValue(Convert.ToDouble(v, Inv)); return; }
            if (v is DateTime) { w.WriteStringValue(((DateTime)v).ToString("o", Inv)); return; }
            var d = v as IDictionary;
            if (d != null)
            {
                w.WriteStartObject();
                foreach (DictionaryEntry e in d) { w.WritePropertyName(ToText(e.Key)); WriteJson(w, e.Value, depth + 1); }
                w.WriteEndObject();
                return;
            }
            if (v is PSCustomObject || (value is PSObject && ((PSObject)value).BaseObject is PSCustomObject))
            {
                w.WriteStartObject();
                foreach (var p in PSObject.AsPSObject(value).Properties) { w.WritePropertyName(p.Name); WriteJson(w, p.Value, depth + 1); }
                w.WriteEndObject();
                return;
            }
            var e2 = v as IEnumerable;
            if (e2 != null)
            {
                w.WriteStartArray();
                foreach (var x in e2) { WriteJson(w, x, depth + 1); }
                w.WriteEndArray();
                return;
            }
            w.WriteStringValue(ToText(v));
        }
    }

    /// <summary>Rows of the report: one array of values per row, in the order of the columns.</summary>
    public sealed class Table
    {
        public Table(params string[] columns) { Columns = columns; Rows = new List<object[]>(); }
        public string[] Columns { get; private set; }
        public List<object[]> Rows { get; private set; }
        public int Count { get { return Rows.Count; } }

        /// <summary>The rows as PowerShell objects, one property per column.</summary>
        public List<PSObject> ToObjects()
        {
            var list = new List<PSObject>();
            foreach (var r in Rows)
            {
                var o = new PSObject();
                for (int i = 0; i < Columns.Length; i++) { o.Properties.Add(new PSNoteProperty(Columns[i], r[i])); }
                list.Add(o);
            }
            return list;
        }
    }

    /// <summary>A calendar item kept by the search (Fast.MatchEvents).</summary>
    public sealed class EventMatch
    {
        public object Event { get; set; }
        public string Key { get; set; }
        public bool Own { get; set; }
        public string OrganizerKey { get; set; }
        public bool IsSeriesMaster { get; set; }
        public object Recurrence { get; set; }
    }

    // ---- window ------------------------------------------------------------------------------------------

    /// <summary>A meeting in the list of the window.</summary>
    public sealed class MeetingRow : INotifyPropertyChanged
    {
        bool _selected;
        bool _canTick;
        public bool Selected { get { return _selected; } set { if (_selected != value) { _selected = value; Notify("Selected"); } } }
        public bool CanTick { get { return _canTick; } set { if (_canTick != value) { _canTick = value; Notify("CanTick"); } } }
        public string Start { get; set; }
        public string Subject { get; set; }
        public string Who { get; set; }
        public string Kind { get; set; }
        public string OrganizerCopy { get; set; }
        public int Copies { get; set; }
        public int Rooms { get; set; }
        public string CopiesText { get; set; }
        public string Status { get; set; }
        public object Meeting { get; set; }
        public event PropertyChangedEventHandler PropertyChanged;
        void Notify(string name) { var h = PropertyChanged; if (h != null) { h(this, new PropertyChangedEventArgs(name)); } }
    }

    /// <summary>A copy of the meeting selected, in the window.</summary>
    public sealed class CopyRow
    {
        public string Mailbox { get; set; }
        public string Role { get; set; }
        public string Occurrence { get; set; }
        public string Via { get; set; }
        public string Result { get; set; }
        public string Detail { get; set; }
    }

    /// <summary>A list bound to the window, replaced in one go (one refresh instead of one per row).</summary>
    public class BulkCollection : ObservableCollection<object>
    {
        public void ReplaceAll(IEnumerable items)
        {
            Items.Clear();
            if (items != null) { foreach (var i in items) { Items.Add(i); } }
            OnPropertyChanged(new PropertyChangedEventArgs("Count"));
            OnPropertyChanged(new PropertyChangedEventArgs("Item[]"));
            OnCollectionChanged(new NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Reset));
        }
    }

    public static class GuiRows
    {
        /// <summary>The rows of the meetings of a result (Update-MclGuiRows).</summary>
        public static List<object> ForMeetings(object meetings, bool acted)
        {
            var rows = new List<object>();
            foreach (var m in Fast_Items(meetings))
            {
                var copies = Fast.RealCopies(m);
                int rooms = 0;
                foreach (var c in copies) { if (Fast.Text(c, "Role") == "Room") { rooms++; } }
                var name = Fast.Text(m, "OrganizerName");
                rows.Add(new MeetingRow
                {
                    Selected = Fast.Flag(m, "Selected"),
                    CanTick = !acted,
                    Start = Fast.Text(m, "StartText"),
                    Subject = Fast.Text(m, "Subject"),
                    Who = name.Length > 0 ? name : Fast.Text(m, "Organizer"),
                    Kind = Fast.Text(m, "Scope") == "Occurrences" ? string.Format(CultureInfo.InvariantCulture, "{0} occ.", Fast.Prop(m, "Occurrences")) : Fast.Text(m, "Kind"),
                    OrganizerCopy = Fast.Text(m, "OrganizerCopy"),
                    Copies = copies.Count,
                    Rooms = rooms,
                    CopiesText = rooms > 0 ? string.Format(CultureInfo.InvariantCulture, "{0} ({1} room{2})", copies.Count, rooms, rooms > 1 ? "s" : "") : copies.Count.ToString(CultureInfo.InvariantCulture),
                    Status = Fast.Text(m, "Status"),
                    Meeting = m
                });
            }
            return rows;
        }

        /// <summary>The rows of the copies of one meeting (Update-MclGuiCopies).</summary>
        public static List<object> ForCopies(object copies)
        {
            var rows = new List<object>();
            foreach (var c in Fast_Items(copies))
            {
                var result = Fast.Text(c, "Result");
                if (result.Length == 0 && Fast.Text(c, "EventId").Length > 0) { result = "Found"; }
                rows.Add(new CopyRow { Mailbox = Fast.Text(c, "Mailbox"), Role = Fast.Text(c, "Role"), Occurrence = Fast.Text(c, "Occurrence"), Via = Fast.Text(c, "Via"), Result = result, Detail = Fast.Text(c, "Detail") });
            }
            return rows;
        }

        /// <summary>How many organizers the meetings come from (the Organizer column is shown from two).</summary>
        public static int OrganizerCount(object meetings)
        {
            var keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var m in Fast_Items(meetings)) { var k = Fast.Text(m, "OrganizerKey"); keys.Add(k.Length > 0 ? k : Fast.Text(m, "Organizer")); }
            return keys.Count;
        }

        public static int CountSelected(IEnumerable rows)
        {
            int n = 0;
            if (rows != null) { foreach (var r in rows) { var row = r as MeetingRow; if (row != null && row.Selected) { n++; } } }
            return n;
        }

        /// <summary>Ticks or unticks every row (Tick all / Untick all).</summary>
        public static void SetSelected(IEnumerable rows, bool value)
        {
            if (rows == null) { return; }
            foreach (var r in rows) { var row = r as MeetingRow; if (row != null) { row.Selected = value; } }
        }

        /// <summary>The boxes ticked in the window, written back to the meetings of the result.</summary>
        public static void ApplySelection(IEnumerable rows)
        {
            if (rows == null) { return; }
            foreach (var r in rows)
            {
                var row = r as MeetingRow;
                if (row == null || row.Meeting == null) { continue; }
                var p = PSObject.AsPSObject(row.Meeting).Properties["Selected"];
                if (p != null) { p.Value = row.Selected; }
            }
        }

        static IEnumerable Fast_Items(object o)
        {
            if (o == null) { yield break; }
            var p = o as PSObject;
            var b = p != null && !(p.BaseObject is PSCustomObject) ? p.BaseObject : o;
            var e = b as IEnumerable;
            if (e == null || b is string) { yield return o; yield break; }
            foreach (var x in e) { if (x != null) { yield return x; } }
        }
    }
}
