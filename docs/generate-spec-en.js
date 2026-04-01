const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        Header, Footer, AlignmentType,
        HeadingLevel, BorderStyle, WidthType, ShadingType,
        PageNumber, PageBreak } = require('docx');
const fs = require('fs');

const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };
const cellMargins = { top: 60, bottom: 60, left: 100, right: 100 };
const headerShading = { fill: "2C3E50", type: ShadingType.CLEAR };
const altShading = { fill: "F8F9FA", type: ShadingType.CLEAR };
const noteShading = { fill: "FFF8E1", type: ShadingType.CLEAR };
const newShading = { fill: "E8F5E9", type: ShadingType.CLEAR };

function hCell(text, width) {
  return new TableCell({
    borders, width: { size: width, type: WidthType.DXA },
    shading: headerShading, margins: cellMargins,
    children: [new Paragraph({ children: [new TextRun({ text, bold: true, color: "FFFFFF", font: "Arial", size: 20 })] })]
  });
}
function cell(text, width, shading) {
  return new TableCell({
    borders, width: { size: width, type: WidthType.DXA },
    shading: shading || undefined, margins: cellMargins,
    children: [new Paragraph({ children: [new TextRun({ text, font: "Arial", size: 20 })] })]
  });
}
function mcell(runs, width, shading) {
  return new TableCell({
    borders, width: { size: width, type: WidthType.DXA },
    shading: shading || undefined, margins: cellMargins,
    children: [new Paragraph({ children: runs })]
  });
}

function heading1(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 360, after: 200 },
    children: [new TextRun({ text, bold: true, font: "Arial", size: 32 })] });
}
function heading2(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 280, after: 160 },
    children: [new TextRun({ text, bold: true, font: "Arial", size: 26 })] });
}
function heading3(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_3, spacing: { before: 200, after: 120 },
    children: [new TextRun({ text, bold: true, font: "Arial", size: 22 })] });
}
function para(text) {
  return new Paragraph({ spacing: { after: 120 },
    children: [new TextRun({ text, font: "Arial", size: 20 })] });
}
function boldPara(label, text) {
  return new Paragraph({ spacing: { after: 100 },
    children: [
      new TextRun({ text: label, bold: true, font: "Arial", size: 20 }),
      new TextRun({ text, font: "Arial", size: 20 })
    ]
  });
}
function bullet(text) {
  return new Paragraph({ spacing: { after: 80 },
    bullet: { level: 0 },
    children: [new TextRun({ text, font: "Arial", size: 20 })] });
}
function bullet2(text) {
  return new Paragraph({ spacing: { after: 60 },
    bullet: { level: 1 },
    children: [new TextRun({ text, font: "Arial", size: 20 })] });
}
function newTag() {
  return new TextRun({ text: " [NEW] ", bold: true, color: "2E7D32", font: "Arial", size: 20 });
}
function changedTag() {
  return new TextRun({ text: " [CHANGED] ", bold: true, color: "E65100", font: "Arial", size: 20 });
}
function notePara(text) {
  return new Paragraph({ spacing: { after: 120 },
    shading: noteShading,
    children: [new TextRun({ text: "NOTE: " + text, italics: true, font: "Arial", size: 20 })] });
}

function makeTable(headers, rows) {
  const widths = headers.map(() => Math.floor(9000 / headers.length));
  return new Table({
    rows: [
      new TableRow({ children: headers.map((h, i) => hCell(h, widths[i])) }),
      ...rows.map((row, ri) => new TableRow({
        children: row.map((c, i) => cell(c, widths[i], ri % 2 === 1 ? altShading : undefined))
      }))
    ]
  });
}

// ========== DOCUMENT CONTENT ==========
const children = [];

// COVER
children.push(
  new Paragraph({ spacing: { before: 3000 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "AIS Client Portal", bold: true, font: "Arial", size: 52 })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 },
    children: [new TextRun({ text: "Development Specification", font: "Arial", size: 36, color: "666666" })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 400 },
    children: [new TextRun({ text: "For IT Development Team", font: "Arial", size: 28, color: "999999" })] }),
  new Paragraph({ alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "Version 2.0 — April 2026", font: "Arial", size: 22, color: "999999" })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 },
    children: [new TextRun({ text: "Prepared by: Arches Consulting", font: "Arial", size: 22, color: "999999" })] }),
  new Paragraph({ children: [new PageBreak()] })
);

// ========== 1. OVERVIEW ==========
children.push(
  heading1("1. Overview"),
  para("This document defines the functional specifications for the AIS Client Portal redesign. It accompanies the HTML/CSS mockup files and describes the backend behavior, business rules, API requirements, and data models that the mockup alone cannot express."),
  para("The mockup repository is available at: github.com/yoshitakasakamoto-collab/client-portal-mockup"),
  heading2("1.1 What Changed vs. Current AIS"),
  para("The following table summarizes all changes from the current production AIS. Items marked [NEW] are entirely new features. Items marked [CHANGED] are modifications to existing behavior."),
  new Paragraph({ spacing: { after: 120 } }),
  makeTable(
    ["Area", "Change", "Type"],
    [
      ["Candidate Page — Layout", "Tabs (Expert Info / Activities / Availability) removed. All sections displayed in single scrollable view: Working History → Experience → Screening Answers → Availability → Comments & Activity Log", "CHANGED"],
      ["Candidate Page — Comments", "Comments section moved from hidden Activities tab to main page. Always visible. Activity Log shown in collapsible <details> element below comments.", "CHANGED"],
      ["Candidate Page — Excel Export", "Export List button (all experts) at top of left sidebar. Bulk action bar appears when checkboxes selected with Export Selected button.", "NEW"],
      ["Booking Modal — Timezone", "Timezone options changed to city-based neutral labels (e.g. 'UTC-05:00 New York, Toronto' instead of 'EST'). Browser auto-detects user TZ on page load (DST-aware).", "CHANGED"],
      ["Booking Modal — Duration", "Changed from text input to dropdown. Options: 30/45/60/75/90/105/120 min in 15-min increments.", "CHANGED"],
      ["Booking Modal — Attendees", "Removed 'If more than 1, please note in comments'. Default attendee pre-filled. Add button (+) to add more email fields.", "CHANGED"],
      ["Booking Modal — Meeting Method", "Dropdown with 3 options: Arches Zoom Link, Arches Zoom (Call-in), Custom Meeting Link. Conditional URL input field for Custom.", "NEW"],
      ["Booking Modal — Confirmation", "Confirmation dialog now shows: expert, date/time, duration, meeting method, attendees. Includes notice: 'This is not a final confirmation...'", "CHANGED"],
      ["Booking Modal — Availability Check", "If selected interview duration exceeds expert's available time window, show warning and block confirmation.", "NEW"],
      ["Email Thread Linking", "Project-level email thread linking via Gmail API. All client notifications sent to linked thread.", "NEW"],
      ["Notification Language", "Client notification language selectable per project (EN/JP).", "NEW"],
    ]
  ),
  new Paragraph({ children: [new PageBreak()] })
);

// ========== 2. PAGE SPECIFICATIONS ==========
children.push(
  heading1("2. Page Specifications"),

  // 2.1 Dashboard
  heading2("2.1 Dashboard (index.html)"),
  para("Main landing page after authentication. Displays all projects the client has access to."),
  heading3("2.1.1 Layout"),
  bullet("Sidebar navigation: Project Management, User Management"),
  bullet("Notification bell with unread badge count"),
  bullet("Language toggle (EN / JP)"),
  bullet("Two tabs: Ongoing Projects / Past Projects"),
  heading3("2.1.2 Project Table Columns"),
  makeTable(
    ["Column", "Data Type", "Notes"],
    [
      ["Project", "String", "Project code + title, links to project-general page"],
      ["Status", "Badge", "On going / Completed"],
      ["Start Date", "Date", "YYYY-MM-DD format"],
      ["CDD", "String", "Proposed/Approved counts (e.g. '5/3')"],
      ["IV", "String", "Waiting/Finished counts"],
      ["Total Price", "Currency", "Sum of all billable calls"],
      ["Arches Team", "String", "Team member names"],
      ["Geography", "String", "Target market/region"],
      ["Tags", "Array", "Project/expert keyword tags"],
    ]
  ),
  heading3("2.1.3 Backend Requirements"),
  bullet("GET /api/projects — Returns list of projects for authenticated user"),
  bullet("GET /api/notifications — Returns unread notifications with badge count"),
  bullet("PATCH /api/notifications/read-all — Mark all notifications as read"),
  new Paragraph({ spacing: { after: 120 } }),

  // 2.2 General
  heading2("2.2 Project General (project-general.html)"),
  para("Overview tab for a specific project. Displays statistics, team info, activity timeline, and project briefing."),
  heading3("2.2.1 Statistics Cards"),
  bullet("Experts Proposed: total (Pending / Approved / Declined breakdown)"),
  bullet("Interviews: total (Finished / Booked breakdown)"),
  bullet("Total Duration: hours (avg per call)"),
  heading3("2.2.2 Sections"),
  bullet("Segment Breakdown: table with Proposed/Approved/Calls Done per segment"),
  bullet("Team Activity: chronological list of actions by team members"),
  bullet("Project Details: team members, geography, tags, billing code"),
  bullet("Project Briefing: original inquiry email content + screening questions"),
  new Paragraph({ spacing: { after: 120 } }),

  // 2.3 Candidates
  heading2("2.3 Candidates (project-candidates.html)"),
  para("Split-view layout with expert list on left, expert detail on right. This is the most heavily modified page."),

  heading3("2.3.1 Left Sidebar — Expert List"),
  new Paragraph({ spacing: { after: 80 },
    children: [newTag(), new TextRun({ text: "Export List button at top — exports ALL experts to .xlsx", font: "Arial", size: 20 })] }),
  new Paragraph({ spacing: { after: 80 },
    children: [newTag(), new TextRun({ text: "Bulk action bar appears when 1+ checkboxes selected — shows count + Export Selected button", font: "Arial", size: 20 })] }),
  bullet("Experts grouped by segment (collapsible headers)"),
  bullet("Each expert shows: checkbox, ID, name, company, role, cost, status badge"),
  bullet("Click expert to load detail in right panel"),

  heading3("2.3.2 Right Panel — Expert Detail [CHANGED]"),
  para("Tabs have been completely removed. All information is displayed in a single scrollable page in this order:"),
  makeTable(
    ["Order", "Section", "Content", "Change"],
    [
      ["1", "Header", "Expert ID, name, status badge, updated date", "No change"],
      ["2", "Action Buttons", "Book Interview / Request Availability / Not Interested", "No change"],
      ["3", "Working History", "Table: Company, Position, From, To", "Was inside Expert Info tab → now top-level"],
      ["4", "Experience", "Paragraph text of professional experience", "Was inside Expert Info tab → now top-level"],
      ["5", "Screening Answers", "Q&A format", "Was inside Expert Info tab → now top-level"],
      ["6", "Availability", "Available time slots + Copy as text button", "Was separate Availability tab → now inline"],
      ["7", "Comments", "Comment list, input textarea, Send button. Info banner above.", "Was inside Activities tab → now always visible"],
      ["8", "Activity Log", "Collapsible (<details>) log of system events", "Was inside Activities > Log System sub-tab → now collapsible section"],
    ]
  ),
  notePara("The Availability quick-view that was above the tabs has been removed since Availability is now a full section."),

  heading3("2.3.3 Excel Export [NEW]"),
  para("Uses SheetJS (xlsx) library loaded via CDN."),
  makeTable(
    ["Function", "Trigger", "Output"],
    [
      ["exportAllExperts()", "Export List button", "expert_list.xlsx with ALL experts"],
      ["exportSelectedExperts()", "Export Selected button (bulk bar)", "expert_list_selected.xlsx with checked experts only"],
    ]
  ),
  boldPara("Export Columns: ", "ID, Name, Company, Position, Status, Cost, Availability"),
  heading3("2.3.4 Backend Requirements"),
  bullet("GET /api/projects/{id}/experts — Expert list with all fields"),
  bullet("GET /api/projects/{id}/experts/{expertId} — Single expert detail"),
  bullet("POST /api/projects/{id}/experts/{expertId}/comments — Add comment"),
  bullet("GET /api/projects/{id}/experts/{expertId}/activity-log — Activity log entries"),
  bullet("GET /api/projects/{id}/experts/export — Server-side Excel generation (optional, can be client-side)"),
  new Paragraph({ children: [new PageBreak()] }),

  // 2.4 Booking Modal
  heading2("2.4 Booking Modal (within Candidates page)"),
  para("Opened when user clicks 'Book for Interview' button on an expert with availability slots."),

  heading3("2.4.1 Fields"),
  makeTable(
    ["Field", "Type", "Required", "Default", "Change"],
    [
      ["Your Timezone", "Dropdown (18 options)", "Yes", "Auto-detected from browser", "CHANGED — city-based labels, auto-detect"],
      ["Interview Duration", "Dropdown", "Yes", "60 min", "CHANGED — was text input, now dropdown (30-120 min, 15-min steps)"],
      ["2-Week Calendar", "Interactive grid", "Yes", "(none)", "No change — 15-min slots, green=available, orange=selected, red=booked"],
      ["Preferred Date & Time", "Readonly text", "Yes", "(from calendar)", "No change"],
      ["Attendees", "Email fields", "Yes", "Logged-in user email", "CHANGED — removed helper text, added + button for additional attendees"],
      ["Meeting Method", "Dropdown", "Yes", "Arches Zoom Link", "NEW"],
      ["Interpretation", "Checkbox", "No", "unchecked", "No change"],
      ["Additional Notes", "Textarea", "No", "(empty)", "No change"],
    ]
  ),

  heading3("2.4.2 Timezone Options"),
  para("18 timezone options using neutral city-based labels. No country-specific abbreviations (no JST, EST, etc.)."),
  makeTable(
    ["Label", "UTC Offset"],
    [
      ["(UTC-08:00) Los Angeles, Vancouver", "-8"],
      ["(UTC-07:00) Denver, Phoenix", "-7"],
      ["(UTC-06:00) Chicago, Mexico City", "-6"],
      ["(UTC-05:00) New York, Toronto", "-5"],
      ["(UTC-03:00) São Paulo, Buenos Aires", "-3"],
      ["(UTC+00:00) London, Dublin, Lisbon", "0"],
      ["(UTC+01:00) Paris, Berlin, Amsterdam", "1"],
      ["(UTC+02:00) Cairo, Helsinki, Bucharest", "2"],
      ["(UTC+03:00) Moscow, Istanbul, Riyadh", "3"],
      ["(UTC+04:00) Dubai, Baku", "4"],
      ["(UTC+05:00) Karachi, Tashkent", "5"],
      ["(UTC+05:30) Mumbai, New Delhi, Colombo", "5.5"],
      ["(UTC+06:00) Dhaka, Almaty", "6"],
      ["(UTC+07:00) Bangkok, Jakarta, Hanoi", "7"],
      ["(UTC+08:00) Singapore, Hong Kong, Taipei, Perth", "8"],
      ["(UTC+09:00) Tokyo, Seoul", "9"],
      ["(UTC+10:00) Sydney, Melbourne", "10"],
      ["(UTC+12:00) Auckland, Wellington", "12"],
    ]
  ),
  boldPara("Auto-detection: ", "On page load, browser's Intl.DateTimeFormat().resolvedOptions().timeZone is read. UTC offset is calculated from new Date().getTimezoneOffset(). Closest matching dropdown option is auto-selected. DST is handled automatically because the offset is calculated at runtime."),

  heading3("2.4.3 Meeting Method Options [NEW]"),
  makeTable(
    ["Option", "Value", "Behavior"],
    [
      ["Arches Zoom Link", "arches_zoom", "Arches provides a Zoom meeting link (default)"],
      ["Arches Zoom (Call-in)", "arches_callin", "Arches provides Zoom dial-in phone number"],
      ["Custom Meeting Link", "client_link", "Shows URL input field. Client pastes their own Zoom/Teams/Meet link."],
    ]
  ),

  heading3("2.4.4 Availability Validation [NEW]"),
  para("When user selects a time slot and clicks Confirm Booking:"),
  bullet("System checks if the selected interview duration fits entirely within the expert's availability window"),
  bullet("Example: Expert available 10:00-10:15 only. User selects 10:00 start with 60-min duration → BLOCKED"),
  bullet("Warning message displayed, Confirm Booking button disabled"),
  boldPara("Rule: ", "interview_end_time <= expert_availability_block_end_time. Checked in 15-minute increments."),

  heading3("2.4.5 Confirmation Dialog"),
  para("After clicking Confirm Booking and passing validation:"),
  makeTable(
    ["Field", "Value"],
    [
      ["Expert", "EX-XXXXXX - Name"],
      ["Preferred Date & Time", "Mar 20 (Fri), 10:00 AM - 11:00 AM UTC-05:00"],
      ["Interview Duration", "60 min"],
      ["Meeting Method", "Arches Zoom Link / Call-in / Client link URL"],
      ["Attendees", "List of all attendee emails"],
    ]
  ),
  para("Warning notice (yellow box):"),
  para("'This is not a final confirmation. A scheduling request will be sent to the expert for the date/time shown above. Once the expert accepts, a calendar invitation will be issued.'"),

  heading3("2.4.6 Backend Requirements"),
  bullet("POST /api/projects/{id}/bookings — Create booking request"),
  boldPara("  Request body: ", "{ expertId, dateTime, duration, timezone, meetingMethod, clientMeetingLink?, attendees[], interpretation, notes }"),
  bullet("Booking status flow: Requested → Confirmed by Expert → Calendar Issued → Conducted"),
  new Paragraph({ children: [new PageBreak()] }),

  // 2.5 Interview
  heading2("2.5 Interview Management (project-interview.html)"),
  para("Displays all interviews for the project, organized by segment."),
  heading3("2.5.1 Statistics"),
  bullet("Booked / Conducted / Canceled counts"),
  heading3("2.5.2 Interview Table Columns"),
  makeTable(
    ["Column", "Data", "Notes"],
    [
      ["Expert", "ID + Name (link to candidate page)", "Clicking opens candidate detail"],
      ["Status", "Badge", "Booked / Conducted / Canceled"],
      ["Date & Time", "Date + time range + TZ", ""],
      ["Duration", "Minutes", "'--' if not yet conducted"],
      ["Cost", "Currency", "'(est.)' if not yet confirmed"],
      ["Recording", "Download button", "Available after interview"],
      ["Transcript", "Download button", "Available after processing"],
      ["Summary", "Download button", "AI-generated summary"],
      ["Actions", "Rate / Cancel buttons", "Rate opens feedback modal"],
    ]
  ),
  heading3("2.5.3 Cancel Modal"),
  bullet("Required: Reason dropdown (Schedule conflict / No longer needed / Project scope changed / Client-side delay / Expert unavailable / Other)"),
  bullet("Optional: Additional details textarea"),
  heading3("2.5.4 Feedback Modal"),
  bullet("5-star rating (clickable)"),
  bullet("Free-text feedback textarea"),
  new Paragraph({ spacing: { after: 120 } }),

  // 2.6 Billing
  heading2("2.6 Billing (project-billing.html)"),
  para("Displays cost breakdown and invoice management."),
  heading3("2.6.1 Summary Cards"),
  bullet("Total Billed (with call count)"),
  bullet("Discount Applied"),
  bullet("Billing Code"),
  heading3("2.6.2 Cost Breakdown Table"),
  makeTable(
    ["Column", "Data"],
    [
      ["Expert", "Name"],
      ["Date", "Interview date"],
      ["Type", "Initial Call / Follow-up / Interpretation / etc."],
      ["Duration", "Minutes"],
      ["Rate", "$/hour or flat rate"],
      ["Discount", "Amount or '--'"],
      ["Cost", "Final cost"],
      ["Status", "Invoiced / Pending / Paid"],
    ]
  ),
  heading3("2.6.3 Invoice Request"),
  bullet("Button opens modal to request invoice"),
  bullet("Optional notes field"),
  bullet("POST /api/projects/{id}/invoice-request"),
  new Paragraph({ children: [new PageBreak()] }),

  // ========== 3. EMAIL THREAD LINKING ==========
  heading1("3. Email Thread Linking [NEW]"),
  para("This is a new feature that does not exist in current AIS. It allows Arches operators to link a Gmail thread to a project, so that all client notifications are sent to that existing thread instead of creating new separate emails."),

  heading2("3.1 Problem Statement"),
  para("Currently, each system notification creates a new email thread. Clients receive multiple disconnected emails and don't know which thread to use for communication. This causes confusion and fragmented conversations."),

  heading2("3.2 Solution"),
  para("A project can be linked to a specific email thread. Once linked, all client-facing notifications are sent as replies to that thread, keeping everything in one conversation."),

  heading2("3.3 User Flow"),
  bullet("1. Arches operator opens the project General page"),
  bullet("2. Clicks 'Link Email Thread' button"),
  bullet("3. A search box appears"),
  bullet("4. Operator types a keyword (e.g. client name, subject) to search"),
  bullet("5. System searches the operator's Gmail inbox via Gmail API"),
  bullet("6. Results show recent matching threads (subject line + date + participants)"),
  bullet("7. Operator selects the correct thread"),
  bullet("8. System stores the thread's Message-ID in the database, linked to the project"),
  bullet("9. Done. All future notifications go to that thread."),

  heading2("3.4 Technical Implementation"),
  heading3("3.4.1 Gmail API Integration"),
  bullet("OAuth2 authentication for Arches operators' Gmail accounts"),
  bullet("Scope required: gmail.readonly (for searching threads)"),
  bullet("Endpoint: GET /gmail/v1/users/me/messages?q={searchQuery}"),
  bullet("Search returns threads matching the query from the operator's inbox"),
  bullet("Display: Subject line, date, participant names/emails"),

  heading3("3.4.2 Data Model"),
  makeTable(
    ["Field", "Type", "Description"],
    [
      ["project_id", "String", "Project identifier (e.g. A109919)"],
      ["gmail_thread_id", "String", "Gmail thread ID for API operations"],
      ["message_id", "String", "RFC 2822 Message-ID header of the anchor message (e.g. <abc123@mail.gmail.com>)"],
      ["linked_by", "String", "Operator who linked the thread"],
      ["linked_at", "DateTime", "When the thread was linked"],
    ]
  ),

  heading3("3.4.3 Sending Notifications to Linked Thread"),
  para("When sending a notification email for a project that has a linked thread:"),
  bullet("Set In-Reply-To header to the stored message_id"),
  bullet("Set References header to the stored message_id"),
  bullet("Keep subject line consistent: '[PROJECT_CODE] notification subject'"),
  para("This causes both Gmail and Outlook to group the notification into the existing thread."),

  heading3("3.4.4 Gmail vs Outlook Compatibility"),
  makeTable(
    ["Client", "Threading Method", "Notes"],
    [
      ["Gmail", "References header + subject similarity", "Subject should contain project code for reliable threading"],
      ["Outlook", "In-Reply-To + References headers", "Subject is less important; headers are primary"],
    ]
  ),
  boldPara("Recommendation: ", "Always prefix subject with project code: [A109919] New expert proposed: Luci Dao"),

  heading3("3.4.5 Thread Re-linking"),
  bullet("Operators can change the linked thread at any time"),
  bullet("Previous notifications remain in the old thread (no migration)"),
  bullet("New notifications go to the newly linked thread from that point forward"),

  heading2("3.5 Notification Types Sent to Thread"),
  makeTable(
    ["Notification", "Trigger", "Send Mode"],
    [
      ["Recording / Transcript / Summary ready", "Processing complete", "Automatic — sent immediately upon completion"],
      ["New expert proposed", "Arches operator promotes expert", "Manual — operator selects experts and clicks 'Promote to Client'"],
    ]
  ),
  notePara("Existing notification sending functionality in AIS is preserved. This feature only changes the DESTINATION (linked thread) instead of creating a new email."),

  heading2("3.6 Notification Language"),
  para("Per-project setting to choose notification email language (EN or JP). Stored in project settings. Email templates are rendered in the selected language."),

  heading2("3.7 Backend Requirements"),
  bullet("POST /api/projects/{id}/link-thread — Store thread link { gmail_thread_id, message_id }"),
  bullet("DELETE /api/projects/{id}/link-thread — Remove thread link"),
  bullet("GET /api/projects/{id}/link-thread — Get current linked thread info"),
  bullet("GET /api/gmail/search?q={query} — Proxy to Gmail API for thread search (requires OAuth)"),
  bullet("PATCH /api/projects/{id}/settings — Update notification language { notification_lang: 'en'|'jp' }"),
  new Paragraph({ children: [new PageBreak()] }),

  // ========== 4. CROSS-CUTTING CONCERNS ==========
  heading1("4. Cross-Cutting Concerns"),

  heading2("4.1 Authentication"),
  para("Current mockup uses a simple password gate (auth.js). Production should use proper OAuth2/SSO authentication."),
  bullet("Session management with token refresh"),
  bullet("Role-based access: Client users vs Arches operators"),
  bullet("Arches operators need Gmail OAuth scope for thread linking feature"),

  heading2("4.2 Internationalization (i18n)"),
  para("The portal supports EN and JP languages. All UI text uses data-lang attributes mapped to translation keys in lang.js."),
  bullet("213 translation keys currently defined"),
  bullet("Language preference stored per user session"),
  bullet("Dynamic content (rendered via JavaScript) calls applyLanguage() after DOM update"),
  boldPara("Implementation note: ", "The applyLang() function stores original EN text in data-lang-en attribute on first run. When switching to EN, it restores from this attribute (not from the key name)."),

  heading2("4.3 Responsive Design"),
  para("Modal widths use max-width with vw fallback (e.g. 860px / 95vw). Split-view layout should stack vertically on mobile."),

  heading2("4.4 Timezone Handling"),
  bullet("All times stored in UTC in the database"),
  bullet("Displayed in user's selected timezone"),
  bullet("Expert availability stored with original timezone info"),
  bullet("Conversion happens client-side using UTC offset"),
  bullet("DST handled by browser's Intl API at runtime"),
  new Paragraph({ children: [new PageBreak()] }),

  // ========== 5. MOCKUP FILE REFERENCE ==========
  heading1("5. Mockup File Reference"),
  para("The HTML mockup files serve as the visual specification. Below is the file inventory:"),
  makeTable(
    ["File", "Purpose"],
    [
      ["index.html", "Project dashboard — project list, notifications"],
      ["project-general.html", "Project overview — stats, team, briefing (A109919)"],
      ["project-general-a110031.html", "Project overview (A110031 — JP project)"],
      ["project-candidates.html", "Expert list + detail panel + booking modal (A109919)"],
      ["project-candidates-a110031.html", "Expert list + detail panel + booking modal (A110031)"],
      ["project-interview.html", "Interview management — recordings, cancel, feedback (A109919)"],
      ["project-interview-a110031.html", "Interview management (A110031)"],
      ["project-billing.html", "Billing — cost breakdown, invoice request (A109919)"],
      ["project-billing-a110031.html", "Billing (A110031)"],
      ["lang.js", "Language switching system (213 EN/JP translation keys)"],
      ["auth.js", "Password gate (mockup only — replace with OAuth in production)"],
    ]
  ),
  notePara("A109919 represents a US-market project. A110031 represents a Japan-market project. Both use identical page structure but different data and some UI variations (e.g. screening answers format)."),
);

// Build document
const doc = new Document({
  sections: [{
    properties: {
      page: { margin: { top: 1000, bottom: 1000, left: 1200, right: 1200 } }
    },
    headers: {
      default: new Header({
        children: [new Paragraph({ alignment: AlignmentType.RIGHT,
          children: [new TextRun({ text: "AIS Client Portal — Development Specification", font: "Arial", size: 16, color: "999999" })] })]
      })
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({ alignment: AlignmentType.CENTER,
          children: [
            new TextRun({ text: "Page ", font: "Arial", size: 16, color: "999999" }),
            new TextRun({ children: [PageNumber.CURRENT], font: "Arial", size: 16, color: "999999" }),
          ] })]
      })
    },
    children
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(__dirname + '/Client_Portal_Dev_Spec_EN.docx', buf);
  console.log('EN spec generated.');
});
