const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        Header, Footer, AlignmentType, LevelFormat,
        HeadingLevel, BorderStyle, WidthType, ShadingType,
        PageNumber, PageBreak, TableOfContents } = require('docx');
const fs = require('fs');

const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };
const cellMargins = { top: 60, bottom: 60, left: 100, right: 100 };
const headerShading = { fill: "2C3E50", type: ShadingType.CLEAR };
const altShading = { fill: "F8F9FA", type: ShadingType.CLEAR };

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
function para(text, opts = {}) {
  return new Paragraph({ spacing: { after: 120 }, ...opts,
    children: [new TextRun({ text, font: "Arial", size: 20, ...opts.run })] });
}
function boldPara(label, text) {
  return new Paragraph({ spacing: { after: 100 },
    children: [
      new TextRun({ text: label, bold: true, font: "Arial", size: 20 }),
      new TextRun({ text, font: "Arial", size: 20 })
    ]
  });
}
function bulletItem(text, ref) {
  return new Paragraph({ numbering: { reference: ref, level: 0 }, spacing: { after: 60 },
    children: [new TextRun({ text, font: "Arial", size: 20 })] });
}
function numberItem(text, ref) {
  return new Paragraph({ numbering: { reference: ref, level: 0 }, spacing: { after: 60 },
    children: [new TextRun({ text, font: "Arial", size: 20 })] });
}

const doc = new Document({
  styles: {
    default: { document: { run: { font: "Arial", size: 20 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "Arial", color: "2C3E50" },
        paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Arial", color: "E67E22" },
        paragraph: { spacing: { before: 280, after: 160 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 22, bold: true, font: "Arial", color: "34495E" },
        paragraph: { spacing: { before: 200, after: 120 }, outlineLevel: 2 } },
    ]
  },
  numbering: {
    config: [
      { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "numbers", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "numbers2", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "numbers3", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "numbers4", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "numbers5", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "numbers6", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "numbers7", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "bullets2", levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "bullets3", levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "bullets4", levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    ]
  },
  sections: [
    // ===== TITLE PAGE =====
    {
      properties: {
        page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } }
      },
      children: [
        new Paragraph({ spacing: { before: 3000 } }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 },
          children: [new TextRun({ text: "Client Portal", font: "Arial", size: 56, bold: true, color: "2C3E50" })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 100 },
          children: [new TextRun({ text: "Development Specification", font: "Arial", size: 40, color: "E67E22" })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 600 },
          border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "E67E22", space: 1 } },
          children: [new TextRun({ text: " ", font: "Arial", size: 20 })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 100 },
          children: [new TextRun({ text: "Arches Consulting", font: "Arial", size: 28, color: "7F8C8D" })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 100 },
          children: [new TextRun({ text: "Version 1.0 | March 2026", font: "Arial", size: 22, color: "7F8C8D" })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 100 },
          children: [new TextRun({ text: "CONFIDENTIAL", font: "Arial", size: 20, bold: true, color: "E74C3C" })] }),
      ]
    },

    // ===== TOC + MAIN CONTENT =====
    {
      properties: {
        page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } }
      },
      headers: {
        default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT,
          children: [new TextRun({ text: "Client Portal - Development Spec", font: "Arial", size: 16, color: "999999" })] })] })
      },
      footers: {
        default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "Page ", font: "Arial", size: 16 }), new TextRun({ children: [PageNumber.CURRENT], font: "Arial", size: 16 })] })] })
      },
      children: [
        new TableOfContents("Table of Contents", { hyperlink: true, headingStyleRange: "1-3" }),
        new Paragraph({ children: [new PageBreak()] }),

        // ===== 1. SYSTEM OVERVIEW =====
        heading1("1. System Overview"),
        para("The Client Portal is a module within the AIS (Arches Intelligence System) that provides clients with self-service access to expert management, interview scheduling, and billing. This document specifies the functional and technical requirements for the portal redesign."),

        heading2("1.1 Architecture Context"),
        boldPara("Master System: ", "AIS (Arches Intelligence System) - the internal expert database and project management system."),
        boldPara("Client Portal: ", "A client-facing module within AIS. All data originates from AIS and is surfaced to clients through the portal."),
        boldPara("Data Flow: ", "AIS (master) -> Client Portal (read + limited write). Client actions (booking, comments, decline) are written back to AIS and trigger notifications."),

        heading2("1.2 Notification Architecture"),
        para("Two notification channels operate in parallel:"),
        new Table({
          width: { size: 9360, type: WidthType.DXA }, columnWidths: [2000, 3680, 3680],
          rows: [
            new TableRow({ children: [hCell("Channel", 2000), hCell("Client-side", 3680), hCell("Arches-side", 3680)] }),
            new TableRow({ children: [
              cell("Portal", 2000), cell("In-app notification bell (badge count + dropdown list)", 3680), cell("N/A - Arches uses AIS internal dashboard", 3680)] }),
            new TableRow({ children: [
              cell("Email", 2000, altShading), cell("Email sent to project members on key events (new expert, IV confirmed, transcript ready)", 3680, altShading),
              cell("N/A", 3680, altShading)] }),
            new TableRow({ children: [
              cell("Slack", 2000), cell("N/A", 3680),
              cell("Each project has a dedicated Slack channel. A Slack email address is registered per project. All client actions trigger emails to this address, posting to the Slack channel.", 3680)] }),
          ]
        }),

        heading2("1.3 User Roles (Initial Release)"),
        para("All client users have identical permissions. Role-based access control (Admin/Viewer/Booker) is out of scope for v1 but should be designed for future extensibility."),
        para("Arches staff access the system through the AIS internal dashboard, not the client portal."),

        new Paragraph({ children: [new PageBreak()] }),

        // ===== 2. DATA MODELS =====
        heading1("2. Data Models"),

        heading2("2.1 Entity Relationship Overview"),
        para("The following entities are managed in AIS and surfaced to the client portal:"),

        new Table({
          width: { size: 9360, type: WidthType.DXA }, columnWidths: [2000, 4360, 3000],
          rows: [
            new TableRow({ children: [hCell("Entity", 2000), hCell("Description", 4360), hCell("Source", 3000)] }),
            new TableRow({ children: [cell("Project", 2000), cell("A client engagement with segments, team members, and briefing", 4360), cell("AIS (master)", 3000)] }),
            new TableRow({ children: [cell("Expert", 2000, altShading), cell("A subject-matter expert with profile, experience, availability", 4360, altShading), cell("AIS (master)", 3000, altShading)] }),
            new TableRow({ children: [cell("Interview", 2000), cell("A scheduled/completed call between client and expert", 4360), cell("AIS + Portal (write-back)", 3000)] }),
            new TableRow({ children: [cell("Billing Item", 2000, altShading), cell("A cost line item (interview, interpretation, follow-up)", 4360, altShading), cell("AIS (master, existing logic)", 3000, altShading)] }),
            new TableRow({ children: [cell("Comment", 2000), cell("A threaded comment on an expert by client or Arches staff", 4360), cell("Portal (write) -> AIS", 3000)] }),
            new TableRow({ children: [cell("Notification", 2000, altShading), cell("An in-app + email alert for key events", 4360, altShading), cell("AIS generates", 3000, altShading)] }),
          ]
        }),

        heading2("2.2 Expert Status Flow"),
        para("The expert lifecycle in AIS follows this progression:"),
        new Table({
          width: { size: 9360, type: WidthType.DXA }, columnWidths: [2000, 4680, 2680],
          rows: [
            new TableRow({ children: [hCell("Status", 2000), hCell("Description", 4680), hCell("Visible in Portal?", 2680)] }),
            new TableRow({ children: [cell("Prospect", 2000), cell("Arches is reaching out to potential candidates. Not yet vetted.", 4680), cell("NO - internal only", 2680)] }),
            new TableRow({ children: [cell("Proposed", 2000, altShading), cell("Arches has vetted and proposed the expert to the client.", 4680, altShading), cell("YES - first visible state", 2680, altShading)] }),
            new TableRow({ children: [cell("Approved", 2000), cell("Client has reviewed and approved the expert for interview.", 4680), cell("YES", 2680)] }),
            new TableRow({ children: [cell("Declined", 2000, altShading), cell("Client has declined the expert (with reason).", 4680, altShading), cell("YES (grayed out)", 2680, altShading)] }),
            new TableRow({ children: [cell("Interview", 2000), cell("Interview is booked or in progress.", 4680), cell("YES (shown in IV page)", 2680)] }),
            new TableRow({ children: [cell("Billing", 2000, altShading), cell("Interview completed, billing generated.", 4680, altShading), cell("YES (shown in billing)", 2680, altShading)] }),
          ]
        }),

        heading2("2.3 Expert Data Model"),
        new Table({
          width: { size: 9360, type: WidthType.DXA }, columnWidths: [2200, 1500, 5660],
          rows: [
            new TableRow({ children: [hCell("Field", 2200), hCell("Type", 1500), hCell("Notes", 5660)] }),
            new TableRow({ children: [cell("id", 2200), cell("String", 1500), cell("Format: EX-XXXXXXX (7 digits)", 5660)] }),
            new TableRow({ children: [cell("name", 2200, altShading), cell("String", 1500, altShading), cell("Full name of the expert", 5660, altShading)] }),
            new TableRow({ children: [cell("company", 2200), cell("String", 1500), cell("Current employer", 5660)] }),
            new TableRow({ children: [cell("role", 2200, altShading), cell("String", 1500, altShading), cell("Current job title", 5660, altShading)] }),
            new TableRow({ children: [cell("status", 2200), cell("Enum", 1500), cell("Proposed | Approved | Declined | Interview | Billing", 5660)] }),
            new TableRow({ children: [cell("cost", 2200, altShading), cell("Number", 1500, altShading), cell("Per-call rate in USD (set per expert, per project)", 5660, altShading)] }),
            new TableRow({ children: [cell("segment", 2200), cell("String", 1500), cell("Which project segment this expert belongs to", 5660)] }),
            new TableRow({ children: [cell("experience[]", 2200, altShading), cell("String[]", 1500, altShading), cell("Array of paragraph-length experience descriptions", 5660, altShading)] }),
            new TableRow({ children: [cell("history[]", 2200), cell("Object[]", 1500), cell("{company, position, from, to} - chronological work history", 5660)] }),
            new TableRow({ children: [cell("screening[]", 2200, altShading), cell("Object[]", 1500, altShading), cell("{question, answer} - screening Q&A pairs", 5660, altShading)] }),
            new TableRow({ children: [cell("availability[]", 2200), cell("Object[]", 1500), cell("{date, startTime, endTime, timezone} - available time blocks. Entered by Arches staff in AIS.", 5660)] }),
            new TableRow({ children: [cell("comments[]", 2200, altShading), cell("Object[]", 1500, altShading), cell("{author, org, text, timestamp, replies[]} - threaded comments", 5660, altShading)] }),
            new TableRow({ children: [cell("logs[]", 2200), cell("Object[]", 1500), cell("{action, timestamp} - activity log (system-generated)", 5660)] }),
          ]
        }),

        new Paragraph({ children: [new PageBreak()] }),

        // ===== 3. PAGE SPECIFICATIONS =====
        heading1("3. Page Specifications"),

        // --- 3.1 Project Dashboard ---
        heading2("3.1 Project Dashboard (index.html)"),
        para("Lists all projects the client has access to, divided into Ongoing and Past tabs."),
        heading3("3.1.1 Data Requirements"),
        bulletItem("Fetch all projects associated with the authenticated client organization", "bullets"),
        bulletItem("Each project shows: name, status, start date, CDD stats (proposed/approved counts), IV stats (waiting/finished counts), total price, Arches team, geography, tags", "bullets"),
        bulletItem("Notification count badge on bell icon (unread count)", "bullets"),
        heading3("3.1.2 API Endpoints Needed"),
        new Table({
          width: { size: 9360, type: WidthType.DXA }, columnWidths: [1500, 3180, 4680],
          rows: [
            new TableRow({ children: [hCell("Method", 1500), hCell("Endpoint", 3180), hCell("Description", 4680)] }),
            new TableRow({ children: [cell("GET", 1500), cell("/api/projects", 3180), cell("List all projects for current client org", 4680)] }),
            new TableRow({ children: [cell("GET", 1500, altShading), cell("/api/notifications", 3180, altShading), cell("Unread notifications for current user", 4680, altShading)] }),
            new TableRow({ children: [cell("PATCH", 1500), cell("/api/notifications/read-all", 3180), cell("Mark all notifications as read", 4680)] }),
          ]
        }),

        // --- 3.2 General ---
        heading2("3.2 Project General Page"),
        para("High-level dashboard showing project summary, segment breakdown, team activity, and project briefing."),
        heading3("3.2.1 Key Sections"),
        bulletItem("Summary cards: Experts Proposed count, Interviews count, Total Duration", "bullets2"),
        bulletItem("Segment breakdown table: proposed/approved/calls-done per segment", "bullets2"),
        bulletItem("Team activity feed: timeline of recent actions (approve, decline, book, complete)", "bullets2"),
        bulletItem("Project details: team members, geography, tags, billing code", "bullets2"),
        bulletItem("Project briefing: the original inquiry context and screening questions", "bullets2"),
        heading3("3.2.2 API Endpoints Needed"),
        new Table({
          width: { size: 9360, type: WidthType.DXA }, columnWidths: [1500, 3680, 4180],
          rows: [
            new TableRow({ children: [hCell("Method", 1500), hCell("Endpoint", 3680), hCell("Description", 4180)] }),
            new TableRow({ children: [cell("GET", 1500), cell("/api/projects/:id", 3680), cell("Full project details including segments, team, briefing", 4180)] }),
            new TableRow({ children: [cell("GET", 1500, altShading), cell("/api/projects/:id/activity", 3680, altShading), cell("Activity feed for the project", 4180, altShading)] }),
          ]
        }),

        new Paragraph({ children: [new PageBreak()] }),

        // --- 3.3 Candidates ---
        heading2("3.3 Candidates Page (Expert Management)"),
        para("Two-panel split view: left panel shows expert list grouped by segment, right panel shows selected expert detail."),

        heading3("3.3.1 Left Panel - Expert List"),
        bulletItem("Grouped by segment (collapsible sections)", "bullets3"),
        bulletItem("Each expert shows: ID, name, company, role, cost, status badge", "bullets3"),
        bulletItem("Checkboxes for bulk selection", "bullets3"),
        bulletItem("Export All button (top of list) - exports all experts to Excel (.xlsx)", "bullets3"),
        bulletItem("Bulk action bar (appears when checkboxes selected) - Export Selected", "bullets3"),

        heading3("3.3.2 Right Panel - Expert Detail (Single Scrollable Page)"),
        para("No tabs. All sections displayed vertically in this order:"),
        numberItem("Availability - available time slots with copy-to-clipboard button", "numbers"),
        numberItem("Working History - table (Company, Position, From, To)", "numbers"),
        numberItem("Experience - paragraph descriptions of expertise", "numbers"),
        numberItem("Screening Answers - Q&A pairs from expert vetting", "numbers"),
        numberItem("Comments & Activity - threaded comments with reply support + collapsible activity log", "numbers"),

        heading3("3.3.3 Action Buttons"),
        new Table({
          width: { size: 9360, type: WidthType.DXA }, columnWidths: [2500, 3430, 3430],
          rows: [
            new TableRow({ children: [hCell("Button", 2500), hCell("Condition", 3430), hCell("Action", 3430)] }),
            new TableRow({ children: [cell("Book for Interview", 2500), cell("Expert has availability slots", 3430), cell("Opens booking modal", 3430)] }),
            new TableRow({ children: [cell("Request Availability", 2500, altShading), cell("Expert has no availability", 3430, altShading), cell("Sends request to Arches (Slack notification)", 3430, altShading)] }),
            new TableRow({ children: [cell("Not Interested", 2500), cell("Expert not yet declined", 3430), cell("Opens decline modal with reason dropdown", 3430)] }),
          ]
        }),

        heading3("3.3.4 Decline Flow"),
        bulletItem("Decline reason (required): Not relevant, Too expensive, Already have coverage, Competitor, Other", "bullets4"),
        bulletItem("Additional comments (optional): free-text", "bullets4"),
        bulletItem("On submit: expert status changes to Declined, notification sent to Arches Slack channel", "bullets4"),

        heading3("3.3.5 Comment System"),
        bulletItem("Client users can post comments on any expert", "bullets4"),
        bulletItem("Arches staff can reply (visible in portal, posted from AIS)", "bullets4"),
        bulletItem("On new comment: notification email sent to project Slack address -> posts to Slack channel", "bullets4"),
        bulletItem("Comments are per-expert, per-project (not global to the expert)", "bullets4"),

        heading3("3.3.6 Excel Export"),
        boldPara("Library: ", "SheetJS (xlsx)"),
        boldPara("Columns: ", "ID, Name, Company, Position, Status, Cost, Availability"),
        boldPara("Two modes: ", "Export All (all experts in project) and Export Selected (checked experts only)"),

        heading3("3.3.7 API Endpoints Needed"),
        new Table({
          width: { size: 9360, type: WidthType.DXA }, columnWidths: [1200, 4000, 4160],
          rows: [
            new TableRow({ children: [hCell("Method", 1200), hCell("Endpoint", 4000), hCell("Description", 4160)] }),
            new TableRow({ children: [cell("GET", 1200), cell("/api/projects/:id/experts", 4000), cell("All experts for project, grouped by segment", 4160)] }),
            new TableRow({ children: [cell("GET", 1200, altShading), cell("/api/experts/:expertId", 4000, altShading), cell("Full expert detail (profile, history, screening, availability)", 4160, altShading)] }),
            new TableRow({ children: [cell("POST", 1200), cell("/api/experts/:expertId/decline", 4000), cell("Decline expert {reason, comment}", 4160)] }),
            new TableRow({ children: [cell("GET", 1200, altShading), cell("/api/experts/:expertId/comments", 4000, altShading), cell("Get comments for expert in project context", 4160, altShading)] }),
            new TableRow({ children: [cell("POST", 1200), cell("/api/experts/:expertId/comments", 4000), cell("Post new comment {text}", 4160)] }),
            new TableRow({ children: [cell("POST", 1200, altShading), cell("/api/experts/:expertId/request-availability", 4000, altShading), cell("Request availability from Arches", 4160, altShading)] }),
          ]
        }),

        new Paragraph({ children: [new PageBreak()] }),

        // --- 3.4 Booking ---
        heading2("3.4 Booking Module (Modal)"),
        para("The booking modal is the most complex UI component. It handles timezone conversion, calendar visualization, duration validation, attendee management, and meeting method selection."),

        heading3("3.4.1 Timezone Handling"),
        bulletItem("Auto-detect client timezone on page load using browser Intl API (DST-aware)", "bullets"),
        bulletItem("Dropdown with common timezone labels: e.g. UTC-05:00 (Eastern Time), UTC+09:00 (Japan/Korea)", "bullets"),
        bulletItem("Expert availability is stored in their local timezone in AIS", "bullets"),
        bulletItem("Calendar display converts all times to client-selected timezone", "bullets"),
        bulletItem("DST transitions: browser auto-detects current offset; timezone labels should use IANA database for accuracy", "bullets"),

        heading3("3.4.2 Calendar Timeline"),
        bulletItem("Displays 2-week rolling window of expert availability", "bullets2"),
        bulletItem("15-minute time blocks as clickable grid", "bullets2"),
        bulletItem("Color coding: Green = available, Orange = selected, Red = already booked", "bullets2"),
        bulletItem("Prev/Next week navigation (hidden if no slots in that direction)", "bullets2"),

        heading3("3.4.3 Duration Selection"),
        boldPara("Format: ", "Dropdown (not free-text)"),
        boldPara("Options: ", "30, 45, 60, 75, 90, 105, 120 minutes (15-minute increments)"),

        heading3("3.4.4 Availability Validation (CRITICAL)"),
        para("When the user selects a start time and duration, the system MUST validate that the entire interview duration falls within the expert's available blocks:"),
        numberItem("Convert selected start time + duration to expert's timezone", "numbers2"),
        numberItem("Check that every 15-minute block from start to end is within an availability slot", "numbers2"),
        numberItem("If any block falls outside availability, show warning and BLOCK confirmation", "numbers2"),
        numberItem("Error message: \"The selected duration extends beyond the expert's available time. Please select a shorter duration or a different time slot.\"", "numbers2"),
        para("This prevents scenarios where a user clicks a slot with only 15 minutes of remaining availability and books a 60-minute interview.", { run: { bold: true, color: "E74C3C" } }),

        heading3("3.4.5 Attendees"),
        bulletItem("Default: pre-filled with the logged-in user's email (read-only)", "bullets3"),
        bulletItem("Add attendee: \"+\" button adds a new email input field", "bullets3"),
        bulletItem("Remove attendee: \"x\" button on each added row (cannot remove default)", "bullets3"),
        bulletItem("Validation: at least 1 email required, all must be valid email format", "bullets3"),

        heading3("3.4.6 Meeting Method"),
        boldPara("Format: ", "Dropdown (not radio buttons)"),
        para("Options:"),
        new Table({
          width: { size: 9360, type: WidthType.DXA }, columnWidths: [2800, 3280, 3280],
          rows: [
            new TableRow({ children: [hCell("Option", 2800), hCell("Behavior", 3280), hCell("Integration", 3280)] }),
            new TableRow({ children: [cell("Arches Zoom Link", 2800), cell("System auto-generates Zoom meeting", 3280), cell("Zoom API: Create meeting, include link in calendar invite", 3280)] }),
            new TableRow({ children: [cell("Arches Zoom (Call-in)", 2800, altShading), cell("System generates Zoom with dial-in numbers", 3280, altShading), cell("Zoom API: Create meeting with telephony enabled", 3280, altShading)] }),
            new TableRow({ children: [cell("Client-provided Link", 2800), cell("Client pastes their own meeting URL", 3280), cell("URL input field appears; include in calendar invite as-is", 3280)] }),
          ]
        }),

        heading3("3.4.7 Booking Confirmation"),
        para("After clicking \"Confirm Booking\", a confirmation dialog displays:"),
        bulletItem("Summary: Expert name, Date & Time, Duration, Meeting Method, Attendees", "bullets4"),
        bulletItem("Warning notice: \"This is not a final confirmation. A scheduling request will be sent to the expert for the date/time shown above. Once the expert accepts, a calendar invitation will be issued.\"", "bullets4"),
        bulletItem("Buttons: Back (return to form), Send Request (submit)", "bullets4"),

        heading3("3.4.8 Calendar Invitation (Privacy Requirement)"),
        para("IMPORTANT: Two separate calendar invitations must be created for privacy protection:", { run: { bold: true, color: "E74C3C" } }),
        new Table({
          width: { size: 9360, type: WidthType.DXA }, columnWidths: [2000, 3680, 3680],
          rows: [
            new TableRow({ children: [hCell("Invite", 2000), hCell("Recipients", 3680), hCell("Content", 3680)] }),
            new TableRow({ children: [cell("Client invite", 2000), cell("Client attendee emails only", 3680), cell("Expert name (no personal email/phone), meeting link, date/time", 3680)] }),
            new TableRow({ children: [cell("Expert invite", 2000, altShading), cell("Expert email only", 3680, altShading), cell("Client company name (no individual names/emails), meeting link, date/time", 3680, altShading)] }),
          ]
        }),
        para("Personal information must never cross between client and expert calendars."),

        heading3("3.4.9 Post-Booking Flow"),
        numberItem("Booking request created in AIS with status \"Pending\"", "numbers3"),
        numberItem("Notification sent to Arches project Slack channel", "numbers3"),
        numberItem("Arches contacts expert to confirm the proposed time", "numbers3"),
        numberItem("On expert acceptance: Zoom meeting auto-created via Zoom API", "numbers3"),
        numberItem("Two separate calendar invites sent (client + expert)", "numbers3"),
        numberItem("Expert status updated to \"Interview\" in AIS", "numbers3"),
        numberItem("Client receives email + portal notification: \"Interview confirmed\"", "numbers3"),

        heading3("3.4.10 API Endpoints Needed"),
        new Table({
          width: { size: 9360, type: WidthType.DXA }, columnWidths: [1200, 4000, 4160],
          rows: [
            new TableRow({ children: [hCell("Method", 1200), hCell("Endpoint", 4000), hCell("Description", 4160)] }),
            new TableRow({ children: [cell("GET", 1200), cell("/api/experts/:id/availability", 4000), cell("Available time blocks for the expert", 4160)] }),
            new TableRow({ children: [cell("POST", 1200, altShading), cell("/api/bookings", 4000, altShading), cell("Create booking request {expertId, dateTime, duration, attendees, meetingMethod, notes}", 4160, altShading)] }),
            new TableRow({ children: [cell("POST", 1200), cell("/api/bookings/:id/confirm", 4000), cell("Arches confirms booking -> triggers Zoom API + calendar invites", 4160)] }),
          ]
        }),

        new Paragraph({ children: [new PageBreak()] }),

        // --- 3.5 Interview ---
        heading2("3.5 Interview Page"),
        para("Displays all interviews grouped by segment with status tracking, recording access, and post-call actions."),

        heading3("3.5.1 Status Cards"),
        bulletItem("Booked: count of upcoming interviews", "bullets"),
        bulletItem("Conducted: count of completed interviews", "bullets"),
        bulletItem("Canceled: count of canceled interviews", "bullets"),

        heading3("3.5.2 Interview Table Columns"),
        para("Expert, Status, Date & Time, Duration, Cost, Recording/Transcript/AI Summary, Actions"),

        heading3("3.5.3 Post-Interview Pipeline (Automated)"),
        numberItem("Interview completes on Zoom", "numbers4"),
        numberItem("Zoom recording automatically retrieved via Zoom API", "numbers4"),
        numberItem("Audio sent to transcription service (e.g., Whisper, Deepgram)", "numbers4"),
        numberItem("Transcript sent to LLM for AI summary generation", "numbers4"),
        numberItem("Recording, transcript, and summary stored and linked to interview record", "numbers4"),
        numberItem("Client receives notification: \"Transcript ready for [Expert Name]\"", "numbers4"),

        heading3("3.5.4 Client Actions"),
        new Table({
          width: { size: 9360, type: WidthType.DXA }, columnWidths: [2500, 3430, 3430],
          rows: [
            new TableRow({ children: [hCell("Action", 2500), hCell("When Available", 3430), hCell("Details", 3430)] }),
            new TableRow({ children: [cell("View Recording", 2500), cell("After transcription complete", 3430), cell("Stream or download audio recording", 3430)] }),
            new TableRow({ children: [cell("View Transcript", 2500, altShading), cell("After transcription complete", 3430, altShading), cell("View full text transcript with AI summary", 3430, altShading)] }),
            new TableRow({ children: [cell("Rate Expert", 2500), cell("After interview conducted", 3430), cell("Star rating (1-5) + text feedback. Saved to AIS expert profile.", 3430)] }),
            new TableRow({ children: [cell("Cancel Interview", 2500, altShading), cell("While status is Booked", 3430, altShading), cell("Reason dropdown + comment. Notifies Arches Slack.", 3430, altShading)] }),
            new TableRow({ children: [cell("Contest Duration", 2500), cell("After conducted", 3430), cell("If recorded duration differs from actual. Submit dispute.", 3430)] }),
          ]
        }),

        heading3("3.5.5 API Endpoints Needed"),
        new Table({
          width: { size: 9360, type: WidthType.DXA }, columnWidths: [1200, 4500, 3660],
          rows: [
            new TableRow({ children: [hCell("Method", 1200), hCell("Endpoint", 4500), hCell("Description", 3660)] }),
            new TableRow({ children: [cell("GET", 1200), cell("/api/projects/:id/interviews", 4500), cell("All interviews for project", 3660)] }),
            new TableRow({ children: [cell("GET", 1200, altShading), cell("/api/interviews/:id/recording", 4500, altShading), cell("Stream/download recording", 3660, altShading)] }),
            new TableRow({ children: [cell("GET", 1200), cell("/api/interviews/:id/transcript", 4500), cell("Full transcript + AI summary", 3660)] }),
            new TableRow({ children: [cell("POST", 1200, altShading), cell("/api/interviews/:id/feedback", 4500, altShading), cell("Submit star rating + feedback", 3660, altShading)] }),
            new TableRow({ children: [cell("POST", 1200), cell("/api/interviews/:id/cancel", 4500), cell("Cancel interview {reason, comment}", 3660)] }),
            new TableRow({ children: [cell("POST", 1200, altShading), cell("/api/interviews/:id/contest-duration", 4500, altShading), cell("Contest duration {estimatedDuration, description}", 3660, altShading)] }),
          ]
        }),

        new Paragraph({ children: [new PageBreak()] }),

        // --- 3.6 Billing ---
        heading2("3.6 Billing Page"),
        para("Displays project costs, call breakdown, and invoice management."),

        heading3("3.6.1 Billing Summary"),
        bulletItem("Total Billed: sum of all completed call costs", "bullets"),
        bulletItem("Discount Applied: volume or negotiated discounts", "bullets"),
        bulletItem("Billing Code: client's internal accounting code", "bullets"),

        heading3("3.6.2 Cost Calculation"),
        para("Billing logic already exists in AIS. The portal displays calculated values from AIS. Cost types:"),
        new Table({
          width: { size: 9360, type: WidthType.DXA }, columnWidths: [2500, 3430, 3430],
          rows: [
            new TableRow({ children: [hCell("Type", 2500), hCell("Description", 3430), hCell("Rate Basis", 3430)] }),
            new TableRow({ children: [cell("Interview", 2500), cell("Standard expert call", 3430), cell("Per-call rate (varies by expert, 30/60 min)", 3430)] }),
            new TableRow({ children: [cell("Follow-up Q&A", 2500, altShading), cell("Post-interview written follow-up", 3430, altShading), cell("Flat rate per session", 3430, altShading)] }),
            new TableRow({ children: [cell("Interpretation", 2500), cell("Translation/interpretation service", 3430), cell("Flat rate per session", 3430)] }),
          ]
        }),

        heading3("3.6.3 Invoice Management"),
        bulletItem("\"Request Invoice\" button: triggers invoice generation request", "bullets2"),
        bulletItem("Invoice type selection: All calls, Calls only, Services only, Custom", "bullets2"),
        bulletItem("Optional notes field for custom billing details", "bullets2"),

        heading3("3.6.4 API Endpoints Needed"),
        new Table({
          width: { size: 9360, type: WidthType.DXA }, columnWidths: [1200, 4500, 3660],
          rows: [
            new TableRow({ children: [hCell("Method", 1200), hCell("Endpoint", 4500), hCell("Description", 3660)] }),
            new TableRow({ children: [cell("GET", 1200), cell("/api/projects/:id/billing", 4500), cell("Billing summary + line items", 3660)] }),
            new TableRow({ children: [cell("POST", 1200, altShading), cell("/api/projects/:id/invoice-request", 4500, altShading), cell("Request invoice generation", 3660, altShading)] }),
            new TableRow({ children: [cell("GET", 1200), cell("/api/projects/:id/invoices", 4500), cell("List of past invoices with download links", 3660)] }),
          ]
        }),

        new Paragraph({ children: [new PageBreak()] }),

        // ===== 4. INTEGRATIONS =====
        heading1("4. External Integrations"),

        heading2("4.1 Zoom API"),
        new Table({
          width: { size: 9360, type: WidthType.DXA }, columnWidths: [2500, 6860],
          rows: [
            new TableRow({ children: [hCell("Capability", 2500), hCell("Details", 6860)] }),
            new TableRow({ children: [cell("Create Meeting", 2500), cell("Auto-generate Zoom link when expert confirms booking. Include dial-in numbers for Call-in option.", 6860)] }),
            new TableRow({ children: [cell("Recording Retrieval", 2500, altShading), cell("After meeting ends, automatically download cloud recording via Zoom webhook/API.", 6860, altShading)] }),
            new TableRow({ children: [cell("Authentication", 2500), cell("Server-to-Server OAuth app (Arches Zoom account). No client auth needed.", 6860)] }),
          ]
        }),

        heading2("4.2 Transcription Pipeline"),
        new Table({
          width: { size: 9360, type: WidthType.DXA }, columnWidths: [2500, 6860],
          rows: [
            new TableRow({ children: [hCell("Step", 2500), hCell("Details", 6860)] }),
            new TableRow({ children: [cell("Audio Input", 2500), cell("Zoom recording file (mp4/m4a)", 6860)] }),
            new TableRow({ children: [cell("Transcription", 2500, altShading), cell("Speech-to-text API (Whisper, Deepgram, or similar). Output: timestamped text.", 6860, altShading)] }),
            new TableRow({ children: [cell("AI Summary", 2500), cell("LLM processes transcript to generate structured summary (key points, themes, action items)", 6860)] }),
            new TableRow({ children: [cell("Storage", 2500, altShading), cell("Recording (object storage), Transcript + Summary (database)", 6860, altShading)] }),
          ]
        }),

        heading2("4.3 Slack Notification"),
        bulletItem("Each project has a dedicated Slack email address registered in AIS", "bullets"),
        bulletItem("Client actions (comment, booking, decline, cancel) trigger notification emails to this Slack email", "bullets"),
        bulletItem("Email content auto-posts as a message in the project's Slack channel", "bullets"),
        bulletItem("This is the primary internal notification mechanism for Arches operations team", "bullets"),

        heading2("4.4 Email Notifications (Client-side)"),
        para("Client users receive emails for:"),
        bulletItem("New expert proposed for their project", "bullets2"),
        bulletItem("Interview confirmed (with calendar invite attached)", "bullets2"),
        bulletItem("Transcript/recording ready for download", "bullets2"),
        bulletItem("Reply to their comment from Arches staff", "bullets2"),

        new Paragraph({ children: [new PageBreak()] }),

        // ===== 5. i18n =====
        heading1("5. Internationalization (i18n)"),
        para("The portal supports English and Japanese with a toggle in the header."),
        bulletItem("All user-facing text must use translation keys (data-lang attributes in mockup)", "bullets"),
        bulletItem("Language preference stored per user (localStorage in mockup, should be user profile setting in production)", "bullets"),
        bulletItem("200+ translation keys currently defined in lang.js (included in mockup source)", "bullets"),
        bulletItem("Dynamic content (expert names, dates, etc.) is NOT translated", "bullets"),
        bulletItem("Date/time formatting should respect locale (e.g., Mar 18 vs 3/18)", "bullets"),

        // ===== 6. SECURITY =====
        heading1("6. Security Requirements"),
        bulletItem("Authentication: existing AIS auth system (SSO/OAuth as applicable)", "bullets3"),
        bulletItem("Authorization: all API calls scoped to client organization (no cross-org data access)", "bullets3"),
        bulletItem("Calendar privacy: separate invites for client and expert (Section 3.4.8)", "bullets3"),
        bulletItem("Expert PII: expert personal email/phone never exposed to client portal", "bullets3"),
        bulletItem("Data in transit: HTTPS/TLS for all API calls", "bullets3"),
        bulletItem("Session management: token-based with appropriate expiry", "bullets3"),

        // ===== 7. MOCKUP REFERENCE =====
        heading1("7. Mockup Reference"),
        para("The interactive HTML mockup is available at:"),
        boldPara("Repository: ", "https://github.com/yoshitakasakamoto-collab/client-portal-mockup"),
        para("The mockup contains:"),
        bulletItem("All page layouts with realistic sample data", "bullets4"),
        bulletItem("Interactive booking flow with calendar, timezone conversion, and validation", "bullets4"),
        bulletItem("EN/JP language switching", "bullets4"),
        bulletItem("Excel export functionality (SheetJS)", "bullets4"),
        boldPara("Password: ", "Client216"),
        para("Note: The mockup uses static JavaScript data. In production, all data should come from AIS API endpoints described in this document."),
      ]
    }
  ]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("docs/Client_Portal_Dev_Spec_EN.docx", buffer);
  console.log("EN spec generated: docs/Client_Portal_Dev_Spec_EN.docx");
});
