# CPAinc Web App — User Manual

**ConferenceDirect — Kristin House Team**
*Last updated: May 2026 — includes Cost Savings Analysis*

---

## Table of Contents

1. [Overview](#1-overview)
2. [Logging In & Out](#2-logging-in--out)
3. [Status Board](#3-status-board)
4. [Dashboard](#4-dashboard)
5. [Bookings](#5-bookings)
6. [Pickup Tracking and Reporting](#6-pickup-tracking)
7. [RFP Process](#7-rfp-tracking)
8. [Reports](#8-reports)
   - 8.4 Cost Savings Analysis
9. [Contacts](#9-contacts)
10. [Import Tools](#10-import-tools)
11. [Settings](#11-settings)
12. [User Management (Admin Only)](#12-user-management-admin-only)

---

## 1. Overview

CPAinc is a private web application used by the Kristin House team at ConferenceDirect to manage hotel bookings, pickup tracking, RFPs, and commission reporting. Access is role-based — what you can see and do depends on the permissions assigned to your account by an administrator.

### Navigation Bar

The top navigation bar is labelled **Conference Planning Associates Inc.** and contains the following links (left to right):

| Link | Goes to |
|---|---|
| **Status Board** | Action item dashboard for your accounts |
| **Dashboard** | Commission snapshot |
| **Bookings ▾** | Bookings list; New Booking |
| **RFP Process** | RFP tracking |
| **Pickup Tracking & Reporting** | Pickup reporting and tracking |
| **Contracts** | Contract templates |
| **Reports ▾** | Commission and payment reports |
| **Contacts** | Client and hotel contact list |
| **Outlook ▾** | Outlook inbox, calendar, and connection settings |
| **Help ▾** | AI Assistant; Help Guide |
| **Settings ▾** | App settings; User management (admin only) |
| **Import ▾** | Import bookings, payments, HHR, cancelled meetings |
| **Your Name ▾** | My Account; Sign Out |

Links visible to each user depend on their assigned permissions. Admins see all links.

---

## 2. Logging In & Out

- Navigate to the app URL and enter your **email address** and **password**.
- Click **Log In**.
- To log out, click your name in the top-right of the navigation bar, then choose **Sign Out**.

> If you cannot log in or need a password reset, contact your administrator (Peter or Kristin).

---

## 3. Status Board

**Navigation:** Click **Status Board** in the top menu.

**Access required:** Pickup / Payments permission.

The Status Board is your personal action-item dashboard. It scans all of your assigned accounts and surfaces bookings that need attention — missing data, overdue pickups, no recent contact, and similar issues. Items are filtered to only the accounts and bookings you have access to.

At the top of the page, coloured pill badges summarise how many issues exist in each category. Click any pill to jump directly to that section.

### Issue Types

Issues are grouped into the following categories, listed in priority order:

| Category | Colour | What it means |
|---|---|---|
| **Overdue pickup** | 🔴 Red | Event has passed its pickup cutoff date but no final pickup has been entered |
| **Past cutoff — no history** | 🔴 Red | Cutoff has passed and no pickup history exists at all |
| **No recent contact** | 🟠 Orange | No contact logged for this event in the past 14 days and cutoff is approaching |
| **Uniform block** | 🟠 Orange | All nightly block values are identical (possible copy-paste error) |
| **Empty block** | 🟠 Orange | No contracted block data has been entered yet |
| **Missing hotel email** | 🔵 Blue | No hotel contact email on file for this event |
| **Missing client contact** | 🔵 Blue | No client contact email on file for this event |
| **Missing cutoff date** | 🔵 Blue | No cutoff date has been set for this event |
| **Missing room rate** | 🔵 Blue | No contracted room rate has been entered |

### Columns

Each issue row shows:
- **Booking ID** — links directly to the booking detail page
- **Event / Organisation** — event name and client organisation
- **Hotel** — the hotel for this booking
- **Start / End** — event dates
- **Issue** — a short description of the specific problem
- **Fix** — a button that takes you directly to the relevant section to resolve the issue

### Ignoring an Item

If an item is not immediately actionable (e.g. you are waiting on the hotel), you can ignore it for a period of time so it does not clutter the board.

Use the **ignore dropdown** on the right side of each row:

| Option | What happens |
|---|---|
| **Ignore 1 month** | Hides this item for 30 days |
| **Ignore 2 months** | Hides this item for 60 days |
| **Custom date…** | A date picker appears — choose any future date |
| **Permanent** | Hides this item indefinitely (requires confirmation) |

After selecting an option, click the **eye-slash button** (👁‍🗨) to apply it. The row fades out immediately. The item will reappear after the ignore period expires (or never, for permanent ignores).

> **Note:** Permanent ignores do not delete the underlying data — they simply suppress the issue on your board. Contact your administrator if you need to clear a permanent ignore.

---

## 4. Dashboard

**Navigation:** Click **Dashboard** in the top menu.

The Dashboard gives Kristin and the team a financial snapshot of the current year's commission activity.

### Views
- **Kristin House** — shows Kristin's own bookings only.
- **Team** — shows team associate bookings and Kristin's 10% override cut.

Toggle between views using the buttons at the top of the page.

### Summary Cards

| Category | What it means |
|---|---|
| **Completed & Paid** | Meetings that have ended and commission has been received |
| **Future & Paid (Incentive)** | Meetings not yet held but prepayment received |
| **Completed – Has Pickup, Unpaid** | Meeting ended, pickup data exists, but no payment yet |
| **Completed – No Pickup Yet** | Meeting ended, no pickup data entered, no payment |
| **Future / Unpaid** | Upcoming meetings within the current cycle with no payment |

Each card shows the **count**, **estimated revenue**, and **estimated commission value**.

### Meetings Next 30 Days
A count of upcoming meetings starting within the next 30 days.

### Upcoming Meetings Table
A list of the next 10 meetings ordered by start date, showing booking ID, event name, hotel, account, and start date.

---

## 5. Bookings

### 5.1 Bookings List

**Navigation:** Click **Bookings** in the top menu, then **Bookings**.

Displays all active (non-cancelled) bookings. Non-admin users only see bookings for accounts assigned to them.

**Filtering options:**
- **Search** — searches event name, booking name, account, hotel (customer), or booking ID.
- **Status** — filter by booking status (Active, Definite, etc.).
- **Associate** — filter by booking associate.

Each row shows: Booking ID, Account, Event Name, Hotel, Start Date, End Date, Revenue, and Kristin's estimated commission.

### 5.2 Booking Detail

Click any booking row to open the full detail page. This shows:
- All booking fields (dates, hotel, room rate, commission %, revenue, etc.)
- **Pickup history** — actual pickup records entered manually
- **Payment history** — checks/payments received
- **Contract** — any uploaded contract documents
- Links to add new pickups or payments

### 5.3 Adding a New Booking

**Navigation:** Bookings → **+ New Booking** button.

Fill in all available fields. Required fields are marked. Booking ID must match a ConferenceDirect booking number.

### 5.4 Editing a Booking

From the Booking Detail page, click **Edit**. Update any field and save. Changes update the record immediately.

### 5.5 Cancelling a Booking

From the Booking Detail page, click **Cancel Booking**. The booking status is set to Cancelled and it is hidden from the main list.

### 5.6 Contracts

From the Booking Detail page, you can:
- **Upload** a contract PDF/document using the upload button.
- **Download** a previously uploaded contract.
- **Delete** a contract (admin only).

Contract templates can also be managed from the **Contracts** link in the top navigation (requires Contracts permission).

---

## 6. Pickup Tracking and Reporting

**Navigation:** Click **Pickup Tracking & Reporting** in the top menu.

This is the core tool for tracking hotel room block pickup for each event. It shows three sections: **Current Events**, **Past Events**, and **Future Events**.

### 6.1 Pickup Dashboard

#### Toolbar

The toolbar at the top of the page contains the following buttons:

| Button | Purpose |
|---|---|
| **Fill Missing** | Looks up event configurations for Kristin's bookings within your account access that do not yet have a pickup tracking record, and creates placeholder records for them |
| **Customer Report** | Generates a multi-tab Excel pickup report for your clients |

#### Sticky Header

When you scroll down on the Pickup Tracking and Reporting page, the column header row (Booking ID, Event / Organisation, Hotel, etc.) automatically stays pinned near the top of the screen so you always know which column you are reading. The header follows you as you scroll both vertically and horizontally.

#### Current Events

Events with a start date within the next ~120 days that have not been finalised. These are the events you are actively tracking.

**Sort options:**
- **By Date** — sorts by event start date (default).
- **By Customer** — groups and sorts by organisation name.

Each event card shows:
- Event name and hotel
- Contracted block total
- Latest pickup total and % of block
- Pace badge: **On Pace** (≥80%), **Watch** (60–79%), **At Risk** (<60%)
- Days until cutoff
- Last contact date

#### Past Events
Events that have ended or been marked as past. These remain visible until finalised.

#### Future Events
Events starting more than ~120 days out. These are visible for planning but not actively tracked.

#### Archived Events
Events that have been archived and are no longer actively displayed.

### 6.2 Event Detail Page

Click any event name to open its full detail page. This shows:

- **Event info** — hotel, organisation, dates, contracted block by night, cutoff date, attrition %, contracted rate
- **Pickup history table** — all weekly pickup reports, showing total rooms, % of block, and % of attrition for each report date
- **Pace chart** — visual graph of pickup over time vs. contracted block
- **Email tools** — buttons to generate and send emails to the hotel, client, or housing company
- **Rooming list** — upload and review guest rooming lists
- **Contact log** — history of all contact attempts and responses

#### Adding a Weekly Pickup Report

From the Event Detail page, click **+ Add Pickup Report**.
- Enter the **report date** and **pickup by night** (rooms picked up for each date in the block).
- The system automatically calculates the total and % of block.
- Click **Save**.

#### Editing / Deleting a Pickup Report

Click the **Edit** or **Delete** icons next to any weekly report row.

### 6.3 Customer Report (XLSX)

**Navigation:** Pickup Dashboard → **Customer Report** button (top right of Current Events section).

Generates a multi-tab Excel file for sending to customers:
- **Tab 1 — Summary**: all selected events sorted by customer, showing block, pickup, and % of block.
- **Subsequent tabs**: one detailed pace report per event, matching the standard "Pick-Up Update" format.

**Steps:**
1. Click **Customer Report** — a selection screen appears.
2. Check/uncheck customers and individual events as needed.
3. Click **Generate Report** to download the Excel file.

### 6.4 Email Tools

From any Event Detail page, the following emails can be generated:

| Button | Sends to | Purpose |
|---|---|---|
| **Email Hotel** | Hotel contact | Request a pickup report |
| **Email Client** | Client/association | Send pickup update |
| **Email Housing** | Housing company | Send housing form or update |
| **Post-Report Email** | Client | Send final pickup summary after cutoff |

Each button opens a formatted email preview. You can copy/paste into Outlook or launch Outlook directly.

### 6.5 Rooming List

From the Event Detail page, click **Rooming List** to:
- Upload a rooming list file (Excel or CSV).
- Review and confirm the parsed guest list.
- Download a cleaned CSV for the hotel.
- Enter rooming list manually if needed.

### 6.6 Final History

Once a pickup period is complete, click **Enter Final History** from the Event Detail page to record the actual final pickup numbers. This moves the event to the **Past** section on the dashboard.

### 6.7 Event Report (Multi-Hotel)

For events spanning multiple hotels, click **Event Report** to see a combined pickup view across all hotels for that event.

### 6.8 Combined Event Report XLSX

From the Event Report page, click **Download XLSX** to export the full combined pace report in the standard format.

---

## 7. RFP Process

**Navigation:** Click **RFP Process** in the top menu.

Manages the Request for Proposal process for new events.

### RFP Statuses

| Status | Meaning |
|---|---|
| **Sourcing** | Initial stage, identifying hotels |
| **Proposals Received** | Hotel proposals have come in |
| **Negotiating** | Actively negotiating terms |
| **Hotel Selected** | A hotel has been chosen |
| **Contracting** | Working through contract |
| **Contracted** | Contract signed |
| **Dead** | RFP did not proceed |

### 7.1 RFP List

Shows all active RFPs with status badges, client organisation, event name, dates, and hotel count. Click **Show Archived** to see closed/archived RFPs.

Non-admin users only see RFPs for their assigned accounts.

### 7.2 Creating a New RFP

Click **+ New RFP** and fill in:
- Client organisation, event name, RFP name/code
- Dates (primary and alternate), peak rooms, total room nights
- Response due date, decision due date
- Status and notes
- Optionally upload the RFP document

### 7.3 RFP Detail Page

Shows all RFP details plus:
- **Hotels** — list of hotels being considered, with proposal status, proposed rate, commission %, attrition, cutoff days, and F&B.
- **Notes** — running log of notes and updates.
- Links to upload hotel proposals and the RFP document.

#### Adding a Hotel to an RFP
Click **+ Add Hotel** and fill in hotel name, contact details, proposed rate, and terms.

#### Hotel Statuses
Hotels within an RFP can be marked as: Pending, Proposal Received, Shortlisted, Selected, Eliminated, or Declined.

#### Import CRF
If a ConferenceDirect CRF (Compete for Revenue Form) document is available, use **Import CRF** to auto-populate hotel details from the document.

### 7.4 Archiving an RFP

From the RFP Detail page, click **Archive** to move a completed or cancelled RFP to the archived list. Click **Unarchive** to restore it.

---

## 8. Reports

**Navigation:** Click **Reports** in the top menu.

### 8.1 Missing Pickup/Comm Report

**Access required:** "Missing Pickup/Comm Report (Kristin)" or "Missing Pickup/Comm Report (Team)" permission.

Shows bookings within a date range that have **not yet received a commission payment** — either because pickup data is missing or the payment hasn't been recorded.

**How to use:**
1. Select **Kristin** or **Team** view.
2. Enter a **Date From** and **Date To** (meeting start date range).
3. Click **Run Report**.

The report shows each booking with:
- Account, hotel, dates, room rate, commission %
- Estimated commission (based on revenue) and actual commission (based on pickup if available)
- Whether a payment has been received

**Export to Excel:** Click **Export XLSX** to download the report.

### 8.2 Payment Report

**Access required:** "Reports → Payment Report" permission.

Shows all **payments received** within a date range.

**How to use:**
1. Enter a **Date From** and **Date To** (date the check was received).
2. Click **Run Report**.

Each row shows the booking, hotel, check number, payment amount, and flags any payment that appears **out of tolerance** (i.e., the payment received is significantly less than the calculated commission).

**Out of tolerance** payments are highlighted in red. The tolerance threshold is set in Settings.

**Export to Excel:** Click **Export XLSX**.

### 8.3 Customer Summary

**Access required:** "Reports → Customer Summary" permission.

A high-level summary of activity grouped by customer/account. Shows total bookings, revenue, and commission by organisation.

---

### 8.4 Cost Savings Analysis

#### 8.4.1 Overview

The Cost Savings Analysis module produces a per-meeting savings report comparing the hotel's initial **proposal** terms to the signed **contract** terms, then totals the savings across guest rooms, staff rooms, comp rooms, F&B minimum, meeting room rental, attrition, internet, parking, planner points, and any custom line items.

The deliverable is an Excel spreadsheet matching ConferenceDirect's existing Cost Savings Report Template that can be exported one meeting at a time, or as a multi-sheet workbook covering every hotel under an RFP (with a Summary tab).

**Navigation:** **Reports** menu → **Cost Savings Analysis**

---

#### 8.4.2 Where the Data Comes From

A Cost Savings Report is tied to one **RFP hotel** (one row in the RFP's hotel list). The "before" numbers come from the proposal phase; the "after" numbers come from the signed contract.

| "Before" (Proposal Side) | Source |
|---|---|
| Rack rate / group rate | CRF gross room rate; falls back to the hotel's proposed rate |
| F&B Minimum (Initial) | RFP hotel's proposed F&B minimum |
| Meeting Room Rental (Initial) | RFP hotel's proposed rental |
| Proposed attrition % | RFP hotel's proposed attrition |
| Hotel brand | RFP hotel's brand (auto-detected from the hotel name if missing) |

| "After" (Contract Side) | Source |
|---|---|
| Contracted rate | Parsed from the signed contract PDF |
| Total guest/staff nights | Parsed from the contract |
| F&B Minimum (Negotiated) | Parsed from the contract |
| Meeting Room Rental (Negotiated) | Parsed from the contract |
| Comp room policy (1 per N) | Parsed from the contract |
| Final attrition % | Parsed from the contract |

You can always override any value manually on the edit form.

---

#### 8.4.3 Auto-Created Reports

**When you click "Select Hotel" on an RFP**, the app automatically creates a Cost Savings Report for that hotel and pre-fills the proposal-side fields from the CRF data. You'll see a flash message:

> Cost Savings Report initialized from CRF — open it under Reports → Cost Savings.

**When you upload a signed contract** (from a Booking, a Pickup event, or an RFP), the app automatically:

1. Runs the contract through the richer cost-savings AI parser.
2. Caches the parsed JSON so subsequent extractions are instant.
3. Updates **all** Cost Savings Reports linked to that booking's RFP with the new contract-side values.
4. Recomputes totals.

You'll see this flash message after the contract upload:

> Cost Savings reports updated with contract values.

This means **you usually don't need to click any "Auto-fill" button manually** — the values populate themselves through the normal RFP workflow.

---

#### 8.4.4 The Cost Savings Dashboard

**Navigation:** Reports → Cost Savings Analysis (`/cost-savings`)

Lists every Cost Savings Report across all RFPs, sorted by most recently updated.

**Summary cards at the top:**

- **Reports** — total report count
- **Total Savings** — sum of all grand totals (year-to-date)
- **Avg Savings / Meeting** — average grand total
- **Total Hours Worked** — sum of `Hours Worked` field across reports

**Search:** Type any text in the search box to filter by meeting name, hotel name, or RFP code.

**Per-row actions:**

| Icon | Action |
|---|---|
| ✏️ Pencil | Open the report for editing |
| 📗 Excel | Download the single-meeting XLSX |

Reports show a **Draft** (gray) or **Final** (green) status badge.

---

#### 8.4.5 The Per-RFP View

**Navigation:** From the RFP detail page, click the piggy-bank icon (🐷) next to any hotel row.

This page shows one card per hotel on the RFP. Each card displays:

- Hotel name, city, state
- Proposed rate
- Current Cost Savings Report status (No Report / Draft / Final)
- Total Cost Savings (if a report exists)
- Hours worked

**Buttons per card:**

- **Create Report** — manually creates a new report for that hotel (usually unnecessary since selection auto-creates).
- **Open** — jumps to the edit form.
- **Export** — single-meeting XLSX.

**Top-right button:** **Export All (Multi-Sheet)** — only appears when at least one report exists under this RFP. Produces a workbook with one sheet per hotel + a Summary tab that aggregates totals across hotels (matches the NCSL 2026 format).

---

#### 8.4.6 Editing a Report

**Navigation:** Open any report from the dashboard or per-RFP view.

The edit form is organized into cards:

#### Meeting Details

- **Meeting Name** — used as the sheet name and report title
- **Hotel Name**
- **Meeting Dates** — free text, e.g. `12/8/26-12/10/26`
- **Original Lead Date Request** — date the lead was first sent
- **Booked Date** — date the contract was signed
- **Hours Worked** — total hours you spent on the meeting (informational)
- **Hotel Brand** — controls planner points calculation (see [Calculation Types](#8-calculation-types))
- **Status** — Draft or Final

#### Total Cost Savings Banner

A green banner directly below shows the live grand total. **Numbers update as you type** anywhere on the form — no need to click Save to see them refresh.

#### Guest Rooms

| Field | Meaning |
|---|---|
| Rack Rate (Proposal) | Hotel's published rate before negotiation |
| Contracted Rate | Final negotiated rate |
| Total Guestrooms Contracted | Total room nights in the block |
| Notes | Optional free text shown in Excel column C |

**Computed below the inputs:**

- Savings per Room = Rack − Contracted
- Savings on Guestrooms = Savings per Room × Total Nights

#### Staff Rooms

Same shape as Guest Rooms but for the staff block (group rate vs. contracted rate × staff nights).

#### Complimentary Rooms

- **Industry Standard (1 per …)** — default 50 (one comp room per 50 paid)
- **Negotiated Policy (1 per …)** — what you contracted (commonly 40)

Savings = (Guest Nights ÷ Negotiated Policy) × Contracted Rate

#### Meeting Room & F&B

A two-row table:

- **Meeting Room Rental** — Initial Quote vs. Negotiated Quote
- **F&B Minimum** — Initial vs. Negotiated

Savings = Initial − Negotiated (per row). Add free-text Notes per row if helpful.

#### Other Cost Savings Elements

A long table of line items (concessions, internet, parking, attrition, planner points, travel expenses, etc.). Each row has:

| Column | Meaning |
|---|---|
| Item | The concession name (editable) |
| Calc Type | Simple / Attrition / Points / Note Only — drives the formula |
| Standard $ | Pre-negotiation price |
| Negotiated $ | Post-negotiation price |
| Qty / % | Quantity or percentage |
| Savings | Computed automatically |
| Notes | Free text (shown in Excel column F) |

A green badge in the section header shows the **Subtotal** of all item savings.

#### Save Button

Click **Save** at the bottom to persist edits to the database. The Save button is required to lock in any changes — the live total banner shows what *will* be saved, but only Save writes it.

---

#### 8.4.7 Auto-fill Buttons

Two yellow buttons at the top of the edit page:

#### Auto-fill from Proposal

Pulls proposal-side values from the hotel's data:

1. If a PDF/DOCX proposal has been uploaded directly to the hotel row, AI parses it and fills rack/group rates, F&B min, MR rental, attrition, comp ratio, brand, concessions list.
2. Otherwise, falls back to the structured CRF columns (proposed rate, F&B min, MR rental, attrition, brand) and the CRF Q&A free-text (gross room rate is parsed out of the "What is your gross room rate?" answer when present).

You'll see one of these flash messages:

- "Proposal values extracted from uploaded PDF — review and Save."
- "Proposal values pulled from the CRF — review and Save."
- "No proposal data found on this hotel — upload a proposal PDF or import a CRF first."

#### Auto-fill from Contract

Pulls contract-side values:

1. If a cached extraction exists (because someone uploaded the contract via the normal Booking/Pickup/RFP flows), the cached JSON is applied instantly — no AI re-run.
2. Otherwise, the contract PDF is parsed fresh and the JSON is cached for next time.
3. Falls back to a booking-level contract if the RFP's contract slot is empty.

You'll see:

- "Contract values pulled from cached extraction — review and Save."
- "Contract values extracted — review and Save."
- "No contract PDF found on this RFP — upload one first."

**Both buttons prompt for confirmation before overwriting fields** — they show a browser confirm dialog warning that fields will be overwritten.

**Always click Save after auto-fill** to persist the extracted values. Auto-fill writes to the database immediately, but the form is then re-displayed for review.

---

#### 8.4.8 Calculation Types

The **Calc Type** dropdown on each line item controls how the Savings column is computed:

#### Simple

`Savings = (Standard $ − Negotiated $) × Qty`

Use for any concession where the hotel waived or discounted a fixed price.
**Example:** Complimentary Internet in Guest Rooms — Standard $12.95, Negotiated $0, Qty 411 → $5,322.45 savings.

#### Attrition

`Savings = Contracted Rate × Guest Nights × MAX(0.90 − Contracted Attrition, 0)`

Industry baseline is **90% attrition** — your room block is on the hook for 90% of contracted nights even if you don't pick them up. Any contracted attrition lower than 90% is a savings.

- Enter the **contracted attrition** in the Qty column as a decimal (`0.70` for 70%, `0.85` for 85%).
- If contracted attrition is ≥ 90%, savings = $0.

**Example:** $224 rate × 411 nights × MAX(0.90 − 0.70, 0) = $224 × 411 × 0.20 = **$18,412.80**.

#### Points

`Savings = (Brand $/point − Negotiated) × ((Guest Nights × Contracted Rate) + MR Negotiated)`

Auto-computes the dollar value of meeting planner points earned. The Brand $/point comes from the **Hotel Brand** dropdown in Meeting Details:

| Brand | $/point |
|---|---|
| Hyatt | $0.018 |
| Marriott | $0.008 |
| IHG | $0.007 |
| Hilton | $0.005 |
| Preferred | $0.005 |

The Negotiated column should usually stay at 0 (you're "buying" points at $0 vs. their value).

#### Note Only

No calculation. The Savings column shows $0. Use for descriptive concessions that don't have a dollar amount yet (the Notes column in Excel will display "TBD" or your custom text).

**Example:** "Group may bring in own laptops and projectors at no cost" — no fixed dollar value, just a noted concession.

---

#### 8.4.9 Adding & Removing Line Items

#### Default Seed Items

When a report is first created, 17 default line items are seeded:

1. Attrition Savings vs 90% Baseline
2. Complimentary Internet in Guest Rooms
3. Complimentary Internet in Meeting Space
4. Suite upgrades per night at group rate
5. VIP Amenities
6. Complimentary Overnight Valet Vouchers
7. 15% Discount on A/V
8. Group may bring in own laptops and projectors
9. Reduced valet parking fee
10. Complimentary easels and podiums
11. Complimentary meeting planner room
12. No Resort Fees or Destination Fees
13. Waived storage and handling fees
14. No deposit required
15. Travel Expenses – Site Visits
16. Travel Expenses – Attending Meeting Dates
17. Meeting Planner Points

Most are pre-populated with industry-standard "Standard $" values you can edit. The Travel Expenses rows start at $0 — enter your actual travel cost in the Standard column.

#### Add Item

Click **+ Add Item** at the bottom of the Other Cost Savings Elements table. A new "New Item" row appears with Calc Type = Simple. Rename it and fill in values.

#### Remove Item

Click the trash icon (🗑) at the end of any row. A browser confirm dialog appears — click OK to delete.

#### Reorder

Items are stored with a `sort_order` value. Currently there is no drag handle; items appear in the order they were added. (If you need to reorder, delete and re-add, or contact Peter.)

---

#### 8.4.10 Exporting to Excel

#### Single Meeting

Click the green **Export** button at the top of the edit form, or the green Excel icon on the dashboard / per-RFP view.

The output is a single-sheet XLSX matching the ConferenceDirect Cost Savings Report Template:

- A1: Meeting Name
- A2: Hotel Name
- A3: Meeting Dates
- A4: Lead Date Request
- A5: Booked
- B8/B9: Rack / Contracted rate
- B11: Total Guestrooms Contracted
- B12: `=B11*B10` (Savings on Guestrooms)
- B15–B19: Staff rooms section
- B22–B24: Comp rooms section
- B27–D28: Meeting Room and F&B
- Rows 31+: Other Cost Savings Elements (formulas inserted per calc type)
- `TOTAL` row sums the Other items
- `Hours Worked` row
- `TOTAL COST SAVINGS:` final formula `=B12+B19+B24+D27+D28+E<total_items_row>`

Filename: `<Meeting_Name>_Cost_Savings.xlsx` (special characters stripped)

#### Multi-Sheet (One RFP, All Hotels)

On the per-RFP view, click **Export All (Multi-Sheet)** at the top-right. Produces a workbook containing:

- One Summary sheet with one row per hotel + a Totals row
- One per-hotel sheet (same shape as the single-meeting export)

The Summary tab uses cross-sheet references (e.g. `='2026 Opioid Policy'!B44`) so editing any individual sheet automatically updates the summary — matches the NCSL 2026 multi-meeting format.

Filename: `<RFP_Code>_Cost_Savings_Report.xlsx`

#### Final Manual Editing

Open the exported XLSX in Excel and:

- Re-arrange rows if needed
- Add or remove rows beyond what the app produced
- Re-style headers, colors, fonts to match client deliverable

The exported file is your deliverable — the app's job ends at this point.

---

#### 8.4.11 Marking a Report Final

Change the **Status** dropdown at the top of the edit form from **Draft** to **Final**, then click Save.

The Final badge (green) appears on the dashboard and per-RFP view. This is informational only — it doesn't lock the report; you can still edit it.

---

#### 8.4.12 Common Tasks — Quick Reference

#### "I just selected a hotel on an RFP. Where's the Cost Savings Report?"

Reports → Cost Savings Analysis (or click the piggy-bank icon on the RFP detail page → Open). The report was auto-created with proposal-side values pre-filled.

#### "I just uploaded a signed contract. Did the cost savings update?"

Yes — the upload flow automatically ran the cost savings extraction and updated every linked report. Open the report to verify; the contract-side fields should now be filled in. No manual click required.

#### "Auto-fill says 'No proposal data found.' Why?"

This hotel has no proposal PDF AND no CRF data (no `proposed_rate`, no `crf_row_data`). Either:

- Upload a proposal PDF on the hotel row (Upload button in the RFP detail page), or
- Import a CRF Excel for this RFP, or
- Fill the proposal-side fields manually on the form.

#### "Savings on Guestrooms shows $0 but I entered values."

Check that you entered both **Rack Rate** AND **Contracted Rate** AND **Total Guestrooms Contracted**. All three are required for the calc. The live update should happen as you type — if it still shows $0, refresh the page.

#### "The grand total didn't update after I changed something."

The grand total banner updates as you type. If it didn't, refresh the page — the JS may have errored. Saving and re-opening always shows the correct total.

#### "How do I add a custom concession not in the seed list?"

Click **+ Add Item**, rename "New Item" to your concession (e.g. "Complimentary Cocktail Reception"), pick Calc Type, fill Standard/Negotiated/Qty. Click Save.

#### "Can I delete a report?"

Yes — open the report and click the red trash button at the top right. A confirm dialog appears. Deletion is permanent (it also removes all line items).

#### "How do I produce the deliverable my client expects?"

1. Open the report.
2. Click **Auto-fill from Proposal** (or verify the proposal-side fields are correct).
3. Click **Auto-fill from Contract** (or verify the contract-side fields are correct).
4. Adjust any line items, add custom concessions if needed.
5. Set Status to **Final**, click Save.
6. Click **Export** to download the XLSX.
7. Open in Excel, do any final styling, send to client.

For a multi-hotel event, use **Export All (Multi-Sheet)** from the per-RFP view instead of step 6.

---

---

## 9. Contacts

**Navigation:** Click **Contacts** in the top menu.

Maintains two contact lists used throughout the app (e.g., for auto-filling email fields):

- **Client contacts** — association/organisation contacts.
- **Hotel contacts** — hotel sales and catering contacts.

### Adding a Contact
Click **+ Add Client Contact** or **+ Add Hotel Contact** and fill in name, title, email, phone, and company.

### Editing a Contact
Click the **Edit** icon on any contact row, update the fields, and save.

### Deleting a Contact
Click the **Delete** icon. Deletions are permanent.

---

## 10. Import Tools

**Navigation:** Click **Import** in the top menu.

All import tools require the corresponding permission. Imports are additive — they add or update records but do not delete existing data.

### 10.1 Import Bookings

Imports booking data from the ConferenceDirect Pipeline report.

**Accepted formats:** Excel (.xlsx) or CSV.

**Steps:**
1. Export the Pipeline report from ConferenceDirect.
2. Go to **Import → Bookings**.
3. Upload the file and click **Import**.
4. Review the summary (records added/updated/skipped).

### 10.2 Import Payments

Imports check/payment records.

**Steps:**
1. Upload your payments file.
2. The system matches payments to bookings by Booking ID.
3. Review the import result for any unmatched records.

### 10.3 Import Cancelled Meetings

Imports cancellation records to update booking statuses to Cancelled.

### 10.4 Import HHR (Housing History Report)

Imports Housing History Report data to update pickup records for events.

**Steps:**
1. Go to **Import → HHR**.
2. Upload the HHR file.
3. Review and confirm the import.

### 10.5 Import Voucher

Imports voucher/commission payment data.

---

## 11. Settings

**Navigation:** Click **Settings** in the top menu, then **Settings** (admin only).

### Commission Split
The default percentage of commission that goes to Kristin's side of the split. Enter as a percentage (e.g., `60` for 60%).

### Kristin House Settings
- **Kristin Split** — Kristin's personal share of her own bookings (e.g., 70%).
- **Team Cut** — Kristin's override percentage on team associate bookings (e.g., 10%).

### Payment Tolerance
The threshold (%) below which a received payment is flagged as "out of tolerance" on the Payment Report. For example, `5` means payments more than 5% below the calculated commission are flagged.

### Date Display Format
Controls how dates are shown throughout the entire site (dashboards, reports, pickup tracking, status board, etc.).

Select your preferred format from the dropdown. A **live preview** updates immediately so you can see what the selected format looks like before saving.

Available formats include:

| Format | Example |
|---|---|
| MM/DD/YYYY | 05/10/2026 |
| M/D/YYYY | 5/10/2026 |
| YYYY-MM-DD | 2026-05-10 |
| DD/MM/YYYY | 10/05/2026 |
| Mon DD, YYYY | May 10, 2026 |
| DD Mon YYYY | 10 May 2026 |
| Month DD, YYYY | May 10, 2026 |
| Weekday, Month DD, YYYY | Sunday, May 10, 2026 |

Click **Save** to apply the format site-wide.

### Account-Level Commission Overrides
Some clients have a negotiated commission split that differs from the default. Add overrides here by specifying:
- **Account name** — must match exactly as it appears in the system.
- **Split rate** — the override percentage.
- **Countries** — optionally limit the override to specific countries (leave blank for all).

---

## 12. User Management (Admin Only)

**Navigation:** Click **Settings** in the top menu, then **Users**.

### 12.1 Permissions Table

Each user has a set of permissions that control what they can access. Toggle permissions on/off with the switches in each column.

| Permission | Controls access to |
|---|---|
| **Dashboard** | The commission dashboard page |
| **Admin Panel** | User management and admin tools |
| **Contracts** | The Contracts nav link and contract template pages |
| **Import → Bookings** | Booking import tool |
| **Import → Payments** | Payment import tool |
| **Import → HHR** | Housing History Report import |
| **Import → Cancelled Meetings** | Cancelled meetings import |
| **Missing Pickup/Comm Report (Kristin)** | Kristin's missing commission report |
| **Missing Pickup/Comm Report (Team)** | Team missing commission report |
| **Reports → Payment Report** | Payment report |
| **Reports → Customer Summary** | Customer summary report |
| **View Bookings** | Bookings list and detail pages |
| **Add / Edit Bookings** | Creating and editing booking records |
| **Add Pickups / Payments** | Adding pickup/payment records; access to Status Board, Pickup Tracking and Reporting, RFP's, and Contacts |

Admins automatically have all permissions.

### 12.2 Account Access

Controls **which client accounts** a non-admin user can see across Bookings, Pickup Tracking and Reporting, RFPs, and the Status Board.

**To set account access:**
1. Click the **buildings icon** button in the Account Access column for a user.
2. In the popup, check the accounts this user should be able to see.
3. Use **Search** to filter the list, or **Select All / Clear All** for bulk changes.
4. Click **Save Access**.

If a user has no accounts assigned, they will see no records. Admins always see everything.

### 12.3 Active / Inactive Users

Toggle the **Active** switch to disable a user's login without deleting their account. Inactive users cannot log in but their data is preserved.

### 12.4 Reset Password

Scroll to the **Reset Password** section at the bottom of the User Management page.
1. Select the user from the dropdown.
2. Enter a new password (minimum 6 characters).
3. Click **Reset**.

---

*For technical issues or questions about the app, contact Peter Wann.*
