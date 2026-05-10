# CPAinc Web App — User Manual

**ConferenceDirect — Kristin House Team**
*Last updated: May 2026*

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
