# Cost Savings Analysis — User Manual

**ConferenceDirect — Promagent & CPAinc**
*Last updated: May 2026*

---

## Table of Contents

1. [Overview](#1-overview)
2. [Where the Data Comes From](#2-where-the-data-comes-from)
3. [Auto-Created Reports](#3-auto-created-reports)
4. [The Cost Savings Dashboard](#4-the-cost-savings-dashboard)
5. [The Per-RFP View](#5-the-per-rfp-view)
6. [Editing a Report](#6-editing-a-report)
7. [Auto-fill Buttons](#7-auto-fill-buttons)
8. [Calculation Types](#8-calculation-types)
9. [Adding & Removing Line Items](#9-adding--removing-line-items)
10. [Exporting to Excel](#10-exporting-to-excel)
11. [Marking a Report Final](#11-marking-a-report-final)
12. [Common Tasks — Quick Reference](#12-common-tasks--quick-reference)

---

## 1. Overview

The Cost Savings Analysis module produces a per-meeting savings report comparing the hotel's initial **proposal** terms to the signed **contract** terms, then totals the savings across guest rooms, staff rooms, comp rooms, F&B minimum, meeting room rental, attrition, internet, parking, planner points, and any custom line items.

The deliverable is an Excel spreadsheet matching ConferenceDirect's existing Cost Savings Report Template that can be exported one meeting at a time, or as a multi-sheet workbook covering every hotel under an RFP (with a Summary tab).

**Navigation:** **Reports** menu → **Cost Savings Analysis**

---

## 2. Where the Data Comes From

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

## 3. Auto-Created Reports

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

## 4. The Cost Savings Dashboard

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

## 5. The Per-RFP View

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

## 6. Editing a Report

**Navigation:** Open any report from the dashboard or per-RFP view.

The edit form is organized into cards:

### Meeting Details

- **Meeting Name** — used as the sheet name and report title
- **Hotel Name**
- **Meeting Dates** — free text, e.g. `12/8/26-12/10/26`
- **Original Lead Date Request** — date the lead was first sent
- **Booked Date** — date the contract was signed
- **Hours Worked** — total hours you spent on the meeting (informational)
- **Hotel Brand** — controls planner points calculation (see [Calculation Types](#8-calculation-types))
- **Status** — Draft or Final

### Total Cost Savings Banner

A green banner directly below shows the live grand total. **Numbers update as you type** anywhere on the form — no need to click Save to see them refresh.

### Guest Rooms

| Field | Meaning |
|---|---|
| Rack Rate (Proposal) | Hotel's published rate before negotiation |
| Contracted Rate | Final negotiated rate |
| Total Guestrooms Contracted | Total room nights in the block |
| Notes | Optional free text shown in Excel column C |

**Computed below the inputs:**

- Savings per Room = Rack − Contracted
- Savings on Guestrooms = Savings per Room × Total Nights

### Staff Rooms

Same shape as Guest Rooms but for the staff block (group rate vs. contracted rate × staff nights).

### Complimentary Rooms

- **Industry Standard (1 per …)** — default 50 (one comp room per 50 paid)
- **Negotiated Policy (1 per …)** — what you contracted (commonly 40)

Savings = (Guest Nights ÷ Negotiated Policy) × Contracted Rate

### Meeting Room & F&B

A two-row table:

- **Meeting Room Rental** — Initial Quote vs. Negotiated Quote
- **F&B Minimum** — Initial vs. Negotiated

Savings = Initial − Negotiated (per row). Add free-text Notes per row if helpful.

### Other Cost Savings Elements

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

### Save Button

Click **Save** at the bottom to persist edits to the database. The Save button is required to lock in any changes — the live total banner shows what *will* be saved, but only Save writes it.

---

## 7. Auto-fill Buttons

Two yellow buttons at the top of the edit page:

### Auto-fill from Proposal

Pulls proposal-side values from the hotel's data:

1. If a PDF/DOCX proposal has been uploaded directly to the hotel row, AI parses it and fills rack/group rates, F&B min, MR rental, attrition, comp ratio, brand, concessions list.
2. Otherwise, falls back to the structured CRF columns (proposed rate, F&B min, MR rental, attrition, brand) and the CRF Q&A free-text (gross room rate is parsed out of the "What is your gross room rate?" answer when present).

You'll see one of these flash messages:

- "Proposal values extracted from uploaded PDF — review and Save."
- "Proposal values pulled from the CRF — review and Save."
- "No proposal data found on this hotel — upload a proposal PDF or import a CRF first."

### Auto-fill from Contract

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

## 8. Calculation Types

The **Calc Type** dropdown on each line item controls how the Savings column is computed:

### Simple

`Savings = (Standard $ − Negotiated $) × Qty`

Use for any concession where the hotel waived or discounted a fixed price.
**Example:** Complimentary Internet in Guest Rooms — Standard $12.95, Negotiated $0, Qty 411 → $5,322.45 savings.

### Attrition

`Savings = Contracted Rate × Guest Nights × MAX(0.90 − Contracted Attrition, 0)`

Industry baseline is **90% attrition** — your room block is on the hook for 90% of contracted nights even if you don't pick them up. Any contracted attrition lower than 90% is a savings.

- Enter the **contracted attrition** in the Qty column as a decimal (`0.70` for 70%, `0.85` for 85%).
- If contracted attrition is ≥ 90%, savings = $0.

**Example:** $224 rate × 411 nights × MAX(0.90 − 0.70, 0) = $224 × 411 × 0.20 = **$18,412.80**.

### Points

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

### Note Only

No calculation. The Savings column shows $0. Use for descriptive concessions that don't have a dollar amount yet (the Notes column in Excel will display "TBD" or your custom text).

**Example:** "Group may bring in own laptops and projectors at no cost" — no fixed dollar value, just a noted concession.

---

## 9. Adding & Removing Line Items

### Default Seed Items

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

### Add Item

Click **+ Add Item** at the bottom of the Other Cost Savings Elements table. A new "New Item" row appears with Calc Type = Simple. Rename it and fill in values.

### Remove Item

Click the trash icon (🗑) at the end of any row. A browser confirm dialog appears — click OK to delete.

### Reorder

Items are stored with a `sort_order` value. Currently there is no drag handle; items appear in the order they were added. (If you need to reorder, delete and re-add, or contact Peter.)

---

## 10. Exporting to Excel

### Single Meeting

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

### Multi-Sheet (One RFP, All Hotels)

On the per-RFP view, click **Export All (Multi-Sheet)** at the top-right. Produces a workbook containing:

- One Summary sheet with one row per hotel + a Totals row
- One per-hotel sheet (same shape as the single-meeting export)

The Summary tab uses cross-sheet references (e.g. `='2026 Opioid Policy'!B44`) so editing any individual sheet automatically updates the summary — matches the NCSL 2026 multi-meeting format.

Filename: `<RFP_Code>_Cost_Savings_Report.xlsx`

### Final Manual Editing

Open the exported XLSX in Excel and:

- Re-arrange rows if needed
- Add or remove rows beyond what the app produced
- Re-style headers, colors, fonts to match client deliverable

The exported file is your deliverable — the app's job ends at this point.

---

## 11. Marking a Report Final

Change the **Status** dropdown at the top of the edit form from **Draft** to **Final**, then click Save.

The Final badge (green) appears on the dashboard and per-RFP view. This is informational only — it doesn't lock the report; you can still edit it.

---

## 12. Common Tasks — Quick Reference

### "I just selected a hotel on an RFP. Where's the Cost Savings Report?"

Reports → Cost Savings Analysis (or click the piggy-bank icon on the RFP detail page → Open). The report was auto-created with proposal-side values pre-filled.

### "I just uploaded a signed contract. Did the cost savings update?"

Yes — the upload flow automatically ran the cost savings extraction and updated every linked report. Open the report to verify; the contract-side fields should now be filled in. No manual click required.

### "Auto-fill says 'No proposal data found.' Why?"

This hotel has no proposal PDF AND no CRF data (no `proposed_rate`, no `crf_row_data`). Either:

- Upload a proposal PDF on the hotel row (Upload button in the RFP detail page), or
- Import a CRF Excel for this RFP, or
- Fill the proposal-side fields manually on the form.

### "Savings on Guestrooms shows $0 but I entered values."

Check that you entered both **Rack Rate** AND **Contracted Rate** AND **Total Guestrooms Contracted**. All three are required for the calc. The live update should happen as you type — if it still shows $0, refresh the page.

### "The grand total didn't update after I changed something."

The grand total banner updates as you type. If it didn't, refresh the page — the JS may have errored. Saving and re-opening always shows the correct total.

### "How do I add a custom concession not in the seed list?"

Click **+ Add Item**, rename "New Item" to your concession (e.g. "Complimentary Cocktail Reception"), pick Calc Type, fill Standard/Negotiated/Qty. Click Save.

### "Can I delete a report?"

Yes — open the report and click the red trash button at the top right. A confirm dialog appears. Deletion is permanent (it also removes all line items).

### "How do I produce the deliverable my client expects?"

1. Open the report.
2. Click **Auto-fill from Proposal** (or verify the proposal-side fields are correct).
3. Click **Auto-fill from Contract** (or verify the contract-side fields are correct).
4. Adjust any line items, add custom concessions if needed.
5. Set Status to **Final**, click Save.
6. Click **Export** to download the XLSX.
7. Open in Excel, do any final styling, send to client.

For a multi-hotel event, use **Export All (Multi-Sheet)** from the per-RFP view instead of step 6.

---

*End of Cost Savings Analysis User Manual.*
