# Build space in the FBA app: design

Date: 2026-11-03. Status: all four sections approved in chat; waiting for review of this written spec.

Builds on `2026-11-03-cartons-shipments-loads-design.md`, "the records spec". Its **Records**, **Bonded or general**, **Header names** and **Evidence** sections still hold. This spec replaces that spec's **Flow** section and its PL-export lines.

## Goal and decisions (user, 2026-11-03)
1. **All data is entered in the FBA app.** The PL is **generated from the FBA app**, from the records.
2. **Data comes from the container's emails.** When ops start building a container, they **confirm which file(s)** to use, and only then does the app read them.
3. **A "Build" page with a to-do list** (approach A). It is **the builder's work space**: everything the PL builder does today happens here, on the records.
4. Built and tested on the local `fba_cartons` database first. It goes online only when the user says so.

Assumptions (stated, not contradicted):
- The PL keeps today's layout.
- The standalone PL builder site stays up until the user is happy, then is retired for making PLs; Group & Sum can stay.

## The Build page

```
┌ Build ─────────────────────────────────────────────────────────────┐
│ To build (12)   In progress (3)   Built (today 5)        [search]  │
├──────────────┬─────────────────────────────────────────────────────┤
│ to-do list   │ <container> · <file no.> · General|BOND · ETA        │
│              │ ① Files   ② Records   ③ PL                          │
└──────────────┴─────────────────────────────────────────────────────┘
```

**To-do list (left)**

| Tab | Containers |
|---|---|
| **To build** | emails in the archive, no records yet; ETA first |
| **In progress** | read, records not confirmed |
| **Built** | records confirmed and PL exported. A container whose records changed after export shows "PL out of date" |

Each entry shows the client and **General / BOND**. Bonded means `-SUF-` in the file no. (or the BD record's `if_bonded`). A container can also be opened from its row in the Loads grid ("Open in Build").

### ① Files
- Lists the container's attachments from the mail archive, using the archive's container match.
- The app **pre-ticks** the likely load file(s), chosen by content (FBA IDs, FC codes, cartons), and says why, e.g. "15 FBA · 9 FC · 734 ctns".
  - Bonded: the customer PL(s) and the loading list whose box is this container.
- Maple's own replies, invoices, photos and BOLs are listed under "not a load file", folded away.
- **+ Add a file** takes a file that came another way (WeChat, portal).
- **Nothing is read until ops press Read.**

### ② Records
- **The records:** loads → shipments → cartons (the records spec), shown like the carton test page. The user's 16 columns are visible, and every cell is bordered and centred.
- **Checks** are updated after every change. They warn and never block:
  - cartons equal the file's total row (bonded: equal the loading-list total);
  - a shipment with no FBA# or PO#;
  - an FC that is not an Amazon code (`/^[A-Z]{3}\d$/`) and not PA / SELF PICKUP / HOLD;
  - the same FBA–PO twice;
  - a shipment whose cartons don't add up to its file line.
- **What ops can change:**

  | Action | Result |
  |---|---|
  | **Move ▾** a shipment to a load / HOLD / SELF PICKUP / a new FC | The shipment moves; the target load is made if missing; both loads' totals update |
  | **Move ▾ → part of it**: chosen cartons (with a carton list) or a count (without) | With cartons: they move. Without: the shipment splits into two lines with the same FBA–PO |
  | Edit a cell (cartons, pcs, kg, CBM, FBA#, PO#, Job #, remark, …) | That value; totals above update |
  | + Add shipment / ⋯ Remove | For lines the file lacks, or a wrong line |
  | Bonded: change CCN# | The shipment's CCN# block |
  | PA load: address, postal code | Load address |

- Every change is written to the app's change history (`edit_history`: who, when, subject).
- **Read again** (after a corrected file) warns that it replaces this container's shipments and cartons, and says how many hand fixes would be lost.
- **Confirm records** writes load totals and estimated skids (`ceil(CBM/1.7 − 0.05)`). The container moves to "ready for PL".

### ③ PL
- A **preview** of the PL, then **Export PL** (to SharePoint) or **Download only**.
- **Layout: today's PL builder output**, built from the records:
  - **General:** title `DESTUFFING REPORT`; header (client, file #, container #, # of destinations, destuffing time); one table of shipment lines; then a summary per destination.
  - **Bonded:** title `DESTUFFING REPORT (Bonded)`; one `CONTROL NUMBER n:` block per CCN# (EJFB HBLs first, then the others, each in number order); a summary per CCN# × destination.
  - **Columns:** CTNS, G.W., CBM, FBA No., Reference ID, Amazon FC, 单号, 箱唛.
  - **One line per shipment**; Amazon FC, 单号 and 箱唛 merged down a load's lines.
  - **Summary:** total cartons and estimated skids from the records. The actual-skid columns (Standard / 60 / 75 / 96 / overhanging / self-palletized) stay blank.
  - **Extra sheets:** the customer's original load file; and, when the file had a carton list, a **Cartons** sheet with one row per carton (the 16 columns).
- **Made in the browser** with ExcelJS (version pinned, the same as the PL builder), using the PL builder's writer logic (`public/js/writer.js`) ported to the FBA app's frontend.
- **Export PL** saves `PL(<container>).xlsx` to the container's SharePoint folder (where the builder saves it today, through the app's existing SharePoint code) and marks the container **Built**.
- Changing records after export marks the PL **out of date**.

## Errors

| Situation | Behaviour |
|---|---|
| No load file in the emails | Stays in To build with "no load file found"; + Add a file |
| `.xls` file | **Open:** read it with SheetJS (one dependency), or show "can't read .xls". User to decide |
| Unknown layout (no FBA / cartons header) | "not read: no FBA / cartons header found"; ops tick or add another file |
| Totals don't match | Warning with the difference; confirm still allowed |
| Mail archive / SharePoint unreachable | That step says so; the rest of the page works |
| Two people on one container | The second sees "being built by X since hh:mm" (the app's existing editing notice) |

## Testing
- **Reader unit tests** on fixtures made from the table parts of 7 real containers (list in the records spec), with no names or addresses.
- **History check:** a script re-builds the 24 past containers and prints the scoreboard (lines, FBA–PO pairs, cartons vs totals). A change may not lower it.
- **PL check:** for the 13 bonded containers with a saved `PL(<container>).xlsx`, the exported PL's lines and summary are compared with the saved one. Known exceptions are listed (hand edits such as YYZ7 17/18; NW used as GW on BMOU5486746).
- **Browser run** on the local app (port 8789, `fba_cartons`): to-do list → files → read → move to HOLD → confirm → export.

## Local first, then online
- New branch `cartons-shipments-loads` in a **new worktree** of FBA_DATABASE (the main checkout has someone's uncommitted work), running against **`fba_cartons`**.
- The live database, `fba_test` and the live app are not touched.
- SharePoint exports from the local app go to the **test folder** (`_fba-app-test`), never a real container folder.
- **Online:** add the new tables/columns to the live database and deploy, only when the user says so, after comparing the live copy with the branch.

## Out of scope
- Reading email text for HOLD / relabel / self pick-up hints.
- OCR of scanned PDFs.
- The "Amazon Shipping List" PL layout.
- Back-filling old containers.
- Retiring the old PL builder site (later, on the user's word).
