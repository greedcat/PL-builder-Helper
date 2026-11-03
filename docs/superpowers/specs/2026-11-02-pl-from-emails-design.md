# PL from email attachments: design

Date: 2026-11-02. Status: SUPERSEDED by 2026-11-03-cartons-shipments-loads-design.md (the user moved the flow to customer files → system records → PL export). Its reading rules carry over.

## Goal

Ops can drop a container's email attachments onto the PL builder all at once:
- the customer PL(s) (派送资料),
- the loading list (拼箱配载清单),
- the EMF CLOSE PDFs,
- and anything else that came in the thread.

The builder then makes the packing list itself, bonded when the cargo moves in bond. It is reviewed and edited in the preview as today.

Worked example: container CIMU1164134. The correct result is `PL(CIMU1164134).xlsx` on SharePoint (5106Team). How it was derived is written up in the session notes, `pl-from-email-CIMU1164134-full-read.md`.

## What the user said / decided
- Approach A: everything runs in the browser, and the combined table goes into the existing pipeline.
- The agents and layouts sending these files are "not sure / mixed". So sorting the files and finding columns must work from headers, not from fixed positions.
- Estimated skids keep the builder's rule `ceil(CBM/1.7 − 0.05)`. The 17 for YYZ7 on the correct PL was a hand edit, so the builder shows 18. The history test found that saved PLs agree with `round(CBM/1.7)` more often (63/81 vs 49/81). The user chose to keep the current rule anyway.
- Decisions made in the email text are not read. Ops apply them in the preview as today:
  - customer self pick-up ("…这票是客户自提的"),
  - relabel to another FC (换标),
  - HOLD (客户还在刷, the customer is still deciding).

## Assumptions (not stated by the user)
- A container can have several customer PLs, one per FBA HBL.
- Dropping or pasting a single Excel file keeps working exactly as today.
- No server change, and no new npm dependency. pdf.js comes from a CDN, pinned like SheetJS and ExcelJS.

## User flow
1. The drop zone and file picker accept several files (`multiple`, accepting `.xlsx,.xls,.pdf` and anything else).
   - One Excel file and no PDF: the current path, unchanged.
   - Two or more files, or any PDF: the email import.
2. A "From emails" panel under the drop zone lists every file with what it was read as. For example:
   - `EJFBTOR26S1870-派送资料.xlsx`: Customer PL, HBL EJFBTOR26S1870, 7 rows, 1060 ctns
   - `拼箱配载清单HLCUTA12608FRQA0.xls`: Loading list, 7 HBLs, 1077 ctns
   - `EMF CLOSE EJFBTOR26S1870.pdf`: EMF, in bond, CCNs EJTOR26S0097 / EJTOR26S0099 / EJFBTOR26S1870
   - `Arrival Notice.pdf`: not used (with a short reason)
3. A check line compares cartons:
   - It reads, for example, `Customer PLs 1060 + loading list only 2 + 15 = 1077 ✓ loading list total 1077`.
   - On a mismatch it shows ✗ and the difference.
   - It is a warning, never a block.
4. The combined table loads as a one-sheet workbook named "From emails". It goes through `loadPastedRows(rows, merges)`, the paste path, so the grid, range detection, column mapping, preview, edits and writer are all unchanged.
5. The Bonded switch is turned on for every email import. In the history test all 13 saved PLs were Bonded, but one container had no EMF and some EMFs say "Import". The CCN# list is filled in the combined table's order. The user can still switch Bonded off.

## Components

### `public/js/from-emails.js` (new; pure logic, no DOM)
It is loaded in the page, and Node tests can `require` it, the same pattern as `shared/paste.js`.

- `classifySheet(aoa) → { kind: 'customerPL' | 'loadingList' | null, headerRow, cols }`
  - Customer PL: one header row has FBA No + CTNS + Amazon FC. Headers are matched with the builder's existing keyword lists where they exist (columns.js / detect.js).
  - Loading list: one header row has 业务编号 + 件数 (or JOB NO + PKGS).
- `readCustomerPL(aoa, fileName, merges) → { hbl, rows: [{ctns, gw, cbm, fba, ref, fc, job, mark, extra…}], total }`
  - The data rows run from the header down to the first total row or blank stretch.
  - Every column of the source is kept, so the builder's own mapping and dropped columns work as for a file.
  - HBL, first match wins:
    1. an HBL-shaped token in the file name (`[A-Z]{2,}\d{2}S\d{4}` style, the same as the loading-list job numbers);
    2. a cell above the header that equals a loading-list job number;
    3. the loading-list job whose cartons equal this PL's total.
  - **Total row.** A row with no FBA No / FC / 单号 / Reference ID, and only numbers (or nothing) to the left of CTNS. Any text there, such as a description, makes it a data row.
    - Do **not** test "cartons = sum of the rows above". That test cut two real PLs short when a row happened to equal the sum.
  - **Layout not supported.** A customer PL with no FBA/CTNS/FC header row is listed as "not read". Example: a goods table plus a separate "Amazon Shipping List" block. Its HBL then falls back to its loading-list row.
- `readLoadingList(aoa) → { box, jobs: [{job, entry, pkgs, kg, cbm, fc, note}], total }`
  - `box`: the container number in the rows above the header (CONTAINER NO.).
  - `fc` is 进仓数据 with a leading `CA` removed.
  - The TOTAL row gives `total`.
- `readEmf(text) → { own, inBond, ccns: [...] }`
  - `own`: this EMF's HOUSE BILL CCN without its carrier code.
  - `inBond`: the text has "In bond".
  - `ccns`: the lines under "RELATED CCN", with the 4-character carrier code (`8FXN`) removed.
- `combine({ customerPLs, loadingList, emfs }) → { headers, rows, merges, ccnOrder, bonded, check }`
  - **Headers.** The first customer PL's header row, plus `CCN#`.
    - With no customer PL, it is a minimal header: `CTNS, G.W.(KG), CBM, FBA No., Reference ID, Amazon FC, 单号, 箱唛, CCN#`.
  - **Customer-PL rows.** Copied as-is, with CCN# = that PL's HBL.
  - **A blank FC on a customer-PL row:**
    - if the row's 单号 matches a loading-list job, that job's FC;
    - otherwise, left blank, merged under the FC above (see merges).
  - **Which files are this container's.** Threads can mention several containers, and users may drop extra files.
    - With several loading lists, the one whose `box` is the container wins. The container number comes from the files' names or the first loading list.
    - Only EMFs whose `own` CCN is one of its jobs are used. One real EMF named the wrong container but had the right CCN.
    - Only customer PLs whose HBL is one of its jobs are used.
  - **Private addresses.** An FC that is not an Amazon code (`[A-Z]{3}\d`), e.g. `HWC`, becomes `Private Address 1`, `Private Address 2`, … The number goes up at each row that has its own 单号; rows below without one continue the same address. The address text stays in the FBA No. column as the customer wrote it. A loading-list-only job with 进仓数据 `HWC` is a private address too. The FBA app calls these PA.
  - **Loading-list jobs not covered by a customer PL:** one row each. A job is covered when its 业务编号, or its 进仓编号, appears anywhere in a customer PL's 单号 or 箱唛 cells. Those cells can be "+"-joined (`EJFBTOR26S1701+EJFBTOR26S1702`) or written as an entry number with a QM suffix (`EJFBTOR26S1796QM`).
    - CTNS / G.W. / CBM come from 件数 / 重量 / 体积, and CCN# = the job.
    - FC is the job's FC. If the FC is blank, or the note has 自提 or 自清, FC = `Self Pick Up`.
    - In the CIMU example, 1830 / 1831 / 1846 / 1876 appear in 1870's 单号 column, so they are not added again. 0097 and 0099 are added.
  - **Merges.** In FC, 单号 and 箱唛, a blank cell directly under a filled cell in the same CCN block is merged into it. The source file's merges are also kept for customer-PL rows.
  - **Order.** CCN blocks: the FBA HBLs (`EJFB…`) first, then the others. Each group is in ascending number order, as on every saved PL. The EMF RELATED CCN order is not used.
  - **`check`.** `{ plTotal, extraTotal, sum, listTotal, ok }`.

### `public/js/ui-load.js` (changed)
- The drop and file-input handlers take every file.
- They route to `loadFromEmails(files)` when there are 2 or more files or any PDF. Otherwise the current `setExcelFile` runs.
- `loadFromEmails`:
  - reads each sheet with SheetJS (`!merges` kept) and each PDF with pdf.js;
  - calls `combine`;
  - renders the panel;
  - sets Bonded and `ccnOrder`;
  - calls `loadPastedRows` with the file name "From emails".
- The panel is cleared whenever a new file or paste is loaded.

### `public/packing_list_builder.html`, `public/css/packing-list.css` (changed)
- The file input gets `multiple`, and the drop text reads "an Excel file — or all of a container's email attachments".
- The panel markup and its styles are added.
- The script tag for `from-emails.js` is added.
- pdf.js is loaded on demand (a dynamic import of a pinned version) only when a PDF is dropped.

## Errors
- **A file that can't be read** (corrupt, password-protected, an image, a scanned PDF with no text): listed as "not used: couldn't read", and the rest still build.
- **No customer PL and no loading list:** a status error ("No packing list or loading list found among the files"), and nothing is loaded.
- **A customer PL whose HBL can't be found:** its rows still load, with an empty CCN#. The panel says so, and the user assigns the CCN# with the existing "Assign CCN#" bar.
- **pdf.js fails to load** (offline): PDFs are listed as not used, and Bonded stays as the user set it.

## Testing
- `test/from-emails.test.js` (node:test), on fixtures under `test/fixtures/cimu1164134/`.
  - The fixtures are small arrays-of-arrays made from the real files' table parts only: no shipper/consignee addresses, emails or phone numbers. The EMF fixtures are the extracted RELATED CCN / In bond text lines.
  - Assertions:
    - the files are sorted correctly;
    - the HBL found is EJFBTOR26S1870;
    - there are 9 combined rows: 7 FBA plus 0097 and 0099 as Self Pick Up (2 / 2930 / 1.39 and 15 / 8477 / 9.75);
    - the FC of row 2 is filled or merged as YOO1;
    - the CCN order is 1870, 0097, 0099;
    - the check is 1060 + 17 = 1077 ok;
    - in bond = true.
- More fixtures, from the history test, each built from its table parts only:
  - **CMAU8584800:** private addresses (`HWC` → Private Address 1 = 360, Private Address 2 = 124). It also has a PL that covers another job through 箱唛 `EJFBTOR26S1796QM` (1818 = 119, with 1796 not added again).
  - **TLLU4625944:** a row whose cartons equal the sum above it is **not** the total row. The PL reads 486 cartons.
  - **TCKU9024600:** other containers' loading lists, PLs and EMFs in the same drop are ignored.
- Existing tests still pass.
- Browser check: drop the CIMU files on the local server.
  - The preview should match the correct PL's rows and merges.
  - Summary: YOO1 355/14, YYZ4 140/2, YYZ7 565/18, 0097 2/1, 0099 15/6.

## History test (2026-11-02)
These rules were tested with a throwaway prototype on 13 past E-Cargo Master containers. Each was built from its email attachments alone and compared with the PL saved on SharePoint and the FBA app's loads:
- Cartons equal the loading-list total: 13/13.
- PL lines (CCN × destination × cartons) match: 11/13. The 2 misses are decisions made in the email text (self pick-up; relabel + HOLD).
- CCN list matches: 10/13. The 3 misses are hand slips on the saved PLs.

The containers: CIMU1164134, SMCU1006072, TCNU3827209, TLLU4625944, SMCU1185050, CMAU8584800, BMOU5486746, TCKU9024600, CMAU6295351, TCNU1717326, FFAU7809134, CAIU4362993, ECMU5353380. The browser check uses several of them, not only CIMU1164134.

## Out of scope
- Pulling attachments straight from the mail archive by container number. That could come later on top of `combine`.
- OCR of scanned PDFs.
- Reading B/Ls, arrival notices or invoices beyond listing them as not used.
- Reading the email conversations (self pick-up / relabel / HOLD hints).
- The "Amazon Shipping List" customer-PL layout.
