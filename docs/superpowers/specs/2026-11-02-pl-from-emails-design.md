# PL from email attachments: design

Date: 2026-11-02. Status: approved in chat; waiting for review of this written spec.

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
- Estimated skids keep the builder's rule `ceil(CBM/1.7 − 0.05)`. The 17 for YYZ7 on the correct PL was a hand edit. The builder will show 18.

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
5. If any EMF says "In bond", the Bonded switch is turned on, and the CCN# list is filled in the combined table's order.

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
- `readLoadingList(aoa) → { jobs: [{job, entry, pkgs, kg, cbm, fc, note}], total }`
  - `fc` is 进仓数据 with a leading `CA` removed.
  - The TOTAL row gives `total`.
- `readEmf(text) → { inBond, ccns: [...] }`
  - `inBond`: the text has "In bond".
  - `ccns`: the lines under "RELATED CCN", with the 4-character carrier code (`8FXN`) removed.
- `combine({ customerPLs, loadingList, emfs }) → { headers, rows, merges, ccnOrder, bonded, check }`
  - **Headers.** The first customer PL's header row, plus `CCN#`.
    - With no customer PL, it is a minimal header: `CTNS, G.W.(KG), CBM, FBA No., Reference ID, Amazon FC, 单号, 箱唛, CCN#`.
  - **Customer-PL rows.** Copied as-is, with CCN# = that PL's HBL.
  - **A blank FC on a customer-PL row:**
    - if the row's 单号 matches a loading-list job, that job's FC;
    - otherwise, left blank, merged under the FC above (see merges).
  - **Loading-list jobs that are not a customer PL's HBL** (and are not a sub-job already listed in a customer PL's 单号 column): one row each.
    - CTNS / G.W. / CBM come from 件数 / 重量 / 体积, and CCN# = the job.
    - FC is the job's FC. If the FC is blank, or the note has 自提 or 自清, FC = `Self Pick Up`.
    - In the CIMU example, 1830 / 1831 / 1846 / 1876 appear in 1870's 单号 column, so they are not added again. 0097 and 0099 are added.
  - **Merges.** In FC, 单号 and 箱唛, a blank cell directly under a filled cell in the same CCN block is merged into it. The source file's merges are also kept for customer-PL rows.
  - **Order.** CCN blocks: FBA HBLs (those with a customer PL) first, then the others. Within each group, the EMF RELATED CCN order, else the loading-list order.
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
- Existing tests still pass.
- Browser check: drop the CIMU files on the local server.
  - The preview should match the correct PL's rows and merges.
  - Summary: YOO1 355/14, YYZ4 140/2, YYZ7 565/18, 0097 2/1, 0099 15/6.

## Out of scope
- Pulling attachments straight from the mail archive by container number. That could come later on top of `combine`.
- OCR of scanned PDFs.
- Reading B/Ls, arrival notices or invoices beyond listing them as not used.
