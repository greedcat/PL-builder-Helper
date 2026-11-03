# Cartons → Shipments → Loads, Part 1 (records + reader, local only) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** For one container, the customer's load file is read into load → shipment → carton → carton item records in the local `fba_cartons` database, and shown read-only in a page of the FBA app.

**Architecture:**
- Pure reader modules in `backend/lib/` turn sheets (arrays of rows plus merges) into a neutral `Parsed` structure. They have no database and no I/O.
- An import service writes `Parsed` into the new tables beside the existing `shipments` (load) table, and rolls the totals up.
- A small route and a static page show the result.
- A regression script runs the reader over the 24 real containers from the history test.

Part 2 (a later plan) covers: moving shipments and cartons between loads (HOLD / relabel / self pick-up), the PL export, starting automatically when a container email arrives, and the move online.

**Tech Stack:** Node ≥ 22.12 ESM (`.mjs`), `node:test`, the app's `db` (Postgres via `DATABASE_URL`, SQLite in tests), and the app's zero-dependency `readXlsxSheets`.

**Spec:** `docs/superpowers/specs/2026-11-03-cartons-shipments-loads-design.md` (PL-builder-Helper repo). Read it first. The code goes in the **FBA_DATABASE** repo.

## Global Constraints

- **Never touch the live database or `fba_test`.** Every run uses `DATABASE_URL=postgres:///fba_cartons` (the local copy), or the SQLite temp DB in tests.
- **Work in a new git worktree** of FBA_DATABASE, branch `cartons-shipments-loads`, cut from `main`. The main checkout is on `invoices-truck-view` with someone's uncommitted work: leave it alone. Do not push without the user saying so.
- **The existing `shipments` table is the load.** Add columns to it; never rename it.
- **New tables:** `fba_shipments`, `cartons`, `carton_items`.
- **No placeholder cartons.** `cartons` rows exist only when the file has a carton list.
- **Estimated skids** = `Math.ceil(cbm / 1.7 - 0.05)` (`0` when there is no CBM).
- **Bonded** = the file no. contains `-SUF-` (case-insensitive). Loads key on (CCN#, destination) when bonded, on destination when general.
- **No new npm dependency.** `.xlsx` only, through `readXlsxSheets`. A `.xls` file is reported as "not read (.xls)" in Part 1. Decision for the user, flagged in the handoff: add SheetJS for `.xls`, or have the archive hand over `.xlsx`.
- **Customs-only columns are not stored:** HS code, value, material, function, brand, model, origin, picture, tax payment method.
- **Customer data stays out of git.** Real files go in `testdata/cartons/` (gitignored, like `testdata/pl/`). Committed fixtures are hand-made table rows: FBA IDs, POs, numbers and FC codes only; no names, addresses, emails or phones.
- **Match the house style:** block comments that explain *why*, `db.prepare(...).run/get/all`, `try { db.exec(...) } catch {}` migrations, routes returning `NEXT` when the path isn't theirs.

## Review Focus

1. **A load file re-imported for the same container** (customer sends a corrected file): must replace that container's shipments/cartons, not double them.
   → Task 7 test `reimport replaces, never doubles`.
2. **A sheet whose total row has text to its left** (`*Total（汇总）:` in a description column): must still stop there, and not count 734 again as a carton line.
   → Task 3 test `stops at a labelled total row`.
3. **A shipment with cartons in the file but different totals** (carton rows sum ≠ the shipment row): must keep the carton sum and flag it, never hide it.
   → Task 7 test `carton sum wins and is flagged`.
4. **Numbers written as text** ("1,060", "11.23 ", "") and a blank FBA cell under a merged one: parse the numbers; a merged blank is a continuation, not a new shipment.
   → Task 3 test `numbers as text and merged FBA`.
5. **A container whose attachments include another container's files** (one thread, several boxes): only this container's files are used.
   → Task 6 test `ignores other containers' files`.

---

## File Structure (FBA_DATABASE worktree)

| File | Responsibility |
|---|---|
| `backend/app/migrate.mjs` (modify) | new load columns + 3 tables |
| `backend/schema/schema.sql` (modify) | same tables for SQLite tests |
| `backend/lib/load-file/headers.mjs` | header keywords → column roles; `findHeader` |
| `backend/lib/load-file/read-sheet.mjs` | one sheet → `RawLine[]` (merges, totals, numbers) |
| `backend/lib/load-file/shape.mjs` | `RawLine[]` → shipments + cartons + items (continuation, itemised sums, FBA+PO split, mixed cartons) |
| `backend/lib/load-file/bonded.mjs` | loading list + customer PLs → CCN blocks (covered jobs, Self Pick Up, private address) |
| `backend/lib/load-file/pick.mjs` | choose the load file(s) among attachments by content |
| `backend/lib/load-file/index.mjs` | `parseLoadFiles(files, { containerNo, bonded }) → Parsed` |
| `backend/services/carton-import.mjs` | write `Parsed` to the DB, roll totals up |
| `backend/routes/cartons.mjs` | `POST /api/cartons/import`, `GET /api/cartons` |
| `frontend/cartons.html` + `frontend/app/cartons-view.js` | read-only load → shipment → carton page |
| `bin/cartons-history.mjs` | regression over `testdata/cartons/<box>/` |
| `tests/load-file-*.test.mjs`, `tests/carton-import.test.mjs`, `tests/fixtures/cartons/*.json` | tests and fixtures |

### Shared types (JSDoc in `backend/lib/load-file/index.mjs`)

```js
/** @typedef {{ fba:string|null, po:string|null, job:string|null, customerRef:string|null,
 *   destination:string|null, ccn:string|null, deliveryMethod:string|null, address:string|null,
 *   postal:string|null, appointment:string|null, windowStart:string|null, windowEnd:string|null, poExpiry:string|null,
 *   remark:string|null, cartons:number, totalPcs:number|null, kgs:number|null, nwKgs:number|null,
 *   cbm:number|null, cartonList:Carton[], source:{file:string, rows:number[]} }} Shipment */
/** @typedef {{ mark:string|null, kgs:number|null, nwKgs:number|null, cbm:number|null,
 *   l:number|null, w:number|null, h:number|null, pallet:string|null, tracking:string|null,
 *   loadingOrder:string|null, items:Item[] }} Carton */
/** @typedef {{ sku:string|null, fnsku:string|null, descEn:string|null, descCn:string|null, pcs:number|null }} Item */
/** @typedef {{ bonded:boolean, shipments:Shipment[], fileTotal:{cartons:number|null,cbm:number|null,kgs:number|null},
 *   files:{name:string, used:boolean, kind:string, note:string}[], checks:{name:string, ok:boolean, detail:string}[] }} Parsed */
```

---

### Task 1: Worktree, local server on `fba_cartons`, schema

**Files:**
- Modify: `backend/app/migrate.mjs` (append a block after the `bd_records` block)
- Modify: `backend/schema/schema.sql`
- Modify: `.claude/launch.json` (worktree copy) and `.gitignore` (`testdata/cartons/`)
- Test: `tests/carton-schema.test.mjs`

**Interfaces:**
- Produces:
  - `shipments` gains `ccn text, delivery_method text, delivery_address text, postal_code text, appointment_at text`.
  - Table `fba_shipments(id, load_id → shipments.id, container_id, job_no, fba_id, po_ref, customer_ref, window_start, window_end, po_expiry, remark, cartons int, total_pcs int, kgs real, nw_kgs real, cbm real, source_file, source_rows, carton_sum_mismatch int default 0, created_at)`.
  - Table `cartons(id, shipment_id → fba_shipments.id, load_id → shipments.id, mark, kgs, nw_kgs, cbm, l, w, h, pallet_no, tracking_no, loading_order)`.
  - Table `carton_items(id, carton_id → cartons.id, sku, fnsku, desc_en, desc_cn, pcs int)`.
  - Indexes on `load_id`, `container_id` and `shipment_id`.

- [ ] **Step 1:** Create the worktree with `superpowers:using-git-worktrees`: repo `/Users/mingyiwu/Desktop/github_hub/fba/FBA_DATABASE`, branch `cartons-shipments-loads` from `main`.
- [ ] **Step 2:** Write the failing test `tests/carton-schema.test.mjs`:
  - Use `tempDb('carton-schema')`, then import `backend/app/migrate.mjs`.
  - Insert one load, one `fba_shipments` row, one carton and two items.
  - Assert the join `carton_items → cartons → fba_shipments → shipments` returns 2 rows.
  - Assert `pragma table_info(shipments)` includes `ccn`.
- [ ] **Step 3:** Run `node --test tests/carton-schema.test.mjs`. Expected: FAIL (no such table `fba_shipments`).
- [ ] **Step 4:** Add the DDL to `migrate.mjs` in the house style:
  - `alter table shipments add column …` each in its own `try {} catch {}`;
  - `create table if not exists …`, with `IS_PG ? serial/bigserial : integer primary key autoincrement`, like `account_tokens`.
  - Mirror the tables in `schema.sql`.
- [ ] **Step 5:** Run the test. Expected: PASS.
- [ ] **Step 6:** Add `"cartons-local"` to the worktree's `.claude/launch.json`:
  - `runtimeExecutable: "sh"`;
  - args `-c "cd <worktree> && DATABASE_URL=postgres:///fba_cartons PORT=8789 exec node backend/boot.mjs"`;
  - port `8789`.

  Start it with `preview_start`, and confirm the boot log shows `engine=postgres` and no migration errors. Then run `psql -d fba_cartons -c '\d fba_shipments'` and confirm the table exists. Finally, confirm `psql -d fba_test -c '\d fba_shipments'` reports **no such relation**: `fba_test` must be untouched.
- [ ] **Step 7:** Add `testdata/cartons/` to `.gitignore`. Commit: `Cartons: load columns and fba_shipments / cartons / carton_items tables`.

### Task 2: Header roles

**Files:**
- Create: `backend/lib/load-file/headers.mjs`
- Test: `tests/load-file-headers.test.mjs`

**Interfaces:**
- Produces:
  - `norm(v) → string`: lower-case; spaces, `*`, full- and half-width brackets, colons, `.` and `#` removed.
  - `ROLE_ORDER: string[]`
  - `roleOf(header) → role|null`
  - `findHeader(rows, maxScan = 40) → { row:number, cols:Record<role, number> } | null`: the header row is the row with **fba** plus (**cartons** or **pcs**) that maps the most roles.
- Roles: `fba, po, job, customerRef, sku, fnsku, mark, cartons, pcs, totalPcs, kgsTotal, kgsEach, nwKgs, cbm, l, w, h, dims, destination, address, postal, deliveryMethod, appointment, windowStart, windowEnd, poExpiry, remark, flag, descEn, descCn, pallet, tracking, loadingOrder, ccn`.

- [ ] **Step 1:** Write the failing test. A table-driven `roleOf` test over **every header string in the spec's "Header names" table**, plus these:

  | Header | Role |
  |---|---|
  | `*Shipment ID` | `fba` |
  | `Amazon reference ID` | `po` |
  | `AMAZON ID` | `po` |
  | `运单号` | `job` |
  | `单号` | `job` |
  | `货代单号` | `customerRef` |
  | `Mark of CTN:箱唛` | `mark` |
  | `*Carton箱数` | `cartons` |
  | `*CTNS (件数)` | `cartons` |
  | `*Unit个数` | `pcs` |
  | `单箱产品数量* （PCS）` | `pcs` |
  | `产品总PSC数` | `totalPcs` |
  | `*G.W (KG)单箱毛重` | `kgsEach` |
  | `*Total G.W (KG)汇总毛重` | `kgsTotal` |
  | `*N.W.(KG) (净重)` | `nwKgs` |
  | `CBM汇总体积` | `cbm` |
  | `*L  (cm)长` | `l` |
  | `Carton Dimensions (尺寸)` | `dims` |
  | `派送地址代码（仓库代码）` | `destination` |
  | `进仓数据` | `destination` |
  | `派送地址` | `address` |
  | `派送邮编` | `postal` |
  | `尾程派送方式` | `deliveryMethod` |
  | `预约进仓时间` | `appointment` |
  | `货件配送起始日期` | `windowStart` |
  | `PO有效期` | `poExpiry` |
  | `战略客户标记` | `flag` |
  | `*H.S CODE 海关编码` | `null` |
  | `是否FBA` | `null` |
  | `柜号` | `null` |

  Then a `findHeader` test: rows with a title row, a blank row, then the TRHU header row → `{ row: 1 }`, with `cols.fba === 3` and `cols.cartons === 10`.
- [ ] **Step 2:** Run `node --test tests/load-file-headers.test.mjs`. Expected: FAIL (module not found).
- [ ] **Step 3:** Implement `headers.mjs`. `roleOf` checks an ordered list of `[role, (h) => boolean]`. The order matters:
  - `kgsTotal` before `kgsEach`;
  - `totalPcs` before `pcs`;
  - `fnsku` before `sku`;
  - `address` excludes headers containing `代码`;
  - `l`/`w`/`h` require `cm` or `长`/`宽`/`高` (so `weight` never becomes `w`).
- [ ] **Step 4:** Run the test. Expected: PASS.
- [ ] **Step 5:** Commit: `Load file reader: header roles`.

### Task 3: One sheet → raw lines

**Files:**
- Create: `backend/lib/load-file/read-sheet.mjs`
- Test: `tests/load-file-read-sheet.test.mjs`
- Fixtures: `tests/fixtures/cartons/sheet-*.json`, as `{ rows: string[][], merges: [r0,c0,r1,c1][] }` (0-based, inclusive)

**Interfaces:**
- Consumes: `findHeader`, `roleOf` (Task 2).
- Produces:
  - `readSheet({ rows, merges }) → { header:{row, cols}, lines: RawLine[], total: {cartons, cbm, kgs} | null } | null`
  - `RawLine = { row:number, v: Record<role, string|number|null>, cont: Set<role> }`. `cont` holds the roles whose cell is a merged continuation of the line above.
  - `toNum(v) → number|null`: strips `,` and spaces; `''` → null.

- [ ] **Step 1:** Write the failing tests, each on a small hand-made fixture:
  - `stops at a labelled total row`: lines `1,1` then `*Total（汇总）: | 2 | 0.17` → 2 lines; `total.cartons === 2`; `total.cbm === 0.17`.
  - `a row equal to the sum is not the total`: TLLU-shaped, with lines 8, 25, 2 (sum 35) then **35 with a description cell** → 4 lines.
  - `numbers as text and merged FBA`: `"1,060"` → 1060; `"11.23 "` → 11.23; an FBA cell merged over rows 2–3 → line 3 has `cont.has('fba')`.
  - `blank stretch ends the table`: two empty rows, then a notes block → lines end before the notes.
- [ ] **Step 2:** Run `node --test tests/load-file-read-sheet.test.mjs`. Expected: FAIL.
- [ ] **Step 3:** Implement `readSheet`. The total row is the first row that either:
  - matches `/total|汇总|合计/i` in any cell and has no FBA ID (`/FBA[0-9A-Z]{9,10}/`), or
  - has no fba/po/job/destination and **only numbers** left of the cartons column.

  Two empty rows end the table.
- [ ] **Step 4:** Run the test. Expected: PASS.
- [ ] **Step 5:** Commit: `Load file reader: one sheet to raw lines`.

### Task 4: Raw lines → shipments, cartons, items

**Files:**
- Create: `backend/lib/load-file/shape.mjs`
- Test: `tests/load-file-shape.test.mjs`

**Interfaces:**
- Consumes: `RawLine` (Task 3).
- Produces: `shapeLines(lines, { file, itemised }) → Shipment[]`, using the typedefs above.

Rules (from the spec and the prototype). Each one gets a test:

| Test name | Input lines | Expected |
|---|---|---|
| `carton list: one carton per line` | FBA19Q1XPK2Z, 13 lines with mark `…U000001`…`U000013`, pcs 6, kgsEach 11.23… | 1 shipment, `cartons 13`, `cartonList.length 13`, `totalPcs 78`, `kgs 146.02`, `cbm 1.0838` |
| `mixed carton: line without carton count is another SKU in the box above` | FBA19Q3QYJ6R U000001–U000003, then `sku COC12042402, pcs 5, cartons ''`, then U000004–6 | 6 cartons; carton 3 has 2 items (7 + 5); `totalPcs 66`; `cbm 0.7659` |
| `shipment level: no marks, no cartons list` | 分货明细 rows: FBA19QVWYSLT 77RP95XS 385 ctns 39.421 | `cartons 385`, `cartonList.length 0` |
| `blank FBA continues the one above` | 50 FBA19JFZ8X51, then 200 and 300 with no FBA | 1 shipment, 550 ctns |
| `itemised sheet sums per FBA` | `itemised: true`, FBA19JTNW2YD merged over 12 product lines | 1 shipment 189 ctns / 1929.14 kg / 13.99 CBM |
| `FBA+PO pairs in one cell split into shipments` | `FBA19F70FN08 4X5MC3NT-14件 / FBA19F71BWX5 3KOSMUYE-54件`, 68 ctns | 2 shipments: 14 and 54, each with its PO |
| `private address: non-Amazon FC` | destination `HWC`, FBA cell holds an address | `destination 'PA'`, `address` = that text, `fba null` |
| `dims cell splits into L/W/H` | `40*27.5*22.5` | `l 40, w 27.5, h 22.5` |
| `kgsEach × cartons when no total weight` | 2 cartons, 11.2 kg each, one line | `kgs 22.4` |

- [ ] **Step 1:** Write the tests in the table above.
- [ ] **Step 2:** Run `node --test tests/load-file-shape.test.mjs`. Expected: FAIL.
- [ ] **Step 3:** Implement `shapeLines`:
  - **Carton list:** the sheet has a `mark` column whose values look like `/FBA[0-9A-Z]{9,10}U\d{6}/`, or one line per carton (cartons = 1 on every line with a mark). Those lines become `Carton`s.
  - **Mixed carton:** a line with pcs but no carton count adds an `Item` to the previous carton.
  - **Destination:** an Amazon code is `/^[A-Z]{3}\d$/`; anything else becomes `PA`.
  - **`itemised`:** set by the caller when the header has `unit value` / `total value`.
- [ ] **Step 4:** Run the test. Expected: PASS.
- [ ] **Step 5:** Commit: `Load file reader: shipments, cartons and items`.

### Task 5: Bonded combine

**Files:**
- Create: `backend/lib/load-file/bonded.mjs`
- Test: `tests/load-file-bonded.test.mjs`

**Interfaces:**
- Consumes: `Shipment[]` from customer PLs (Task 4), and a loading-list sheet.
- Produces:
  - `readLoadingList({ rows }) → { box:string|null, jobs:{job, entry, pkgs, kg, cbm, fc, note}[], total:number|null } | null`
  - `combineBonded(customerPLs:{hbl, shipments}[], loadingList) → Shipment[]`: every shipment gets `ccn`.

Rules, from the spec and the bonded history test:
- A loading-list job is **covered** when its job number or 进仓编号 appears in any customer PL's `job`/`mark` text (the `+`-joined and `…QM` forms).
- An uncovered job becomes one shipment with no FBA. Its destination is the job's FC without `CA`; or `SELF PICKUP` when the FC is blank or the note has `自提`/`自清`; or `PA` for a non-Amazon FC.
- A customer PL's HBL comes from its file name; otherwise from the loading-list job whose cartons equal the PL's total.

Tests:
- `CIMU1164134`: one PL, HBL EJFBTOR26S1870, 7 lines, 1060 ctns; loading list of 7 jobs, total 1077.
  - CCNs = `EJFBTOR26S1870`, `EJTOR26S0097`, `EJTOR26S0099`.
  - 0097 = SELF PICKUP 2 / 2930 / 1.39; 0099 = SELF PICKUP 15.
  - Cartons sum 1077.
- `CMAU8584800`: the 1818 PL's 单号 `EJFBTOR26S1796QM` covers job 1796 → no extra 1796 shipment; cartons sum 959.

- [ ] **Step 1:** Write the two tests, with fixtures made from those containers' table cells (FBA IDs, POs, numbers and FC codes only).
- [ ] **Step 2:** Run `node --test tests/load-file-bonded.test.mjs`. Expected: FAIL.
- [ ] **Step 3:** Implement `readLoadingList` and `combineBonded`.
- [ ] **Step 4:** Run the test. Expected: PASS.
- [ ] **Step 5:** Commit: `Load file reader: bonded combine`.

### Task 6: Pick the load file(s) and the entry point

**Files:**
- Create: `backend/lib/load-file/pick.mjs`, `backend/lib/load-file/index.mjs`
- Test: `tests/load-file-pick.test.mjs`

**Interfaces:**
- Consumes: Tasks 2–5, and `readXlsxSheets(buf)` from `backend/lib/pl_parser.mjs`, which returns sheets as rows plus merges. Check its exact return shape and adapt it to `{ rows, merges }` in `index.mjs`.
- Produces:
  - `parseLoadFiles(files:{name:string, buf:Buffer, fromMaple?:boolean}[], { containerNo:string, fileNo:string }) → Parsed`
  - `scoreSheet({ rows }) → { fbaIds:number, fcCodes:number }`

Rules:
- Skip files from Maple (`fromMaple`) and non-`.xlsx` files. A `.xls` is listed with the note `not read (.xls)`.
- **General:** use the sheet(s) with FBA IDs, preferring the most. If there are several, a sheet whose FBA IDs are a subset of a chosen one is skipped (per-destination copies).
- **Bonded** (`-SUF-`): customer PLs plus the loading list whose `box` equals `containerNo`. A loading list for another box is ignored, as is a customer PL whose HBL isn't a job on this box's loading list.
- `checks`:
  - `cartons equal the file's total row`;
  - bonded: `cartons equal the loading-list total`.

Tests:
- `general: picks the sheet with FBA IDs`: one 分货明细-shaped sheet plus a BOL-shaped sheet (no FBA header) → `files[0].used`, 15 shipments, 734 cartons, `checks[0].ok`.
- `ignores other containers' files`: two loading lists (box `TCKU9024600` and another) and a PL whose HBL belongs to the other → only this box's shipments.
- `.xls is reported, not read`.

- [ ] **Step 1:** Write the three tests.
- [ ] **Step 2:** Run `node --test tests/load-file-pick.test.mjs`. Expected: FAIL.
- [ ] **Step 3:** Implement `pick.mjs` and `index.mjs`.
- [ ] **Step 4:** Run the test. Expected: PASS.
- [ ] **Step 5:** Commit: `Load file reader: pick the load file, parseLoadFiles`.

### Task 7: Write the records

**Files:**
- Create: `backend/services/carton-import.mjs`
- Test: `tests/carton-import.test.mjs`

**Interfaces:**
- Consumes: `Parsed` (Task 6), and the tables (Task 1).
- Produces:
  - `importParsed(containerId:number, parsed:Parsed, { by:string }) → { loads:number, shipments:number, cartons:number, flags:string[] }`
  - `containerRecords(containerId) → { loads:[{…load, shipments:[{…, cartons:[{…, items}]}]}] }`

Rules:
- Everything runs in one transaction (`withTx`).
- **Load per (ccn, destination)** for bonded, **per destination** for general:
  - reuse an existing `shipments` row of this container with the same `destination_code` (and `ccn`);
  - otherwise insert one, with `file_no` = container file no. + next `-NN`, `source 'carton-import'`.
- **Re-import:** delete this container's `fba_shipments`, `cartons` and `carton_items` first. Existing loads are kept and updated.
- **Load totals** are written from the sums: `total_cartons`, `weight_kg`, `cbm`, `est_pallets = ceil(cbm/1.7 − 0.05)`, `fba_id` / `po_reference` (space-joined in shipment order).
- **Shipment totals:** when `cartonList` has cartons, totals = carton sums. If those differ from the file's shipment numbers, set `carton_sum_mismatch = 1` and add a flag `"<FBA> cartons/CBM differ from the file's shipment line"`.

Tests (SQLite temp DB):
- `general import`: TRHU-shaped parsed → 9 loads, 15 shipments, 19 cartons (the two sample shipments), YYZ7 load `total_cartons 73`, `est_pallets` per the rule.
- `reimport replaces, never doubles`: import twice → still 15 shipments / 19 cartons; load count unchanged.
- `carton sum wins and is flagged`: a shipment line says 14 ctns, but its cartons sum to 13 → shipment cartons 13, mismatch 1, one flag.
- `bonded import keys loads by CCN and destination`: CIMU-shaped → 5 loads (1870×YOO1/YYZ4/YYZ7, 0097 SELF PICKUP, 0099 SELF PICKUP), `ccn` set on each.

- [ ] **Step 1:** Write the four tests.
- [ ] **Step 2:** Run `node --test tests/carton-import.test.mjs`. Expected: FAIL.
- [ ] **Step 3:** Implement `importParsed` and `containerRecords`.
- [ ] **Step 4:** Run the test. Expected: PASS.
- [ ] **Step 5:** Commit: `Cartons: write load file records, totals rolled up`.

### Task 8: Route, page, and the local check on real containers

**Files:**
- Create: `backend/routes/cartons.mjs`, `frontend/cartons.html`, `frontend/app/cartons-view.js`, `bin/cartons-history.mjs`
- Modify: `backend/app/handler.mjs`: register `cartonsRoutes` next to `loadsRoutes`.
- Test: `tests/cartons-route.test.mjs`

**Interfaces:**
- Consumes: `parseLoadFiles` and `importParsed` / `containerRecords`.
- Produces:
  - `POST /api/cartons/import?container=<container_no>`: body is multipart or a JSON `{ files:[{name, base64}] }` (pick the simpler of the two given `http.mjs`). Answers `{ parsed:{files, checks}, result }`. **Admin only.**
  - `GET /api/cartons?container=<container_no>` → `containerRecords`.
  - `bin/cartons-history.mjs [--db]`: for each `testdata/cartons/<box>/`, it reads `mail/*` and `fba_loads.tsv`, runs `parseLoadFiles`, and prints per box: shipments, cartons vs the file total, FBA–PO pairs found in the app, and destinations matching the app's cartons. It ends with totals, like the prototype's scoreboard. `--db` also imports into `fba_cartons`.

- [ ] **Step 1:** Write the failing route test:
  - a non-admin → 403;
  - admin import of the TRHU-shaped fixture → 200 with `result.shipments 15`;
  - GET → 9 loads.
- [ ] **Step 2:** Run `node --test tests/cartons-route.test.mjs`. Expected: FAIL.
- [ ] **Step 3:** Implement the route. The page is the prototype's carton page (`/Users/mingyiwu/…/scratchpad/testpage/carton_template.html`) ported to the app's style:
  - destination bars → shipment rows (click to open) → carton rows with mixed-SKU rows;
  - the user's 16 columns;
  - every cell bordered and centred;
  - the container picker at the top.
- [ ] **Step 4:** Run the test. Expected: PASS. Then run all tests with `node --test tests/`. Expected: PASS, and no existing test newly failing (note any that already fail on `main`).
- [ ] **Step 5:** Copy the history data into the worktree's `testdata/cartons/`. Source: the PL-builder session scratchpad `hist/` and `gen/` folders (`<box>/mail/*` and `fba_loads.tsv` each).
  - Run `node bin/cartons-history.mjs`. Expected at least:
    - general: 9/9 load files found, all destinations matching the app on TRHU6779290 and the 8 others whose file was found;
    - bonded: cartons equal the loading-list total 13/13.

    Paste the scoreboard into the commit message.
  - Run `DATABASE_URL=postgres:///fba_cartons node bin/cartons-history.mjs --db`, then open http://localhost:8789/cartons.html?container=TRHU6779290 with `preview_start cartons-local` and take a screenshot.
- [ ] **Step 6:** Commit: `Cartons: import route, read-only page, history check on fba_cartons`.

---

## Handoff notes for the user (after Part 1)
- `.xls` files (some agents' loading lists and load files) are listed as "not read" until a choice is made: add SheetJS (one dependency), or have the mail archive convert them.
- Part 2 plan: moving shipments/cartons between loads (HOLD / relabel / self pick-up), the PL export from the records, starting automatically when a container email arrives, and the move online. The move online happens only when the user says so.
