# Build Space, Part 1 (records, reader, Build page: Files + Records; local only) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** In the FBA app's new **Build** page, ops pick a container from the to-do list, confirm its load file(s) from the email archive, read them into load → shipment → carton → item records in the local `fba_cartons` database, then review, move, split, edit and confirm those records.

**Architecture:**
- Pure reader modules in `backend/lib/` turn sheets (arrays of rows plus merges) into a neutral `Parsed` structure. They have no database and no I/O.
- An import service writes `Parsed` into the new tables beside the existing `shipments` (load) table, and rolls the totals up.
- Build routes list containers to build, fetch attachments through the existing mail-archive client (`workerFetch`), read the confirmed files, and apply moves and edits (all logged in `edit_history`).
- The Build page has three parts: the to-do list, ① Files and ② Records.
- A regression script runs the reader over the 24 real containers from the history test.

Part 2 (a later plan) covers: ③ PL (preview, export with ExcelJS through the PL builder's writer logic, save to the SharePoint test folder, "PL out of date"), the PL check against saved PLs, and the move online.

**Tech Stack:** Node ≥ 22.12 ESM (`.mjs`), `node:test`, the app's `db` (Postgres via `DATABASE_URL`, SQLite in tests), and the app's zero-dependency `readXlsxSheets`.

**Spec:** `docs/superpowers/specs/2026-11-03-build-space-design.md`, plus the parts of `docs/superpowers/specs/2026-11-03-cartons-shipments-loads-design.md` it keeps (Records, Bonded or general, Header names, Evidence). Both are in the PL-builder-Helper repo; read both first. The code goes in the **FBA_DATABASE** repo.

## Global Constraints

- **Never touch the live database or `fba_test`.** Every run uses `DATABASE_URL=postgres:///fba_cartons` (the local copy), or the SQLite temp DB in tests.
- **Work in a new git worktree** of FBA_DATABASE, branch `cartons-shipments-loads`, cut from **`release-was-hold`**. That is the live line, with the mail-archive client; `main` is two commits old. The main checkout is on `invoices-truck-view` with someone's uncommitted work: leave it alone. Do not push without the user saying so.
- **Mail archive:** read only, through `workerFetch(path, { att })` in `backend/app/mail-archive.mjs`:
  - `GET /containers/<container id>/threads?possible=1` lists the threads;
  - `GET /threads/<id>` returns a thread with its attachments `{ id, post_id, name, type, size, sha256, link_only }`;
  - `GET /att/<sha256>` returns the bytes.
  Skip `link_only`. A post `from_addr` ending `@maplesupplychain.ca` is Maple's own.
- **Nothing is read until ops press Read** (spec ① Files).
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
| `backend/services/carton-import.mjs` | write `Parsed` to the DB, roll totals up; move / split / edit / confirm |
| `backend/routes/build.mjs` | `/api/build/*`: to-do list, files (via the mail archive), read, records, move/edit/confirm |
| `frontend/app/views/build.js` + `build-records.js` | the Build page: to-do list, ① Files, ② Records |
| `bin/cartons-history.mjs` | regression over `testdata/cartons/<box>/` |
| `tests/load-file-*.test.mjs`, `tests/carton-import.test.mjs`, `tests/build-route.test.mjs`, `tests/build-view.test.mjs`, `tests/fixtures/cartons/*.json` | tests and fixtures |

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
  - `containers` gains:
    - `build_state text`: `null` (to build) / `'read'` (in progress) / `'confirmed'` / `'built'`;
    - `build_by text, build_at text`;
    - `records_changed_at text, pl_exported_at text` (Part 2 uses the last two for "PL out of date").

- [ ] **Step 1:** Create the worktree with `superpowers:using-git-worktrees`: repo `/Users/mingyiwu/Desktop/github_hub/fba/FBA_DATABASE`, branch `cartons-shipments-loads` from **`release-was-hold`**.
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
  - `containerRecords(containerId) → { container, bonded, loads:[{…load, shipments:[{…, cartons:[{…, items}]}]}], checks:{name, ok, detail}[], handFixes:number }`
  - `moveShipment(shipmentId, { toLoadId? , toDestination?, toCcn?, count?, cartonIds? }, { by }) → { fromLoadId, toLoadId, newShipmentId? }`
    - `count` (no carton list) splits the shipment into two rows with the same FBA–PO;
    - `cartonIds` moves those cartons, with their shipment row duplicated into the target load when needed;
    - `toDestination` finds or makes the load (`HOLD`, `SELF PICKUP`, an FC, `PA`).
  - `editShipment(shipmentId, patch, { by })` / `editCarton(cartonId, patch, { by })` / `addShipment(loadId, values, { by })` / `removeShipment(shipmentId, { by })`
  - `confirmRecords(containerId, { by })`: writes load totals + `est_pallets`, sets `build_state 'confirmed'`.

  Every change writes one `edit_history` row (subject = container no., kind `build`), recomputes the affected loads' totals, and stamps `records_changed_at`.

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
- `move to HOLD by count splits`: FBA19QVWYSLT 385 → move 20 to `HOLD` → YOO1 load 365, HOLD load 20, two rows with the same FBA–PO; one `edit_history` row.
- `move chosen cartons`: FBA19Q3QYJ6R, move cartons U000005–6 to `SELF PICKUP` → XYY4 load drops 2 cartons and 0.2001 CBM; the mixed carton's items stay with their carton.
- `edit updates totals`: set a shipment's cartons 13 → 12 → its load `total_cartons` drops by 1.
- `handFixes counted for read again`: after two edits, `containerRecords(...).handFixes === 2`.
- `confirm writes est. skids`: YYZ7 6.1677 CBM → `est_pallets 4`; `build_state 'confirmed'`.

- [ ] **Step 1:** Write the nine tests.
- [ ] **Step 2:** Run `node --test tests/carton-import.test.mjs`. Expected: FAIL.
- [ ] **Step 3:** Implement the functions in Produces. Each one runs in `withTx`.
- [ ] **Step 4:** Run the test. Expected: PASS.
- [ ] **Step 5:** Commit: `Cartons: write records; move, split, edit, confirm`.

### Task 8: Build routes

**Files:**
- Create: `backend/routes/build.mjs`
- Modify: `backend/app/handler.mjs`: register `buildRoutes` next to `mailRoutes`.
- Test: `tests/build-route.test.mjs` (stub `workerFetch` by passing a fake through `buildRoutes.__setArchive(fn)`, exported for tests)

**Interfaces:**
- Consumes: `workerFetch`, `parseLoadFiles` (Task 6), and the Task 7 functions.
- Produces (all `/api/build/*`; they need a signed-in user, and the writes need the `loads_edit` capability or whichever the Loads grid's edits use; check `permissions.mjs` and reuse it):

  | Route | Answers |
  |---|---|
  | `GET /api/build/list?tab=todo\|progress\|built&q=` | `{ rows:[{ container_id, container_no, file_no, client, bonded, eta, build_state, build_by, build_at, files:number, flag }] }`. `todo` = has archive threads, `build_state` null, `load_type='container'`; ETA order |
  | `GET /api/build/files?container_id=` | `{ files:[{ sha256, name, from, at, ours:boolean, readable:boolean, score:{fbaIds, fcCodes, cartons}, kind:'load'\|'loadingList'\|'customerPL'\|'other', ticked:boolean, why:string }] }`. `ticked` is the pick of Task 6 |
  | `POST /api/build/read` `{ container_id, sha256s:[], uploads:[{name, base64}] }` | runs `parseLoadFiles` + `importParsed`; `{ result, files, checks }`; sets `build_state 'read'`. If records already exist and `confirm_replace` is not `true` → 409 `{ handFixes }` |
  | `GET /api/build/records?container_id=` | `containerRecords` |
  | `POST /api/build/move`, `/edit`, `/add`, `/remove`, `/confirm` | the Task 7 functions; each answers the fresh `containerRecords` |

- [ ] **Step 1:** Write the failing tests:
  - `list shows a container with emails and no records under todo`;
  - `files pre-ticks the 分货明细 sheet and folds Maple's own replies` (a fake archive of one customer xlsx, one BOL xlsx, one `@maplesupplychain.ca` reply);
  - `read then read again asks first` (second read → 409 with `handFixes`; with `confirm_replace` → 200);
  - `a user without edit rights gets 403 on move`.
- [ ] **Step 2:** Run `node --test tests/build-route.test.mjs`. Expected: FAIL.
- [ ] **Step 3:** Implement `build.mjs`. The archive container id is `containers.id`. Bytes come from `workerFetch('/att/' + sha, { att: true })`.
- [ ] **Step 4:** Run the test. Expected: PASS.
- [ ] **Step 5:** Commit: `Build: routes for the to-do list, files, read, records and changes`.

### Task 9: The Build page (to-do list, ① Files, ② Records)

**Files:**
- Create: `frontend/app/views/build.js` (view module, the same shape as the other `views/*.js`), `frontend/app/views/build-records.js` (the records table and Move menu)
- Modify: the view registry and menu where the other views are registered (`frontend/app/views/base.js` / `app.js`; follow how `invoices` is added); the Loads grid row menu: "Open in Build" → `#build/<container_id>`.
- Test: `tests/build-view.test.mjs`: pure helpers loaded with `node:vm`, as in `tests/invoices-truck-view.test.mjs`.

**Interfaces:**
- Consumes: the Task 8 routes.
- Produces:
  - pure helpers `recordRows(records) → row[]`, with each row `{ kind:'load'|'shipment'|'carton'|'item', cells:Record<col, value>, … }`;
  - `BUILD_COLS`: the 16 user columns in this order: Job #, FBA#, PO#, Total pcs, Pcs, Cartons, Kgs, CBM, Destination, Remark, Shipping mark, SKU, FNSKU, L, W, H;
  - `checkBadges(checks)`.

Look (from the spec and the user's earlier asks):
- the container list at the top or left with the three tabs and counts;
- the three steps ① ② ③ (③ greyed "Part 2");
- every table cell bordered and **centred**;
- destination bars in blue, shipment rows in purple (click to open cartons), mixed-SKU rows in yellow;
- `—` for a value the file didn't have.

- [ ] **Step 1:** Write the failing tests:
  - `BUILD_COLS order is the user's 16`;
  - `recordRows nests load → shipment → carton → item` (TRHU sample: YYZ7 bar, FBA19Q1XPK2Z row, 13 carton rows);
  - `a mixed item row follows its carton`.
- [ ] **Step 2:** Run `node --test tests/build-view.test.mjs`. Expected: FAIL.
- [ ] **Step 3:** Implement the view.
  - ① Files: the list, with ticks; "not a load file" folded; + Add a file; **Read**. The 409 answer shows "This replaces the records and N hand fixes. Replace?"
  - ② Records: the table, check badges, **Move ▾** (to load / HOLD / SELF PICKUP / new FC / part of it…), cell edit, + Add / ⋯ Remove, **Confirm records**.
- [ ] **Step 4:** Run the test. Expected: PASS.
- [ ] **Step 5:** Browser check on `cartons-local` (port 8789, `fba_cartons`):
  1. Open Build; TRHU6779290 is under To build.
  2. ① shows `DSF260920865_分货明细.xlsx` ticked with "15 FBA · 9 FC · 734 ctns".
  3. Read → ② shows 9 loads, with the check "734 = file total" ✓.
  4. Move 20 of FBA19QVWYSLT to HOLD → YOO1 365, HOLD 20.
  5. Confirm → YYZ7 `est_pallets 4` (`psql -d fba_cartons`).

  Take screenshots. Confirm `fba_test` has no `fba_shipments` table.
- [ ] **Step 6:** Commit: `Build: the Build page — to-do list, files, records`.

### Task 10: History check on real containers

**Files:**
- Create: `bin/cartons-history.mjs`

**Interfaces:**
- Consumes: `parseLoadFiles`, `importParsed`.
- Produces: `node bin/cartons-history.mjs [--db]`. For each `testdata/cartons/<box>/` (`mail/*` + `fba_loads.tsv`), it prints:
  - shipments;
  - cartons vs the file total (bonded: vs the loading-list total);
  - FBA–PO pairs found in the app;
  - destinations whose cartons match the app.

  It ends with totals. `--db` also imports into the database in `DATABASE_URL` (only ever `fba_cartons`; refuse any other name).

- [ ] **Step 1:** Copy the history data into the worktree's `testdata/cartons/` (gitignored). Source: the PL-builder session scratchpad's `hist/` and `gen/` folders (`<box>/mail/*`, `fba_loads.tsv`).
- [ ] **Step 2:** Run `node bin/cartons-history.mjs`. Expected at least:
  - bonded: cartons equal the loading-list total 13/13;
  - FBA–PO pairs 96/105, or better;
  - general: the load file is found for the 9 whose file is in the archive, and TRHU6779290's 9 destinations all match the app.

  `.xls`-only boxes are reported "not read (.xls)", not as failures.
- [ ] **Step 3:** Run `DATABASE_URL=postgres:///fba_cartons node bin/cartons-history.mjs --db`, and spot-check two containers in the Build page.
- [ ] **Step 4:** Commit with the scoreboard in the message: `Cartons: history check over 24 containers`.

---

## Handoff notes for the user (after Part 1)
- `.xls` files (some agents' loading lists and load files) are listed as "not read" until a choice is made: add SheetJS (one dependency), or have the mail archive convert them.
- Part 2 plan: ③ PL in the Build page:
  - preview, and Export via ExcelJS (pinned) with the PL builder's `public/js/writer.js` logic ported;
  - Cartons sheet + the original file sheet;
  - save to the SharePoint **test** folder;
  - "PL out of date";
  - the PL check against the 13 saved bonded PLs.

  Then the move online, only when the user says so, after comparing the live copy with the branch.
