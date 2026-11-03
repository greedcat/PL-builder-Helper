# Cartons → shipments → loads: design

Date: 2026-11-03. Status: structure approved in chat; waiting for review of this written spec.

This spec supersedes `2026-11-02-pl-from-emails-design.md`. That spec built the PL in the browser from dropped files. The user then changed the flow to **customer files → records in the system → PL export**. The reading rules it learned (header detection, merges, total rows, private addresses, covered jobs) carry over unchanged.

## Goal

Every carton has its own record. When a container's emails arrive:
1. The system reads the customer's load file into **loads → shipments → cartons**.
2. Ops review and adjust the records: HOLD, relabel, self pick-up, split.
3. The PL is exported from the records.

The FBA app's load view and the PL read the same records, so they cannot disagree.

## What the user decided
- The structure is **cartons → shipments → loads**. Each carton has its own record (2026-11-03).
- The system record columns (user's list): Job #, FBA#, PO#, total pcs, pcs, cartons, kgs, CBM, destination, remark, shipping mark, SKU, FNSKU, L, W, H. The "B" columns from the header census are added (see Fields).
- Three adjustments were approved:
  1. **No placeholder cartons.** Carton records are created only from real carton data (the customer's carton list, or the warehouse scanning or labelling at destuffing). Until then the shipment holds the file's totals.
  2. **A carton points to its load as well as its shipment**, so a shipment can be split across loads (part HOLD, part relabelled, part self pick-up).
  3. **The FBA app's existing `shipments` table stays the load table.** New tables go beside it.
- Estimated skids keep the builder's rule `ceil(CBM/1.7 − 0.05)`.
- Decisions made in email text (self pick-up, relabel, HOLD) are applied by ops in the review screen, not read automatically.

## Where it is built (user, 2026-11-03)
- **Neither the live database nor the working test copy is touched.**
- Everything is built and tested on a **separate local database, `fba_cartons`**:
  - it is a dump/restore copy of `fba_test` (2,385 containers / 11,024 loads at copy time);
  - the new tables and columns are added only there.
- Moving to the online database happens only after the whole thing is done, tested, and the user says so.

## Bonded or general
- **Bonded:** the file no. contains `-SUF-` (e.g. `2018-SUF-26063`), or the BD record has `if_bonded`. Loads are per **CCN# × destination**. The load file is the customer PL(s) (派送资料 / 清关资料), plus the 拼箱配载清单 and the EMF CLOSE PDFs. The rules are in the superseded spec, "Components", and were tested on 13 containers.
- **General:** loads are per **destination**. The load file is **one Excel file from the agent or customer**, found by its **content** (FBA IDs + Amazon FC codes + cartons), never by its name. Names seen: 派送指引, 派送指示, 货物明细, 出库, 操作指令, 分货明细, FBA TRACKING REPORT, `<container>--<ctns>.xls`. The 清关资料 is **not** the load file for general containers: it has no FBA IDs. It was tested on 11 containers from 11 clients: 9 had the file in the archive and 2 did not.

## Records

```
Load ─< Shipment ─< Carton ─< Carton item
          (FBA# + PO#)   (one box)   (SKU line)
Carton ──> Load (default: its shipment's load; changed when cartons are split off)
```

### Load: the existing `shipments` table (one per container × CCN# × destination)
| Field | Note |
|---|---|
| file_no + file_seq, container | existing |
| **ccn** | **new**; bonded only |
| destination_code | existing: FC code / PA / SELF PICKUP / HOLD |
| **delivery_method** | **new**: truck / express / UPS / self pick-up (派送方式, 尾程派送方式, 渠道, Delivery instruction) |
| **delivery_address, postal_code** | **new** (派送地址, 邮编, FBA address); needed for PA |
| **appointment_at** | **new** (预约时间, 预约进仓时间) |
| est_pallets, act_pallets_* | existing; est = rule above, from the summed CBM |
| total_cartons, weight_kg, cbm, fba_id, po_reference | existing; **written from the sum of its shipments/cartons**, no longer typed |

### Shipment: new table `fba_shipments` (one per FBA# + PO# within a load)
| Field | Note |
|---|---|
| load_id | the load it was read into |
| job_no | 单号 / 运单号 / 业务编号 |
| fba_id, po_ref | FBA# and PO# (Reference ID). A cell holding several "FBA… REF-n件" is split into one shipment each, with its own cartons |
| customer_ref | 货代单号 |
| window_start, window_end / po_expiry | 货件配送起始/结束日期, PO有效期 |
| remark | 备注; 战略客户标记 also goes here |
| cartons, total_pcs, kgs, nw_kgs, cbm | **from the file** when there are no carton records; **the sum of its cartons** once cartons exist |
| source_file, source_rows | where it was read from |

### Carton: new table `cartons` (one per physical box; only from real carton data)
| Field | Note |
|---|---|
| shipment_id, load_id | load_id defaults to the shipment's load |
| carton_id / shipping_mark | e.g. `FBA19Q1XPK2ZU000001` (箱唛 / Mark of CTN) |
| kgs, nw_kgs, cbm, l, w, h | `kgs` is this box's gross weight. A file's per-carton weight (单箱毛重) is stored as is; a line covering n cartons is stored as n cartons |
| pallet_no, tracking_no, loading_order | 托盘编号, 追踪号, 装柜顺序 |

### Carton item: new table `carton_items` (one per SKU line in a box)
| Field | Note |
|---|---|
| carton_id | |
| sku, fnsku, description_en, description_cn, pcs | A file line with no carton number is another SKU in the carton above it (a mixed carton). It adds pcs, not cartons |

Totals only add upward: carton items → carton → shipment → load → container.

## Flow
1. **The email arrives.** It is matched to the container (the archive already does this). Bonded or general comes from the file no. / BD record.
2. **Find the load file(s)** among the container's attachments, by content. Maple's own replies (offload report, invoices) are skipped. When none is found, the container is flagged "no load file". In the history test this happened for 2 of 11 general containers.
3. **Read** with the rules from the prototype, generalised:
   - columns found by header keywords (see "Header names" below);
   - merged cells: a continuation is the same FBA / destination / job; numbers are counted once;
   - lines with no FBA No. continue the FBA above;
   - an itemised 清关资料 is summed per FBA;
   - the total row is recognised;
   - private addresses (`HWC`, non-Amazon FC) become PA;
   - for bonded containers, jobs covered via 单号/箱唛/进仓编号 are not added again.
4. **Review screen** (FBA app): loads → shipments → cartons, with checks:
   - cartons equal the file's total row;
   - bonded: cartons equal the loading-list total.
   Ops apply HOLD / relabel / self pick-up by moving shipments or cartons to another load, then confirm.
5. **Save** the records. Load totals and estimated skids are written from the sums.
6. **PL export** reads the records:
   - bonded: one block per CCN# with its shipment lines, FC/单号/箱唛 merged down, and a summary per CCN# × destination;
   - general: the same without CCN# blocks.
   It replaces typing the PL in the builder.

## Header names the reader must know (from 31 real files)
| Column | Headers seen |
|---|---|
| kgs | G.W., 毛重, 重量, 货物重量, 实际重量, 单箱重量, 汇总毛重, KGS #gwTtl |
| destination | Amazon FC, 仓点, 仓库代码, FBA仓库编码, FBA仓库 简称, 目的仓库, 仓库地址代码, 进仓数据 (CA+FC) |
| PO# | Reference ID, AMAZON ID, PO |
| pcs / total pcs | QTY, Unit 个数, 单箱产品数量, PCS（商品数量） / 产品总PSC数 |
| L/W/H | 长/宽/高 (CM), 货箱长度/宽度/高度; or one "Carton Dimensions" cell "40*27.5*22.5", split |
| shipping mark | 箱唛, Mark of CTN, 分货标识, WAREHOUSE ENTRY NO/MARKS |
| SKU | SKU, 产品SKU |
| description | DESCRIPTION OF GOODS, 英文品名, 中文品名, 品名, 产品, CARGO DESCRIPTION |

Customs-only columns are **not** stored: HS code, value, material, function, brand, model, origin, picture, tax payment method.

## Evidence (2026-11-02/03, throwaway prototype)
- **Bonded, 13 containers** (E-Cargo Master):
  - 214/217 PL lines read exactly (cartons, kg, CBM, FBA No.);
  - 96/105 FBA–PO pairs in the FBA app found;
  - cartons equal the loading-list total 13/13.
- **General, 11 containers / 11 clients:** the load file was found by content in 9.
- **TRHU6779290:** `分货明细` read as 15 shipments / 734 ctns / 72.0836 CBM, equal to its total row. All 9 destinations equal the FBA app.
  - Carton sample: FBA19Q1XPK2Z (13 cartons) and FBA19Q3QYJ6R (6 cartons, 1 mixed) add up exactly to their shipments.
- Test pages (local only; they hold customer data): http://localhost:8765 and /carton.html.

## Testing
- Reader unit tests on fixtures made from the table parts of real files, without addresses, emails or phone numbers:
  - CIMU1164134 (bonded basics);
  - CMAU8584800 (private addresses, a job covered via 箱唛);
  - TLLU4625944 (the total-row trap);
  - TCNU3827209 (itemised 清关资料);
  - TCKU9024600 (other containers' files mixed in);
  - FFAU7809134 (FBA+PO in one cell);
  - TRHU6779290 (carton-level, mixed carton).
- History regression: the comparison script runs over the 24 containers. Each change must not lower line / pair / carton-total matches.
- Database: totals stay equal to the sums after moving cartons between loads.

## Out of scope (for now)
- Reading email text for HOLD / relabel / self pick-up hints.
- OCR of scanned PDFs.
- The "Amazon Shipping List" customer-PL layout (1 of 31 files).
- Back-filling old containers: they keep loads only.

## Open questions
1. Which repo and branch: the tables, reader and review screen belong in **FBA_DATABASE**, and the PL export replaces the builder's typing. Confirm before planning.
2. Whether the warehouse will scan or label cartons at destuffing, which would create carton records for files with no carton list. That decides how many containers get carton-level data at first.
