# Sampada Pharmacy Data Workspace & GST Stock Ledger

Desktop-first, light-theme frontend prototype for independent pharmacy teams to manage multi-store inventory, track batch-wise stock ledgers, generate GST-compliant retail/B2B tax invoices, reconcile ERP imports, and confirm operational requirements with the client and their Chartered Accountant (CA).

> **Disclaimer:** All GST calculation logic and reports in this prototype are **Illustrative: to be verified by the client's Chartered Accountant (CA) before production**.

---

## 1. Core Architecture: Stock Ledger (Opening + Movements)

Stock is no longer treated as a static number. The platform implements an immutable, per-store, per-product, per-batch **Stock Ledger**:

$$\text{Current Stock} = \text{Latest Imported Snapshot (Opening Balance)} + \sum(\text{Movements after snapshot date})$$

### Movement Types & Inventory Effects:
| Movement Type | Stock Effect | Reference Document | Source |
|---|---|---|---|
| **Opening / Stock import** | Baseline snapshot | Snapshot as of date | ERP Import |
| **Purchase (in)** | $+ \text{qty}$ | Supplier Invoice No | ERP Import / App Entry |
| **Sale (invoice)** | $- \text{qty}$ (immediate) | Tax Invoice No (`MS/2026-27/00001`) | App Billing |
| **Sales return / credit note** | $+ \text{qty}$ (restored to batch) | Credit Note No (`CN/2026-27/00001`) | App Billing |
| **Purchase return** | $- \text{qty}$ | Debit Note / Return Voucher | App Billing |
| **Transfer out** | $- \text{qty}$ at source store | Transfer Challan ID (`TR-2025-0014`) | App Billing |
| **Transfer in** | $+ \text{qty}$ at receiving store | Transfer Challan ID (`TR-2025-0014`) | App Billing |
| **Adjustment** | $\pm \text{qty}$ (Damage / Expiry / Count) | Adjustment ID (`ADJ-2025-0008`) | App Billing |

### Source-of-Truth & Duplicate Avoidance Policy:
- Sales generated in the application deduct store batch stock **immediately**.
- **Go-Live Date for App Billing:** Configurable in Admin (default `2025-06-01`). ERP Sales imports dated on or after this date display a prominent warning banner: *"Possible duplicate of app bills"* and are not deducted twice.
- The 3-month average sales and stock cover calculations dynamically combine historical ERP net sales with live App Billing sales and returns.

---

## 2. Screens & Functional Modules

### 2.1 Dashboard (`/`)
- **Today's Sales Card:** Live bill count and value for the active store (or consolidated across all stores).
- **Today's GST Collected Card:** Real-time breakdown of CGST, SGST, and IGST collected.
- **Expiring in 90 Days:** Count of batches requiring FEFO (First Expiry First Out) prioritization.
- **Stock Mismatches (Recon):** Discrepancies between app stock ledger and latest imported ERP stock snapshot.
- Topbar **Store Switcher** (All Stores, Main Store, Branch 2, Branch 3) and **Role Switcher** (Admin, Store Owner, Manager, Pharmacist, Staff).
- Recent App Invoices and Recent Stock Ledger Movements widgets.

### 2.2 Store Stock (`/store-stock`)
- **Store Selector:** Filter by Main Store, Branch 2, Branch 3, or All Stores consolidated.
- **Batch Inventory Table:** Product, Store, Batch, Expiry, Available Qty, Reserved (0 in prototype), MRP, Last Movement, and status badges (`Expiring soon ≤90d`, `Expired`, `Low stock`).
- **Matrix View:** Products $\times$ Stores grid showing physical on-hand quantity at each location.
- **Stock Ledger Drawer:** Click any inventory row to slide out the detailed ledger showing dated movement records, reference links, user, movement type, and running stock balance.
- **Stock Adjustment Modal:** Mandatory reason selection (`Damage`, `Expiry write-off`, `Count correction`, `Other`) with optional notes.
- **Inter-Store Transfers:** Create multi-store transfers with a 3-step state machine (`Draft` $\rightarrow$ `Dispatched` $\rightarrow$ `Received`). Deducts from source store on dispatch and adds to destination store upon receipt confirmation.
- **Reconciliation Tab:** Compare current system ledger stock with imported ERP stock snapshot. One-click **Accept Variance** or **Create Adjustment** with reason. Excel export.

### 2.3 Billing / Counter POS (`/billing`)
- **Counter Terminal Layout:** Store picker (defaults to active store), date/time, customer search (or Walk-in / Cash), optional **Doctor Name**, and **Link Prescription** selector.
- **FEFO Batch Auto-Selection:** On selecting a product, the earliest expiring non-expired batch with positive stock is automatically picked. Dropdown allows manual batch switching.
- **Live Stock Indicator:** Displays "Available after this sale: X units" beside each line item.
- **Tax Calculation Engine:**
  - Place of supply evaluation: Intra-state (Karnataka) applies **CGST + SGST (50% each)**; Inter-state applies **IGST**.
  - MRP-Inclusive vs. Exclusive pricing calculation (toggleable in Admin Settings).
  - Discount % or flat amount per line.
  - Per-line tax rounding to 2 decimals, and separate invoice round-off line.
  - **Grand Total in figures and words** (Indian numbering format: *"Rupees One Thousand Two Hundred Only"*).
- **Payment Modes:** Cash, UPI, Card, or Credit (mock) with amount received and balance/change calculator.
- **Save & Print Tax Invoice:** Assigns next sequential store invoice number, posts sale movements to the stock ledger, displays stock update toast (`Stock updated: Paracetamol 500mg 120 → 110`), and opens print preview.
- **Hold Bill & Save Draft:** Hold active counter bills to serve other walk-in customers and resume anytime.

### 2.4 Invoices (`/invoices`)
- Comprehensive invoice register with filters: Date range, Store, Customer, Payment Mode, Status (`Final`, `Draft`, `Cancelled`), GST Type (`B2B`, `B2C`), and search.
- **A4 Tax Invoice Layout & 80 mm Thermal Receipt:** Built-in print preview with toggle between full A4 tax invoice and compact 80 mm thermal counter receipt (print CSS optimized).
- **Cancel Invoice:** Finalized invoices cannot be deleted or directly edited. Cancellation requires a mandatory reason, reverses stock back to the original batches via Adjustment movements, and permanently preserves the invoice number marked as `Cancelled`.
- **Create Credit Note / Sales Return:** Select returned lines and quantities; automatically returns stock to the original batch and generates a printable Credit Note referencing the original invoice number and date.

### 2.5 Customers (Master)
- Registered B2B customers (hospitals, clinics, nursing homes) and B2C walk-ins.
- Captures Name, Phone, Address, State, and **15-character GSTIN validation** (illustrative regex check).
- Inline customer registration directly from the Billing POS counter.

### 2.6 Purchases (Stock-In Entry)
- Lightweight stock-in entry modal: Supplier, Supplier Invoice No & Date, Product, Batch, Expiry, Billed Qty, Free Bonus Qty, Purchase Rate, and MRP.
- Submitting posts `Purchase (in)` movements directly to the selected store's stock ledger.

### 2.7 GST Reports (`/gst-reports`)
- Period selector (month/year) and store filter. Exportable to Excel:
  1. **Sales Register:** Complete invoice register with Taxable Value, CGST, SGST, IGST, and Grand Total.
  2. **GSTR-1 Style Summary:** B2B invoice table (with GSTIN), B2C summary by tax rate and place of supply, Credit/Debit notes, and **HSN-wise summary** (HSN, description, UQC, total qty, taxable value, tax breakdown).
  3. **Rate-wise Tax Summary:** Tax collected broken down by rate tier (0%, 5%, 12%, 18%).
  4. **Purchase Register:** Input tax breakdown by rate for accountant reference.
  5. **Tally-style Stock Summary:** Opening balance, inwards (purchases & transfers in), outwards (sales & transfers out), and closing balance per product with value at purchase rate.
- Prominent compliance label: *"Prototype report format. Not a filing-ready return. To be verified by CA."*

### 2.8 Prescriptions (`/prescriptions`)
- Patient prescription tracking with pharmacist review stages (`Received` $\rightarrow$ `Review` $\rightarrow$ `Process/Print` $\rightarrow$ `Closed`).
- **"Create Bill" Handoff:** Click "Create Bill" on any prescription to open Billing with the patient name, phone, doctor name, and linked medicines prefilled. Upon invoice generation, the invoice number is linked back to the prescription record.
- Safety notice: *"The system does not interpret prescriptions or recommend or substitute medicines. The pharmacist is responsible for evaluation and dispensing."*

### 2.9 Admin (`/admin`)
- **Stores Master Data:** Legal entity names, addresses, states, GSTINs, Drug Licence numbers, and annual invoice number prefixes (`MS/2026-27/`, `B2/2026-27/`, `B3/2026-27/`).
- **Billing & GST Settings:**
  - *Prices are MRP-inclusive of GST* toggle (with mathematical formula tooltip).
  - *Allow Negative Stock* toggle (default OFF; Admin override only).
  - *Go-live date for app billing* setting.
  - *E-invoicing / E-way bill note*: Guidance on turnover thresholds (confirm with CA).
- **Users & Role Matrix:** Detailed permission grid for all 5 roles.
- **Audit Trail:** Persistent session log of bill creation, cancellations, credit notes, stock adjustments, and transfers.

### 2.10 Client Confirmation (`/client-confirmation`)
- Interactive review screen for all 14 screens.
- **12 Replit Prompt 2 Architecture Questions** covering:
  1. Billing source of truth vs. MediVision
  2. Bill of Supply / Composition scheme requirements
  3. Store GST registration, state, and drug licence numbers
  4. Invoice numbering formats and series
  5. MRP-inclusive discount policies and GST rate maintenance
  6. B2B credit customer ledgers and receivables scope
  7. E-invoicing / E-way bill portal integration
  8. Counter barcode scanners and thermal printers
  9. Batch-wise tracking and FEFO policy
  10. Scheduled drug registers and pharmacist signatures
  11. Inter-store transfer GST delivery challans
  12. Single vs. multiple GSTIN operations
- Export answers and feedback as structured `.json` files.

---

## 3. Role Permissions Matrix

| Capability | Admin | Store Owner | Manager | Pharmacist | Staff |
|---|:---:|:---:|:---:|:---:|:---:|
| **All Stores Consolidated View** | Yes | Own Store | Yes | Own Store | Own Store |
| **Counter Billing POS** | Yes | Yes | Restricted | Yes (Rx / Scheduled) | Draft bills only |
| **Cancel Invoice & Issue Credit Note** | Yes | Yes | No | No | No |
| **Store Stock & Ledger Drawer** | Yes | Yes | Yes | Yes (View) | Yes (View) |
| **Stock Adjustments & Transfers** | Yes | Yes | Yes | No | No |
| **Reconciliations with ERP** | Yes | Yes | Yes | No | No |
| **GST Reports & Tax Registers** | Yes | Yes | Yes | No | No |
| **Excel Workbook Imports** | Yes | No | Yes | No | No |
| **Override Negative Stock** | Yes | No | No | No | No |
| **Admin Store Master & GST Config** | Yes | View only | No | No | No |

---

## 4. How to Test the Guardrails & Validations

1. **Negative Stock Blocking:**
   - Go to **Billing** (`/billing`).
   - Pick a product with low stock (e.g. *Vitamin D3 60K* with 18 units).
   - Enter a quantity of `25` (exceeding on-hand stock).
   - Click *Add Line* $\rightarrow$ The system displays an error: *"Cannot sell more than available batch stock (18 units)"*.
   - In Admin $\rightarrow$ Settings, enable *Allow Negative Stock* (Admin only) to test the override behavior.

2. **Expired Batch Sale Blocking:**
   - Go to **Billing** (`/billing`).
   - Pick a product having an expired batch (e.g., *Cetirizine 10 mg* batch `BT-2200` expiring `2025-04`).
   - Notice the batch selector marks it with an `[EXPIRED]` tag and disables addition to the bill with a clear error: *"Cannot dispense expired batch"*.

3. **Scheduled Drugs Validation (Schedule H / H1):**
   - In Billing, select *Amoxicillin 500 mg* (Schedule H1) or *Atorvastatin 10 mg* (Schedule H).
   - Notice the prominent amber badge: `Schedule H / H1 - Doctor details required`.
   - Clear the Doctor Name field and attempt to Save $\rightarrow$ Blocked with notice: *"Doctor Name and Prescription Link required for scheduled drug sales"*. Enter doctor name to save.

4. **FEFO (First Expiry First Out) Auto-Pick:**
   - In Billing, select *Paracetamol 500 mg*.
   - Notice the batch dropdown automatically selects the earliest expiring batch (`BT-2200` expiring `2025-07`) over later batches (`BT-2400` expiring `2026-11`).

5. **Inter-State Tax Logic (CGST/SGST vs. IGST):**
   - In Billing, select customer *Meera Krishnan* (Karnataka) $\rightarrow$ Taxes display as **CGST (6%) + SGST (6%)**. Place of supply: *Karnataka*.
   - Switch customer to *Apollo First Aid Station (Hosur)* (Tamil Nadu) $\rightarrow$ Taxes display as **IGST (12%)**. Place of supply: *Tamil Nadu*.

6. **Invoice Cancellation & Stock Reversal:**
   - In **Invoices** (`/invoices`), select invoice `MS/2026-27/00001`.
   - Click *Cancel Invoice*, enter a reason (e.g., "Customer returned prior to leaving counter").
   - Notice invoice status changes to `Cancelled`.
   - Open **Store Stock** $\rightarrow$ Click *Stock Ledger* $\rightarrow$ Verify that positive `Adjustment` movements (`CANCEL:MS/2026-27/00001`) were automatically posted, restoring stock to the batch.

7. **Credit Note (Sales Return):**
   - In **Invoices**, click *Credit Note* on an invoice.
   - Select 1 unit to return, enter reason.
   - A Credit Note is generated (`CN/2026-27/00001`) referencing the invoice, and 1 unit is restored to the batch in the Stock Ledger.

8. **Inter-Store Transfer State Machine:**
   - In **Store Stock** $\rightarrow$ click *Inter-Store Transfer*.
   - Create transfer of 10 units from Main Store to Branch 3.
   - Status transitions from `Draft` $\rightarrow$ `Dispatched` (deducts 10 from Main Store) $\rightarrow$ `Received` (adds 10 to Branch 3).

---

## 5. Scope Boundaries: What is NOT in Prototype

In accordance with prompt requirements, the following are simulated or out of scope:
- **No real GST portal / ClearTax / Sandbox API integration:** GSTR-1, GSTR-3B summaries and tax calculations are illustrative for client review and CA confirmation.
- **No real E-invoicing / E-way bill generation:** Informational notes provide statutory thresholds.
- **No accounting double-entry ledgers:** No general ledger, debtors/creditors accounting, balance sheet, or P&L.
- **No live payment gateway:** Cash, UPI, Card, Credit buttons are simulated prototype selections.
- **No hardware printer drivers / scanner drivers:** Print preview uses standard CSS print stylesheets (supporting A4 and 80 mm receipt viewports).
- **No medical advice, recommendation or medicine substitution:** The platform specifically forbids automated drug recommendations.

---

## 6. Verification Test Log

| Test Case | Expected Behavior | Observed Result | Status |
|---|---|---|:---:|
| **Dashboard KPIs** | Displays Today's Sales, GST, Expiring, and Mismatches | Accurate cards rendered with real-time computations | **PASS** |
| **Store Switcher** | Filters data across Main Store, Branch 2, Branch 3, All | All views respect selected store | **PASS** |
| **FEFO Auto-Pick** | Picks earliest valid batch on product selection | Earliest batch selected by default | **PASS** |
| **Negative Stock Block** | Blocks sale if qty > available batch stock | Blocked with warning banner | **PASS** |
| **Expired Batch Block** | Blocks sale of batch with past expiry date | Blocked with badge and error | **PASS** |
| **Grand Total in Words** | Converts rupee amount to words format | Displayed correctly (e.g., *Rupees One Thousand...*) | **PASS** |
| **Stock Deduction on Bill** | Deducts batch stock immediately upon save | Verified in Stock Ledger with toast | **PASS** |
| **Invoice Cancellation** | Reverses stock movements and keeps invoice number | Stock restored, marked `Cancelled` | **PASS** |
| **Credit Note** | Returns stock to original batch with credit document | Stock restored, credit note issued | **PASS** |
| **A4 & 80mm Print** | Toggle between A4 Tax Invoice and 80mm receipt | Clean responsive print styles | **PASS** |
| **Inter-Store Transfer** | Paired out/in movements with Draft/Dispatched/Received | Dispatched deducts; Received increments | **PASS** |
| **Reconciliation Variance** | Shows difference between ledger and ERP snapshot | Difference displayed; adjust/accept works | **PASS** |
| **CA Confirmation Questions**| 12 questions rendered with editable inputs and export | Answers editable, exported to JSON | **PASS** |