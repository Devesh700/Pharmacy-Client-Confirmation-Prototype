# Sampada Pharmacy Data Workspace & GST Stock Ledger Prototype

Desktop-first frontend prototype for an independent multi-store pharmacy business to manage store-wise inventory via immutable stock ledgers, generate GST-compliant retail/B2B tax invoices, reconcile imported ERP stock snapshots, and confirm operational workflows with the client.

## Run & Operate
- `pnpm --filter @workspace/pharmacy-confirmation run dev` — runs the Vite dev server on port 3000
- `pnpm --filter @workspace/pharmacy-confirmation run build` — compiles TypeScript and builds production client assets

## Modules & Architecture (Replit Prompt 2 Implemented)
1. **Per-Store Stock Ledger (`/store-stock`):**
   - Immutable movement records: `Current stock = latest imported snapshot (opening) + sum(movements)`.
   - Batch inventory, expiring soon (≤90 days) and expired badges, FEFO auto-selection.
   - Products × Stores matrix view.
   - Click-to-open Stock Ledger drawer with running balances.
   - Stock Adjustments (Damage, Expiry write-off, Count correction, Other).
   - Inter-Store Transfers (`Draft` → `Dispatched` → `Received` with paired ledger postings).
   - Reconciliation view comparing system stock against imported ERP snapshot with 1-click variance adjustments.

2. **Counter POS Billing (`/billing`):**
   - Counter terminal with store selection, customer search/walk-in, doctor name, and prescription link.
   - Type-ahead product search with automatic FEFO batch selection.
   - Live stock indicator ("Available after this sale").
   - Intra-state (CGST + SGST) vs Inter-state (IGST) place of supply calculation.
   - MRP-inclusive vs Exclusive pricing calculation.
   - Grand total in figures and words (`Rupees ... Only`).
   - Payment modes (Cash / UPI / Card / Credit).
   - Instant stock deduction with visual confirmation toast.
   - Hard blocks on selling expired batches or exceeding on-hand stock (unless Admin override).
   - Scheduled drugs (Schedule H/H1) validation requiring doctor name and prescription link.

3. **Invoices & Returns (`/invoices`):**
   - 25 finalized seed invoices across stores with B2B/B2C indicators.
   - Print preview with toggle between A4 Tax Invoice and 80 mm thermal receipt.
   - Invoice cancellation (mandatory reason, reverses stock movements, permanently keeps number as Cancelled).
   - Credit Note / Sales Return creation (restores stock to original batch, printable credit note).

4. **GST Reports (`/gst-reports`):**
   - Sales Register, GSTR-1 style summary (B2B, B2C, HSN-wise summary), Rate-wise tax breakdown, Purchase register, and Tally-style Stock Summary with Excel export.

5. **Roles & Store Scope:**
   - Persona switcher: **Admin**, **Store Owner**, **Manager**, **Pharmacist**, **Staff**.
   - Topbar store selector for multi-store vs single-branch operations.

6. **Client Confirmation (`/client-confirmation`):**
   - Review space with screen-by-screen feedback and **12 CA & Architecture confirmation questions** exportable to JSON.
