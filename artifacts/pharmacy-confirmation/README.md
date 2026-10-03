# Sampada Pharmacy Data Workspace

Desktop-first, light-theme frontend prototype for independent pharmacy teams to inspect product data, reconcile incoming work and confirm operational requirements. All seeded information and edits live in React state in the browser; there is no server persistence.

## Screens

- **Dashboard** — reporting date context, low-stock count, pending matches, WhatsApp review and open prescriptions.
- **Imports** — multi-file `.xlsx` upload, client-side SheetJS parsing, report type recognition, required-value and numeric validation, duplicate filename handling, simulated Parse → Validate → Normalize → Match → Load progress, batch detail, row errors, retry and error export.
- **Products** — searchable, filterable product master; store, supplier, stock, group and cover filters; URL query persistence; exports/print; product detail with aliases, stock, net sales, average, cover, 12-month chart, stores and purchase history.
- **Suppliers** — supplier search, supplied-product and stock summary, purchase history and per-supplier MRP details.
- **Matching Queue** — explicit select existing / create product / skip decisions; no automatic mapping.
- **Ask** — deterministic local keyword parser, editable filter results and interpreted filter JSON. No model or medical advice.
- **WhatsApp Inbox** — simulated central inbox, local message line parsing, confirmation controls and image placeholder.
- **Prescriptions** — local patient record workflow, role-based transitions, status history, A5/A4/receipt preview and browser print.
- **Admin** — prototype roles/matrix, audit activity, usage simulation and calculation settings.
- **Client confirmation** — per-screen decision/comments, nine open client questions and JSON export.
- A persistent **Client feedback** drawer is available on every screen.

## Sample workbooks

Open **Imports → Download 4 sample workbooks**. The browser downloads four separate `.xlsx` files: Sales, Purchase, Sales Return and Stock. Upload one or multiple workbooks from the same screen. SheetJS reads the actual workbook bytes locally and displays detected type, rows and validation errors. Duplicate uploads are blocked unless the active role is Admin and explicitly overrides.

## Roles

Use the role selector in the top bar to switch among **Admin**, **Manager**, **Pharmacist** and **Staff**. The selector is a prototype persona switch, not authentication. Admin can access all actions; Manager does not see prescriptions; Pharmacist manages and prints prescriptions with product read-only access; Staff uses product lookup and WhatsApp and can create a prescription record for review.

## Simulations and safety

The banner marks simulated data/features. There are no real WhatsApp, LLM, ERP, printer, OCR, authentication, API or database connections. Ask is a small deterministic rule parser; imports are parsed and validated locally but do not update an external inventory system. Prescription text is not interpreted and the system does not recommend or substitute medicines. Browser print is a preview only; Epson M104 output needs client-PC validation. All changes are in-memory for the active page session; refreshing restores seed data.