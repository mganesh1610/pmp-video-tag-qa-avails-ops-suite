# Agency IO Intake, Operations Launch Checklist & Post-Sales Billing Reconciler (`P1 Module`)

Extends `pmp-video-tag-qa-avails-ops-suite` with media agency insertion-order (IO) intake, Operations launch tracking, and exact `Decimal` order-to-billing reconciliation across **40,144** SHA-256-verified public media records (`39,644` UCI Online News Popularity placements + `500` UCI Paid/Organic Creative Deliverables).

## Key Verified Outputs
- Reconciles **24 media agency insertion orders** across **8 agencies** and **6 sales verticals** with **`$0.00` billing ledger variance** (`Booked IO = Billable Invoiced + Make-Good Credits`).
- Exports a formula-driven Microsoft Excel workbook (`outputs/agency_io_billing_reconciliation.xlsx`) and SQLite audit ledger (`outputs/agency_io_billing.db`).

## Quick Run
```bash
python3 -m agency_billing_io_reconciliation.io_billing_reconciler
python3 -m unittest discover -s tests -v
```
