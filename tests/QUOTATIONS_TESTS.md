# v0.2 Quotation Engine — acceptance tests

**Reference:** Windows 11 / Excel 2019 / VirtualBox.
**Status:** basic workflow reported working by the project owner on Windows 11 / Excel 2019 (VirtualBox), after importing the initially missing `modQuotations.bas` module. Detailed numeric and negative tests not individually confirmed.

## Upgrade test
- Back up existing v0.1 workbook; replace `modSetup` and add `modQuotations`.
- Debug > Compile VBAProject: no compile errors.
- Run `InitializeInvoiceApp`; existing customer/product data retained.
- Verify sheets: `Dashboard`, `Customers`, `Products`, `Audit Log`, `Quote Draft`, `Quotations`, `Quotation Lines`.
- Re-run InitializeInvoiceApp: records and existing draft lines remain unchanged.

## Functional test
Create or confirm `CUS-0001` (Active) and `PRD-0001` (Active, unit price 150, tax 0). In Quote Draft enter:
- B4 = `CUS-0001`
- B5 = `10` (discount percent)
- A10 = `PRD-0001`
- B10 = `2`

Click **Create Quote** in Dashboard. Expected quotation:
- Quote ID `QUO-0001` in a fresh workbook
- Net 300.00, discount 30.00, tax 0.00, gross 270.00
- Status `Draft`
- Exactly one quotation line and `CREATE_QUOTATION` audit event
- Repeating with unchanged draft creates `QUO-0002` (this is expected, not deduplicated)

## Negative/edge tests
- Unknown/inactive customer: reject without creating quote
- Unknown/inactive product: reject without creating quote
- Blank/zero/negative quantity: reject
- Missing product ID with a quantity: reject
- Blank/all-empty line area: reject
- Discount outside 0–100: reject
- Different sample tax rates per product: verify tax calculated after discount
- After saving/reopening, quotes and line items persist
- Verify quote numbering when previous quote IDs have gaps
- Rebuild from clean sources with PowerShell

**Limitations:** quote records are drafts, not fiscal invoices. No simultaneous multi-user support or transactional rollback of a partial Excel write. No Word/PDF/Outlook integration yet. No currency/exchange-rate handling. All amounts are demo calculations, not tax advice.

## Owner-reported smoke test

- The Create Quote button initially displayed `Cannot run the macro ... CreateQuotation` because `modQuotations` had not been imported into the existing XLSM.
- After importing that missing module, the owner confirmed that everything works.
- Treat this as a successful basic workflow test, not independent verification of all amounts, edge cases, or build variants.
