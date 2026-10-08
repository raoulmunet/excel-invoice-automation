# v1.0.0 — Excel Invoice & Quotation Automation (Demo Edition)

**Release candidate — requires final Excel 2019 opening/compilation of the metadata-adjusted XLSM before public release.**

## Features
- Customer and product catalogs
- Quotation builder with line items, discounts and sample tax calculation
- Word and PDF quotation export
- Non-fiscal demo invoices with duplicate prevention
- Partial payment ledger, balances, status filters and KPI dashboard
- Classic Outlook 2019 drafts with PDF attachments — **never sends automatically**
- Visual Office 2019 Dashboard

## Environment
Tested by the project owner on Windows 11 / Excel 2019 / Word 2019 / Classic Outlook 2019. Excel Online, LibreOffice and New Outlook do not support these desktop COM workflows.

## Privacy
The downloadable demo workbook is populated only with fictional catalog records, with transactional history cleared. The author and last-modifier metadata were changed to `Demo Project`; the compiled VBA project was kept byte-for-byte unchanged. Static integrity was verified, but the final metadata-adjusted workbook must still be opened and compiled in Excel 2019. No original customer/test-account workbook should be published.

## Important limits
Demonstration only: not fiscal/legal billing software, no RO e-Factura, no payment gateway, no configured currency, no production security assurances. Do not email reserved example-domain addresses.

## Manual publishing checklist
1. Download the v1.0 release candidate ZIP from this conversation.
2. Extract and open `ExcelInvoiceAutomation_v1.0_DEMO.xlsm` in Excel 2019.
3. Ensure Excel does not request workbook repair; run **Alt+F11 > Debug > Compile VBAProject** and verify the demo workflow.
4. Check File > Info > Properties and do not reintroduce author metadata by saving without rechecking.
5. Open https://github.com/raoulmunet/excel-invoice-automation/releases/new and create tag `v1.0.0`.
6. Attach only the sanitized XLSM or ZIP, **not** the original private workbook.
7. Publish as stable only after the owner verifies the final file; otherwise mark **pre-release**.

Core application source and detailed test plans live in this repository.
