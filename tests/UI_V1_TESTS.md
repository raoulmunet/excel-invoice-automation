# v1.0 Dashboard UI Preview — Office 2019 Acceptance

**Status:** source published, Excel 2019 user test pending. Previously confirmed v0.4 and Outlook workflows are preserved; this document is NOT a release sign-off.

## Safe upgrade to your populated XLSM

1. Close Word/Outlook exports, save and back up your WORKING ExcelInvoiceAutomation.xlsm.
2. Download the latest src/modSetup.bas and new src/modUI.bas.
3. In Excel 2019 press Alt+F11. Right-click the old modSetup, Export File (backup), then Remove modSetup (do not export twice).
4. File > Import File: import the latest modSetup.bas. Import modUI.bas.
5. Keep modCatalog, modQuotations, modDocuments and modPayments unchanged; six modules total.
6. Debug > Compile VBAProject. Resolve any error before executing.
7. Alt+F8 > InitializeInvoiceApp. Do not run the fresh-build PowerShell script over populated files.
8. Check that the large dark-blue Dashboard header appears and action buttons are grouped to the right.

## Non-destructive / layout smoke tests

- Existing Customers, Products, Quotations, Quotation Lines, Demo Invoices, Payments, Audit Log rows are unchanged.
- Catalog counters in B6/B7, last refresh in B8, last quotation in B9 are still correct.
- Payment KPI values remain in B20:B24 and agree with demo invoice/payment data.
- Test Add customer; Add product/service; Customers; Products; Refresh KPI buttons.
- Test Quote Draft; Create Quote; Quote Word/PDF; Demo Invoice; Outlook Draft (must never send automatically).
- Test Record Payment; Refresh Payments; Filter Invoices.
- Re-run InitializeInvoiceApp: no double buttons, no lost records, existing quote draft entries preserved.
- Save, close and reopen: theme persists in XLSM and macros work.
- Zoom to 80% / 100% and inspect button labels, overlaps and merged status banner.
- Inspect the Quote Draft and worksheet headers.

## Known v1.0-preview boundaries

This is a **presentation/UX milestone**, not a final v1.0 release package. There is no public downloadable verified XLSM binary yet. Future steps: demo dataset, release packaging, screenshots, comprehensive edge-case tests and optional manual Outlook send verification. This workbook stores fictitious data; it does not issue fiscal invoices.
