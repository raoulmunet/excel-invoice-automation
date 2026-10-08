# v1.0 Dashboard UI Preview — Office 2019 Acceptance

**Status (2026-10-09):** owner-confirmed UI smoke test in Excel 2019: obsolete row-19 text no longer overlaps and button labels are more legible. Full end-to-end regression remains pending. Previously confirmed v0.4 and Outlook workflows are preserved; this document is NOT a release sign-off.

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

## Screenshot review, 2026-10-08

Initial Dashboard screenshot confirms the overall layout: grouped actions, visible catalog and payment KPIs, and distinct commercial/workflow areas. Two cosmetic issues were identified: the legacy payment label was still visible on row 19, and button label text was small. Source fixes were applied to both `modUI.bas` and `modPayments.bas` (remove obsolete A19 title even on future KPI refresh, enlarge action fonts). **The owner subsequently confirmed both visual corrections successful.**

To test the corrections: replace both `modUI` and `modPayments` in the VBA editor with the latest source files, Compile VBAProject, run `InitializeInvoiceApp` and then `Refresh Payments`, and confirm the ghost heading remains gone.

## Owner visual confirmation — 2026-10-09

**PASS (user-reported):** after replacing modUI and modPayments, the previously overlapping text on row 19 disappeared and button labels became easier to read. This validates the two visual fixes only; it does not prove the complete version 1.0 regression suite or a distributable XLSM has been tested.

## Sanitized package validation (owner report, 2026-10-09)

The owner confirmed completion of all steps through step 5 of the local demo-sanitization workflow, including opening and checking the generated workbook. Result: owner-reported PASS for the sanitized demo smoke test. Next release gate: inspect the generated XLSM itself for privacy/embedded metadata and check a reproducible demo scenario before attaching it to the public GitHub v1.0 release.


## Independent static audit of uploaded sanitized XLSM — 2026-10-09

Uploaded file: ExcelInvoiceAutomation_DEMO.xlsm; ZIP/OOXML structure passes integrity check.
- Nine expected worksheets present; all are visible.
- Customers contains three fictional companies and reserved example-domain email addresses. Products contains four fictional services.
- Quotations, Quotation Lines, Demo Invoices, Payments and Audit Log contain header rows only. Quote Draft contains labels but no persisted transaction.
- No external-link, connection, query, embedding or comment package parts identified.
- The VBA project binary is still embedded in the XLSM; this is structural confirmation, **not VBA execution verification**.
- XML/shared-string email search found only reserved example-domain addresses. The original test mailbox was not found in the scanned VBA binary via simple byte/string probes.
- **Remaining publication decision:** workbook core metadata `dc:creator` and `cp:lastModifiedBy` contain the owner's full name. Confirm deliberate attribution or remove metadata before public redistribution.
- **Limit:** the audit is static and does not certify full VBA binary decompilation, Office functionality, absence of all private identifiers, or absence of malicious code. Final macro execution must be verified by the owner.

**Release gate:** awaiting user choice on retaining/removing personal author metadata; then perform final package hash and publish only the vetted demo workbook.
