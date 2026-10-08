# Excel Invoice & Quotation Automation

Microsoft Excel + VBA business automation portfolio project: customer catalog, product/service catalog, audit history and dashboard, with quotation, invoice, PDF and Outlook automation planned in later releases.

> **Status:** v0.1 basic functionality was tested successfully by the project owner in Windows 11 + Excel 2019 (VirtualBox), October 2026. **v0.2 Quotation Engine basic operation is confirmed by the project owner in Excel 2019** after importing the missing `modQuotations.bas` module. **v0.3 Word/PDF quotation export has been confirmed working by the project owner in Windows 11 / Excel and Word 2019.** **Demo invoice creation and duplicate prevention have also been confirmed by the project owner.** **Classic Outlook 2019 is now configured and the Excel VBA Outlook Draft workflow was user-tested successfully**, including test recipient, quote subject, body and attached PDF, without automatic sending. Manual dispatch from the created draft has not been separately verified. **v0.4 payment tracking was tested successfully by the project owner in Windows 11 / Excel 2019; detailed edge cases were not individually reported.**

## Implemented in v0.1

- Build or rebuild the dashboard without deleting catalog records.
- Create customer records with generated IDs, contact email, country and active status.
- Create product/service records with generated IDs, unit price and a configurable *demo* tax percentage.
- Browse catalog worksheets and refresh customer/product counts.
- Record creation events in the `Audit Log` worksheet.
- Use late-bound/standard VBA and worksheet-based buttons rather than ActiveX.

### Explicitly not yet implemented

Legally compliant fiscal invoices, external editable Word templates, payment processing, multi-user control, inventory tracking and tax compliance. They are on the roadmap; v0.1 must **not** be presented as a complete invoicing product.

## v0.2 Quotation Engine (basic Excel 2019 operation confirmed)

- A **Quote Draft** worksheet for Customer ID, discount percent and up to 20 catalog products with quantities.
- `CreateQuotation` validates active customer and products, quantities, prices and sample tax rates before saving.
- Generated `QUO-0001`-style identifiers; immutable snapshots of names, unit prices and tax rates into `Quotations` and `Quotation Lines` tables.
- Line-level net, discount, tax and gross; summary-level totals. Money uses two-decimal rounding. This is a **fictional quotation tool**, not legally compliant invoice software.
- New Dashboard buttons: **Quote Draft** and **Create Quote**.

### Upgrade your already working v0.1 XLSM (recommended)

1. Back up the existing working `.xlsm` file. **Do not rebuild from scratch if you wish to retain your customers/products.**
2. Download the latest `src/modSetup.bas` and `src/modQuotations.bas`.
3. In Alt+F11, remove the old `modSetup` module (export a backup first); import the updated `modSetup.bas`.
4. Import the **new** `modQuotations.bas`. Keep `modCatalog.bas` unchanged.
5. Select **Debug > Compile VBAProject** before proceeding. Any error must be resolved first.
6. Run **`InitializeInvoiceApp`** from Alt+F8. It adds quotation sheets without clearing existing catalogs.
7. On Dashboard use **Quote Draft**, fill `B4 = CUS-0001`, `B5 = 10`, `A10 = PRD-0001`, `B10 = 2`. Return to Dashboard and click **Create Quote**.
8. Verify data in `Quotations`, `Quotation Lines` and `Audit Log`. Save/reopen and verify persistence.

**Expected demo** (assuming PRD-0001 unit price 150, demo tax 0%, quantity 2, discount 10%): net **300.00**, discount **30.00**, tax **0.00**, gross **270.00**. If your catalog contains different values, the totals will differ accordingly.

### Fresh v0.2 build

The updated `build/Create-InvoiceWorkbook.ps1` imports **three modules**: `modCatalog.bas`, `modSetup.bas` and `modQuotations.bas`. It creates a clean `.xlsm` from scratch and refuses to overwrite an existing output. Follow the build instructions below, but use the three modules for manual import.

Read the acceptance checklist: [tests/QUOTATIONS_TESTS.md](tests/QUOTATIONS_TESTS.md).

## v0.3 Word / PDF / Demo Invoice / Outlook (source preview)

**Word and PDF generation owner-tested successfully in Microsoft Office 2019 (2026-10-08); other v0.3 features still await validation.** This version adds `src/modDocuments.bas` and new Dashboard buttons:

- **Quote Word/PDF** — enter saved quotation ID (e.g. `QUO-0001`), choose an output `.docx` name, and produce Word and PDF files with quote lines and totals using Word 2019 Desktop automation.
- **Demo Invoice** — record a **DEMO-INV-0001**-style internal document from an existing quotation in a `Demo Invoices` worksheet. This is a non-fiscal, non-compliant **demonstration record**, not a VAT invoice or legal billing document.
- **Outlook Draft** — select an existing quote and manually select the generated PDF to attach; opens a Classic Outlook draft with the catalog customer email. It **never sends email automatically**. Review recipient, amount and attachments before sending.

**Limitations:** Word document formatting uses programmatically created content; external DOCX template customization is scheduled for a later version. The demo has no configured currency, business issuer details, VAT registration, fiscal invoice numbering compliance, Romanian RO e-Factura integration or payment reconciliation. It must not be used to issue real-world legal invoices. DOCX and PDF exports refuse to overwrite existing files. The Outlook PDF picker does not verify that the manually selected PDF matches the quotation — check it yourself.

### Upgrade from your working v0.2 XLSM

1. **Save a backup of the working XLSM** (or download source from `backup/v0.2-tested`). Do not run the clean-build script on top of your populated workbook.
2. Download the latest `src/modSetup.bas` and new `src/modDocuments.bas` from GitHub.
3. In Excel press **Alt+F11**; export a backup of old `modSetup` and remove that old module. Import the new `modSetup.bas`.
4. Import **`modDocuments.bas`**. Keep `modCatalog.bas` and `modQuotations.bas` unchanged. Four standard modules are expected.
5. Select **Debug > Compile VBAProject** and stop if any error appears.
6. Run **`InitializeInvoiceApp`** via Alt+F8. It creates the `Demo Invoices` sheet and new Dashboard buttons **without deleting your customers, products or quotations**.
7. Ensure Microsoft Word 2019 Desktop is installed before clicking **Quote Word/PDF**. Use fictional quote data and a fresh output path.
8. If Classic Outlook is installed and configured, test **Outlook Draft**, but **do not send** the email.
9. Test **Demo Invoice** separately. Check `Demo Invoices` and `Audit Log`.

See [v0.3 acceptance tests](tests/DOCUMENTS_TESTS.md).

### Fresh installation

The clean workbook builder now imports **four source modules**: `modCatalog`, `modSetup`, `modQuotations` and `modDocuments`. It does not overwrite existing XLSM files.

## v0.4 Payment Tracking & Business Dashboard (owner-tested)

v0.4 adds the standard VBA module `src/modPayments.bas` and works with **fictional demo invoices only**.

- `Payments` worksheet: append-only records (payment ID, timestamp, demo invoice ID, amount, reference).
- `Demo Invoices` columns I–K: amount paid, outstanding balance and calculated status.
- Statuses: **Unpaid**, **Partially Paid**, **Paid**, **Overdue**. Overdue takes precedence over partial payment when balance remains and due date is before today.
- Dashboard KPIs at A19:B24: number of demo invoices, total demo billed, recorded payments, outstanding and overdue balances.
- New commands: **Record Payment**, **Refresh Payments**, **Filter Invoices**. Filters cover statuses or All.

### Upgrade without losing v0.3 records

1. Back up your already working `.xlsm`.
2. Download **the newest** `src/modSetup.bas` and **new** `src/modPayments.bas`.
3. In the VBA editor, export a backup and remove the old `modSetup`; import the updated `modSetup.bas` and new `modPayments.bas`. Retain `modCatalog`, `modQuotations`, `modDocuments`.
4. Run **Debug > Compile VBAProject**. Resolve all errors before testing.
5. Execute `InitializeInvoiceApp` using Alt+F8. Existing catalog, quote and demo invoice records are retained. A new `Payments` sheet and derived invoice columns are added.
6. Test partial and full payments against a fictional invoice, refresh KPI metrics, filter and save/reopen.

See [tests/PAYMENTS_TESTS.md](tests/PAYMENTS_TESTS.md) for exact test values.

**Constraints:** not legal accounting/invoicing software; no payment gateway, currency conversion, refunds or transaction rollback. Payment records are stored locally in one workbook, intended for a single user; manual cell edits and concurrent access are unsupported. The owner confirms that v0.4 payment functionality works in Excel 2019. Individual negative and edge cases have not all been separately verified.

### Fresh build

The PowerShell builder imports **five** source modules: `modCatalog`, `modSetup`, `modQuotations`, `modDocuments`, `modPayments`. Never overwrite an XLSM that already contains your data.

## v1.0 interface preview — modern Excel Dashboard

**Status: published source; awaiting owner tests in Excel 2019.** The earlier v0.4 / Outlook integration smoke tests were successful, but this new presentation layer still needs verification. **This is not the final v1.0 release.**

Changes: a dark-blue dashboard header, grouped quick actions, separate commercial/payment navigation areas, redesigned customer and financial KPI blocks, demo-only warning, and coordinated data-sheet header styling. No existing VBA business logic is changed. Module: [src/modUI.bas](src/modUI.bas).

### Upgrade without losing your data

1. Save a backup of your working ExcelInvoiceAutomation.xlsm.
2. Download the latest [modSetup.bas](src/modSetup.bas) and new [modUI.bas](src/modUI.bas).
3. In the VBA editor export/remove **only the old** `modSetup`, import updated `modSetup.bas`, and import `modUI.bas`.
4. Keep existing `modCatalog`, `modQuotations`, `modDocuments`, `modPayments`. Expected total: **six** standard modules.
5. Run **Debug > Compile VBAProject**. Stop if an error occurs.
6. Use **Alt+F8 > InitializeInvoiceApp**. It reapplies the new Dashboard without clearing existing records. Verify the action buttons and KPI counters.
7. Save/reopen the workbook. Check all existing functionality.

To build a fresh **empty** XLSM instead, the updated PowerShell builder imports six modules, including `modUI`; it refuses to overwrite an existing workbook.

See [tests/UI_V1_TESTS.md](tests/UI_V1_TESTS.md) for visual and regression checks. **Screenshots, fictitious demo dataset, verified binary release package and full end-to-end tests are upcoming tasks.**

## Environment and compatibility

| Environment | Status | Details |
|---|---|---|
| Windows 11 + Excel 2019 Desktop / VirtualBox | **Primary target; not tested yet** | First acceptance environment |
| Windows 11 + Excel 2021, 2024, Microsoft 365 Desktop | Target; untested | Windows desktop VBA enabled |
| Windows + Excel 2016 Desktop | Target; untested | No modern spreadsheet functions required |
| Excel 32-bit / 64-bit | Designed for both; untested | No Windows API declarations |
| Windows 10 | Possibly compatible; not tested | Standard Windows 10 support ended October 2025 |
| macOS Excel | Not supported/tested | Windows-based packaging and later Office COM automation |
| Linux, LibreOffice or Wine | Not supported | No reliable Excel VBA compatibility claim |
| Excel in browser | Not supported | Excel Online does not run VBA |

**Office dependencies:** v0.1/v0.2 require only Excel Desktop and VBA; v0.3 document export additionally requires **Microsoft Word Desktop for Windows**, and email drafts require **Classic Outlook for Windows**. Access, Word, PowerPoint, Outlook, databases, third-party libraries and internet access are not required to run v0.1. All Word and Outlook automation is late-bound COM. **New Outlook does not support this Classic Outlook automation model.**

**Security and support:** Excel 2019 (and Office 2016) reached end of support on 14 October 2025. Prefer supported Office for client production work. Do not enable macros from untrusted sources or globally disable macro security. The optional automated build requires temporary **Trust access to the VBA project object model** in Excel; turn it off when finished. Corporate policies may prohibit this option.

## Build clean workbook on Windows

### Automated build using Excel 2019

1. Download the complete repository and extract it.
2. Review `src/` and `build/Create-InvoiceWorkbook.ps1`.
3. In Excel go to **File > Options > Trust Center > Trust Center Settings > Macro Settings**, temporarily enable **Trust access to the VBA project object model**. Do not enable all macros.
4. Close Excel.
5. In Windows PowerShell, from the repository root, run:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\build\Create-InvoiceWorkbook.ps1
```

The first command is only needed if your environment permits it and the script is blocked; do not override organization-enforced restrictions. The script refuses to overwrite any existing workbook and creates `ExcelInvoiceAutomation.xlsm` in the repository root.
6. Open the workbook and choose **Alt+F11 > Debug > Compile VBAProject**. Verify the workbook has **Dashboard**, **Customers**, **Products** and **Audit Log** sheets.
7. Turn off the temporary VBA object model access setting.

### Manual build without PowerShell

1. In Excel Desktop create a blank workbook; Save As `ExcelInvoiceAutomation.xlsm`.
2. Press **Alt+F11**; import the six modules: `src/modCatalog.bas`, `src/modSetup.bas`, `src/modQuotations.bas`, `src/modDocuments.bas`, `src/modPayments.bas`, and `src/modUI.bas` via **File > Import File**.
3. Run **Debug > Compile VBAProject**.
4. Run `InitializeInvoiceApp` via **Alt+F8**.
5. Remove unused blank worksheets if needed; keep all application sheets.

## Demo / acceptance test

1. Run `InitializeInvoiceApp` again: no catalog records should be deleted.
2. On Dashboard click **Add customer**, enter `Contoso Demo SRL`, `demo@example.invalid`, `Romania`; expected ID `CUS-0001`.
3. Click **Add product/service**, enter `Consulting hour`, unit price `150`, demo tax percentage `0`; expected ID `PRD-0001`.
4. Click **Refresh KPI**; expect one active customer and one active product.
5. Check `Audit Log` has entries `CREATE_CUSTOMER` and `CREATE_PRODUCT`.
6. Save, close and reopen; confirm catalog data persists.

Detailed test cases in [tests/ACCEPTANCE.md](tests/ACCEPTANCE.md). **v0.1 and v0.2 basic operations have been confirmed by the project owner. Individual acceptance assertions remain pending.**

## Repository structure

```text
src/modCatalog.bas
src/modSetup.bas
src/modQuotations.bas
src/modDocuments.bas
src/modPayments.bas
src/modUI.bas
build/Create-InvoiceWorkbook.ps1
samples/demo-customers.csv
samples/demo-products.csv
tests/ACCEPTANCE.md
README.md
LICENSE
```

## Design choices

- The workbook itself is the local data store. It is **single-user**, and simultaneous edits are not supported.
- IDs use an incrementing prefix derived from current worksheet values and are not collision-proof under concurrent/multi-user edits.
- Data validation is basic; monetary and tax numbers follow the local Excel decimal separator.
- The tax rate is **sample configuration**, not automatically updated tax legislation.
- Avoid storing real customer contact data in public demos; use synthetic datasets.
- Subsequent versions may add tables, stronger validation, document numbering, legal disclaimers, and audit hardening.

## Roadmap

- **v0.1** Customer/Product Catalog + Dashboard — basic tests owner-confirmed on Excel 2019.
- **v0.2** Quotation draft, line items, sample tax and discounts, and quote numbering — basic execution confirmed by project owner; detailed test cases pending.
- **v0.3** Word/PDF quote export, demo invoice creation and duplicate prevention — owner-confirmed working. Classic Outlook draft creation owner-confirmed (no automatic sending); manual dispatch not separately tested.
- **v0.4** Demo payment ledger, calculated balances and status, KPI dashboard and invoice filtering — owner-reported successful tests on Excel 2019.
- **v1.0 UI preview** Modern Dashboard, grouped action buttons and visual refinements — source published, Excel 2019 tests pending.
- **v1.0 final** Demo dataset, full regression tests, screenshots, release package and Upwork-ready portfolio — pending.

## License

MIT. See [LICENSE](LICENSE).
