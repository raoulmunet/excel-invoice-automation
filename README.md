# Excel Invoice & Quotation Automation

Microsoft Excel + VBA business automation portfolio project: customer catalog, product/service catalog, audit history and dashboard, with quotation, invoice, PDF and Outlook automation planned in later releases.

> **Status:** v0.1 basic functionality was tested successfully by the project owner in Windows 11 + Excel 2019 (VirtualBox), October 2026. **v0.2 Quotation Engine source is now available but not yet tested in Excel**. Its calculation, storage and compatibility must be validated before production use.

## Implemented in v0.1

- Build or rebuild the dashboard without deleting catalog records.
- Create customer records with generated IDs, contact email, country and active status.
- Create product/service records with generated IDs, unit price and a configurable *demo* tax percentage.
- Browse catalog worksheets and refresh customer/product counts.
- Record creation events in the `Audit Log` worksheet.
- Use late-bound/standard VBA and worksheet-based buttons rather than ActiveX.

### Explicitly not yet implemented

Legal invoice generation, Word templates, PDF exports, Outlook email drafts, payments, multi-user control, inventory tracking and legal tax compliance. They are on the roadmap; v0.1 must **not** be presented as a complete invoicing product.

## v0.2 Quotation Engine (source preview; runtime testing pending)

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

**Office dependencies:** v0.1 requires only Excel Desktop and its included VBA environment. Access, Word, PowerPoint, Outlook, databases, third-party libraries and internet access are not required to run v0.1. Future Word/Outlook integrations will require the respective Windows desktop apps; Outlook COM automation requires **Classic Outlook**, not New Outlook.

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
2. Press **Alt+F11**; import **only** `src/modCatalog.bas`, `src/modSetup.bas`, and `src/modQuotations.bas` via **File > Import File**.
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

Detailed test cases in [tests/ACCEPTANCE.md](tests/ACCEPTANCE.md). **v0.1 basic operations were confirmed by the project owner. v0.2 quotation operations remain untested.**

## Repository structure

```text
src/modCatalog.bas
src/modSetup.bas
src/modQuotations.bas
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
- **v0.2** Quotation draft, line items, sample tax and discounts, and quote numbering — published as source; Excel testing pending.
- **v0.3** Invoice records, PDF and Word document templates, Classic Outlook draft creation.
- **v1.0** Payment-status tracking, dashboards, tests, screenshots, release package and Upwork-ready portfolio.

## License

MIT. See [LICENSE](LICENSE).
