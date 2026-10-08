# Excel Invoice & Quotation Automation

Microsoft Excel + VBA business automation portfolio project: customer catalog, product/service catalog, audit history and dashboard, with quotation, invoice, PDF and Outlook automation planned in later releases.

> **Status:** v0.1 source-code MVP. Files are published on GitHub but **not yet compiled or runtime-tested in Excel**. The reference test environment is Windows 11 running Microsoft Office 2019 Desktop in VirtualBox.

## Implemented in v0.1

- Build or rebuild the dashboard without deleting catalog records.
- Create customer records with generated IDs, contact email, country and active status.
- Create product/service records with generated IDs, unit price and a configurable *demo* tax percentage.
- Browse catalog worksheets and refresh customer/product counts.
- Record creation events in the `Audit Log` worksheet.
- Use late-bound/standard VBA and worksheet-based buttons rather than ActiveX.

### Explicitly not yet implemented

Quotation and invoice generation, Word templates, PDF exports, Outlook email drafts, payments, multi-user control, inventory tracking and legal tax compliance. They are on the roadmap; v0.1 must **not** be presented as a complete invoicing product.

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
2. Press **Alt+F11**; import **only** `src/modCatalog.bas` and `src/modSetup.bas` via **File > Import File**.
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

Detailed test cases in [tests/ACCEPTANCE.md](tests/ACCEPTANCE.md). **No Excel runtime tests have been performed yet.**

## Repository structure

```text
src/modCatalog.bas
src/modSetup.bas
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

- **v0.1** Customer/Product Catalog + Dashboard — published source, needs Excel 2019 verification.
- **v0.2** Quotations, line items, calculations, sample tax rules and document numbering.
- **v0.3** Invoice records, PDF and Word document templates, Classic Outlook draft creation.
- **v1.0** Payment-status tracking, dashboards, tests, screenshots, release package and Upwork-ready portfolio.

## License

MIT. See [LICENSE](LICENSE).
