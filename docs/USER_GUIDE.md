# User Guide — Excel Invoice & Quotation Automation

## Supported environment

- Microsoft Excel 2019 Desktop on Windows 11 (tested by the project owner for the core workflow)
- Microsoft Word 2019 Desktop for DOCX/PDF quotation exports
- Classic Outlook 2019 Desktop for composing email drafts; a configured mail profile is required
- New Outlook, Office Online and LibreOffice are not supported for COM/VBA features.

The workbook is a **portfolio demo with fictional invoices and payments**. It is not legal fiscal billing software and does not integrate with RO e-Factura, bank accounts, or payment gateways.

## First-time use

1. Open the trusted `ExcelInvoiceAutomation.xlsm` in Excel Desktop.
2. Enable only this workbook's macros if you trust its contents.
3. In Excel, run `InitializeInvoiceApp` from Alt+F8, if the Dashboard was not already initialized.
4. Confirm that Customers, Products, Dashboard, Quotations, Quotation Lines, Demo Invoices, Payments, Audit Log and Quote Draft sheets exist.

## Customer and product catalog

- Dashboard > **Add customer**: enter name, optional contact email and country.
- Dashboard > **Add product/service**: enter name, unit price and **demo** tax percent (0–100).
- Created IDs are CUS-0001, PRD-0001 and so on. Catalog entries appear in their respective sheets.

## Quotation workflow

1. Click **Quote Draft**.
2. Enter a valid active Customer ID in cell B4, e.g. CUS-0001.
3. Enter discount percent in B5, e.g. 10 (not 0.10).
4. Enter product ID in A10 and quantity in B10; use up to 20 lines through row 29.
5. Click **Create Quote** on Dashboard.
6. A new QUO-0001-style reference appears in Quotations and Quotation Lines.
7. Click **Quote Word/PDF**; select saved quote ID and a **new** DOCX filename. The PDF is generated alongside the DOCX.
8. Inspect the Word and PDF for the correct customer, details, totals and demo disclaimer.

## Demonstration invoice and payments

1. Click **Demo Invoice** and enter an existing quote ID.
2. Review Demo Invoices: a DEMO-INV-0001-style non-fiscal record is created. The same quotation cannot be invoiced twice.
3. Click **Record Payment**. Enter the demo invoice ID and a positive amount not exceeding the outstanding balance. Confirm.
4. Refresh using **Refresh Payments**. Paid, balance and status are derived from the Payments ledger.
5. Use **Filter Invoices** to show All, Paid, Partially Paid, Unpaid or Overdue entries.

Status calculations depend on the due date. Positive balance with a due date before today is Overdue, even if there has been a partial payment.

## Outlook draft — explicit manual sending only

1. Configure Classic Outlook 2019 with a mailbox **you control**.
2. Replace the demo customer's reserved email address with your owned test mailbox before any actual email use.
3. Click **Outlook Draft**; enter the saved QUO ID.
4. Pick the corresponding existing PDF.
5. Review the To field, subject, message and attachment. The macro opens a draft and **never calls Send**.
6. Sending a real email, if desired, requires a deliberate manual action by the user.

## Troubleshooting

- Cannot run macro: check that all six standard modules are imported and that VBAProject compiles; confirm macros are trusted.
- Word automation failure: verify Word Desktop is installed and available through COM.
- Outlook error: configure Classic Outlook, not New Outlook; restart both Outlook and Excel if COM cannot obtain a profile.
- Export will not overwrite files: choose a new DOCX/PDF filename.
- PowerShell is blocked: use a temporary per-process policy only if trusted and permitted, not machine-wide bypass.
- Unexpected totals: confirm locale decimal separator, product unit price and sample tax rate.
- Layout/UI issues: back up workbook, replace modUI/modSetup with current source, compile and initialize.

## Data safety

Do not use real personal customer data in screenshots or GitHub. Keep populated XLSM files out of version control unless intentionally sanitized. Take a backup before updating VBA modules. Avoid editing the Payments ledger manually.

Demo CSV files in `samples/` are reference data only: the current app does **not** provide a bulk CSV importer.

## Validation

The user has confirmed normal catalog, quotation, Word/PDF, demo invoice, payment and Outlook draft workflows in Office 2019. The updated v1.0 visual layer and comprehensive edge cases still require sign-off.
