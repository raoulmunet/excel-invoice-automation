# v0.4 Payment Tracking Test Plan

Status: SOURCE PREVIEW. The new module has not been tested in Excel. Target: Windows 11, Excel 2019 Desktop in VirtualBox.

## Upgrade existing v0.3 XLSM
1. Back up your working XLSM.
2. Replace the existing modSetup with the latest src/modSetup.bas.
3. Import src/modPayments.bas. Keep modCatalog, modQuotations and modDocuments.
4. Debug > Compile VBAProject; stop if any error.
5. Run InitializeInvoiceApp from Alt+F8. Existing customer, product, quotation and demo invoice records should be preserved.
6. Verify Payments sheet, and new columns I Amount Paid, J Balance, K Payment Status in Demo Invoices.
7. Dashboard should show Record Payment, Refresh Payments, Filter Invoices and KPI values at A19:B24.

## Example: demo invoice with gross 270.00, due tomorrow
- Refresh: paid 0, balance 270, status Unpaid.
- Record 100: expect PAY-0001, paid 100, balance 170, status Partially Paid.
- Record 170: expect PAY-0002, paid 270, balance 0, status Paid.
- Record another positive payment on the same invoice: rejected.
- Enter an amount greater than remaining balance on any test invoice: rejected.
- Filter Paid: fully paid invoices displayed; filter All: all restored.
- Audit Log: RECORD_DEMO_PAYMENT entries for successful records.
- Save, close, reopen: payment records remain intact.

## Other tests
- Unknown demo invoice: rejected.
- Invalid, blank, negative or zero amount: rejected.
- Cancelling the confirmation: no record.
- Change a fictional invoice due date to yesterday and refresh: Overdue if balance remains positive.
- Rerun InitializeInvoiceApp: no existing records deleted.
- Compare total billed, paid and outstanding KPIs against invoice rows.

Status rules: Paid (zero balance); Overdue (balance positive and due date before today); Partially Paid (some payment and not overdue); otherwise Unpaid.

Limitations: synthetic data only, single-user workbook, no refunds/reversals, no currency management, no fiscal or real payment processing. Editing ledger cells directly may invalidate calculations. All v0.4 tests are pending user confirmation.
