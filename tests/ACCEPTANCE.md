# v0.1 Acceptance Tests — Windows 11 / Excel 2019

**Status:** not run. Record the execution date, Excel bitness and result after validation.

| # | Scenario | Expected |
|---|---|---|
| 1 | Import the two .bas files and Debug > Compile VBAProject | No compile errors |
| 2 | Run InitializeInvoiceApp | Dashboard, Customers, Products, Audit Log |
| 3 | Run InitializeInvoiceApp twice | No loss of catalog records |
| 4 | Add first customer | CUS-0001 saved with contact info |
| 5 | Add first product | PRD-0001 saved with numeric price and demo tax |
| 6 | Refresh KPI | Counts match active catalog rows |
| 7 | Check Audit Log | CREATE_CUSTOMER and CREATE_PRODUCT |
| 8 | Add a second customer | ID increments; records retained |
| 9 | Invalid product price | User-friendly error; no record added |
| 10 | Blank customer name | Cancel without writing |
| 11 | Click navigation buttons | Appropriate worksheet opens |
| 12 | Save/close/reopen workbook | Records and UI preserved |
| 13 | Build with PowerShell | Output XLSM created, no overwrite |
| 14 | Test another Excel version | Must be explicitly verified |

Only use fictional sample data. Keep the scope restricted to catalog management; quotes and invoice documents are not yet built.
