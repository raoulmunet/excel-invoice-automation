# v0.3 Word/PDF and Outlook Manual Acceptance Tests

**Target environment:** VirtualBox Windows 11; Excel 2019, Word 2019 and (optional) Classic Outlook Desktop.
**Status:** code on GitHub; **no user runtime verification yet**. v0.2 was owner-confirmed before these modifications.

## Install (upgrade without losing data)
1. Back up the working v0.2 `.xlsm`.
2. Replace the old `modSetup` with the latest `src/modSetup.bas`.
3. Import `src/modDocuments.bas`; keep existing `modCatalog` and `modQuotations`.
4. **Debug > Compile VBAProject**. On error, record the exact message and highlighted line.
5. Run `InitializeInvoiceApp` from `Alt+F8` to rebuild Dashboard buttons and create `Demo Invoices`.

## Demonstration

Use the previously saved quote `QUO-0001` from `CUS-0001` and `PRD-0001`.
1. Click **Quote Word/PDF**; enter `QUO-0001`; choose a new output DOCX filename.
2. Expect both `.docx` and `.pdf` in the chosen folder. Open both files and check **DEMO** banner, quote ID, customer, product lines, quantities, discounted totals, demo tax and gross.
3. **Do not re-use an existing name:** application should refuse an overwrite.
4. Click **Demo Invoice**, enter `QUO-0001`; expect one row in `Demo Invoices` with `DEMO-INV-0001` and status `DEMO - NOT FISCAL`.
5. Repeat the same quote: app should refuse a second demo invoice for that quote.
6. If **Classic Outlook** is installed and configured, click **Outlook Draft**; enter `QUO-0001`; choose the PDF, review the recipient and attachment. The mail must **not** be sent automatically.
7. Confirm the Audit Log contains EXPORT_QUOTE_DOCS, CREATE_DEMO_INVOICE and (if used) PREPARE_OUTLOOK_DRAFT.
8. Save, close and reopen the XLSM; check persistence and the v0.2 flows.

## Negative tests
- Unknown quote ID must return a clear message, not create documents.
- Missing Word Desktop: report explanatory automation error, not silently succeed.
- Missing Classic Outlook: error only for that optional feature.
- Cancel Save As and cancel PDF picker: no output/no send.
- DOCX or PDF target already exists: do not overwrite.
- Invalid/missing customer email: confirm Outlook To field before any manual sending.
- Documents must remain clearly labeled demonstration; no fiscal/legal compliance implied.
- Initial workbook with no quotations: export must not produce output.
- Record any Word formatting/column-width/long-description issues found.

## Release gate

Do not advertise Word/PDF/Outlook features as tested until explicitly verified by the user in Office 2019. Real fiscal invoicing, localized tax rules, numbering compliance and Romanian RO e-Factura are **out of scope**.
