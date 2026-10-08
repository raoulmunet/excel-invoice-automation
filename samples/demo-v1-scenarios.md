# Fictional Demo Scenarios for Excel Invoice Automation v1.0

All company names, contact addresses and amounts in this sample set are fictional. Addresses use reserved example domains and should **never** receive real messages. Do not import these rows into an existing workbook without checking for duplicate IDs. This dataset is a reference, not an executable workbook.

## Scenario A: quote, document and demo invoice

- Customer `CUS-0001` — Northwind Demo Services
- Product `PRD-0001` — Data reconciliation consulting, 150.00 per unit, zero demo tax
- Quote quantity: 2; discount: 10%
- Expected net 300.00, discount 30.00, tax 0.00, gross 270.00
- Export the quotation to Word and PDF; create a demo invoice from the quotation.

## Scenario B: two payments

- Starting demo invoice gross: 270.00
- Record 100.00 -> remaining balance 170.00, status Partially Paid (if not overdue)
- Record 170.00 -> remaining 0.00, status Paid
- Third positive payment should be rejected.

## Scenario C: overdue demonstration

- Create another demo invoice in a separate test workbook.
- Change its fictional due date to yesterday and refresh the payment dashboard.
- If balance remains positive, status should be Overdue.

## Scenario D: Outlook draft

- **Do not** email the reserved example.com/example.org/example.net addresses.
- For an actual Outlook integration test, set a fictional customer's contact email to a mailbox **you own**, create a draft and inspect recipient, subject, message and PDF. Sending must remain a deliberate manual action.

## Important

The demo has no defined currency or fiscal compliance. It illustrates a software workflow, not tax or accounting correctness. The CSV is informational only: it does not automatically populate or generate quotations, invoices or payments. Preserve your tested workbook when preparing screenshots.
