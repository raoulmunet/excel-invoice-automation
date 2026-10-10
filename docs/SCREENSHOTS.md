# Real application screenshots for v1.0.0

**Status:** three real captures supplied by the owner on 2026-10-10, checked visually for readability and obvious personal information. Binary PNG files are not yet uploaded into the repository; the paths below are their intended destinations. Do **not** use stock pictures or generated mockups as evidence of the running application.

## Required captures

1. `docs/screenshots/dashboard-v1.0.png` — show the dark-blue Dashboard header, main actions and KPIs. Use the clean fictional demo workbook; no personal identifiers.
2. `docs/screenshots/quote-draft-v1.0.png` — in a separate testing copy, show Quote Draft: B4 = CUS-0001, B5 = 10, A10 = PRD-0001, B10 = 2, A11 = PRD-0002, B11 = 5. The real screenshot was supplied by the project owner.
3. `docs/screenshots/quotation-pdf-v1.0.png` — show the actual exported demonstration quotation (not a fiscal invoice): 2 × 150.00 plus 5 × 420.00, subtotal 2400.00, 10% discount 240.00, total 2160.00. This two-product screenshot is a separate scenario from the earlier single-product 270.00 acceptance test.

## Capture and privacy steps

- Use `ExcelInvoiceAutomation_TEST.xlsm` for transactions, never modify the untouched downloadable demo.
- In Windows 11 use Win+Shift+S, select the application area, and save as a PNG. Crop to the relevant window; do not include desktop notifications, usernames or real email addresses.
- Inspect each PNG at 100% for private information and legibility.
- Upload approved PNGs to the paths above (create `docs/screenshots/` as needed).
- Only after the files exist, embed the images in README with relative Markdown links:
  `![Dashboard](docs/screenshots/dashboard-v1.0.png)`,
  `![Quote Draft](docs/screenshots/quote-draft-v1.0.png)`, and
  `![Quotation export](docs/screenshots/quotation-pdf-v1.0.png)`.

Screenshots are supplemental documentation. The XLSM is a fictional, non-fiscal, single-user portfolio demo.