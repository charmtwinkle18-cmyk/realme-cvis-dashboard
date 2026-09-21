# realme CVIS Sales Dashboard

A Vercel-ready Next.js dashboard styled to match the CVIS/WVIS sales intelligence layout.

## Included
- Dark CVIS sidebar navigation
- Excel upload and automatic dashboard refresh
- Embedded `public/default-data.xlsx` so the dashboard loads immediately after deployment
- Month / area / subregion / customer / channel / model / price-range / series filters
- KPI cards
- Sales and dealer trend charts
- Model, dealer, subregion, store, promoter, and area ranking pages
- Night mode

## Run locally
```bash
npm install
npm run dev
```
Open http://localhost:3000

## Deploy to Vercel
The files in this ZIP are already at the project root. `package.json`, `app`, and `public` must remain at the same root level.

If this project is uploaded into a subfolder on GitHub, set that folder as the Vercel **Root Directory**.

## Workbook requirements
The dashboard expects these sheets when available:
- Sales_Data
- Sales_Target
- Inventory
- Promoter_Data
- Store_Master
- Model_Master
- Area_Master

`Sales_Data` supports your current columns including `Customer Type` (used as Channel when `Channel` is absent).
