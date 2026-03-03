# csv-to-slides-webapp

Web app to convert **CSV/XLSX** input into a downloadable **PPTX** deck (one row per slide).

## Features

- Upload `.csv`, `.xls`, or `.xlsx`
- One row -> one slide
- Dark theme slide template (16:9)
- Full-width title, metadata block, body text, image placeholder
- Clickable deliverable hyperlink: `https://guilds.reply.com/news/[Id]`
- Browser-side generation (no backend)

## Expected columns

- `Id`
- `Title`
- `Associated Lance` (also accepts `Associated Lances` / `Associated Lens`)
- `Associated Deliverable`
- `Publication Date`
- `Text` (HTML supported)

## Run locally

Just open `index.html` in a browser.

## Deploy

Push to `main` branch; GitHub Actions deploys to GitHub Pages.
