# csv-to-slides-webapp

Web app to convert **CSV/XLSX** input into a downloadable **PPTX** deck (one row per slide).

## Features

- Upload `.csv`, `.xls`, or `.xlsx`
- One row -> one slide
- Slides are ordered by publication date from oldest to newest
- Dark theme slide template (16:9)
- Full-width title, metadata block, body text, right-side cover image
- Clickable deliverable hyperlink: `https://guilds.reply.com/news/[Id]`
- Browser-side generation (no backend)
- Automatic image-proxy fallback for hosts that block direct browser fetches (for example `guilds-cdn.reply.com`)

## Expected columns

- `Id`
- `Title`
- `Associated Lance` (also accepts `Associated Lances` / `Associated Lens`)
- `Associated Deliverable`
- `Publication Date`
- `Cover` (also accepts `Image` / `Cover URL`)
- `Text` (HTML supported)

## Run locally

Just open `index.html` in a browser.

## Deploy

Push to `main` branch; GitHub Actions deploys to GitHub Pages.
The default output file name is `guilds_highlights_<month>_<year>.pptx`.

## Notes on images

Some image hosts, including Reply's Guilds CDN, do not expose CORS headers needed for direct browser fetches from GitHub Pages. The web app now falls back to `images.weserv.nl` for those URLs so cover images can still be embedded into the downloaded PPTX while keeping the app fully static.
