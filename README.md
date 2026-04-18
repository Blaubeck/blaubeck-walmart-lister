# BLAUBECK Walmart Lister

Tool to bulk-list Amazon products on Walmart Marketplace for BLAUBECK (Lackus LLC).

## Pipeline

1. Pull 30-day Amazon sales -> identify active SKUs
2. Diff against current Walmart catalog -> find missing
3. Fetch Amazon listing data (title, images, price, UPC)
4. Generate Walmart-optimized copy (titles, bullets, descriptions)
5. Build `MP_ITEM` feed with all data
6. Submit via Walmart Feeds API, poll status

## Status (2026-04-18)

- 11 SKUs queued for Walmart listing
- 10 UPCs needed from GS1 UK (planned Monday 2026-04-20, GS1 UK is in maintenance)
- Images downloaded and committed to `data/images/` for public URLs
- Copy prepared in `data/walmart_copy.json`

## Ready-to-list products

See `data/review_master.csv` for the final candidate list.
