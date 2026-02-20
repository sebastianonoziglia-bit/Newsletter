# Newsletter Workflow (Excel/Google Sheets -> HTML)

This setup lets you edit a single data source and regenerate the newsletter in the same style.

## Files

- `/Users/sebbo/Desktop/Newsletter/build_newsletter.py` - generator script
- `/Users/sebbo/Desktop/Newsletter/generate_newsletter.sh` - one-command launcher (Google Sheets default)
- `/Users/sebbo/Desktop/Newsletter/src/worker.js` - Cloudflare Worker (online auto-updating newsletter)
- `/Users/sebbo/Desktop/Newsletter/wrangler.toml` - Cloudflare Worker config
- `/Users/sebbo/Desktop/Newsletter/newsletter_data.xlsx` - optional local Excel source
- `/Users/sebbo/Desktop/Newsletter/newsletter.html` - generated output
- `/Users/sebbo/Desktop/Newsletter/history/` - automatic source snapshots/backups

## Data Structure (same for Excel and Google Sheets)

Use these tabs and columns:

1. `meta`
- `key`
- `value`

2. `points`
- `order` (1..10)
- `title`
- `content`
- `image_path` (optional)
- `image_caption` (optional)
- `source` (optional)

3. `distribution` (optional)
- `category`
- `amount_btc`
- `percent` (optional; can be blank)
- `color` (CSS color like `rgb(255, 66, 2)`)

## Edit Weekly / Monthly

1. Update `meta` values (`main_title`, `subtitle`, `block_height`, `max_supply_btc`, `circulating_supply_btc`, `hashrate_eh_s`, `hashrate_scale_eh_s`, `snapshot_*`, `tldr_*`, `conclusion_*`, `cta_*`, `address_line`, `footer_line`).
2. Update `points` rows (max 10, unique `order`).
3. Optionally update `distribution` rows.
4. Optional auto-images:
- Put files as `1.png`, `2.png`, ... `10.png` in `/Users/sebbo/Desktop/Newsletter/`
- If `image_path` is empty, the script auto-uses the matching number if the file exists.
- Extra images per point are supported at the bottom of that point section with names like `1.1.png`, `1.2.png`, `2.1.png`, `3.1.png`.

## Rebuild Newsletter

Default output location:
- `/Users/sebbo/Desktop/Newsletter/newsletter.html` (always, unless you pass `--out`)

Fast app mode (recommended):

```bash
/Users/sebbo/Desktop/Newsletter/generate_newsletter.sh
```

Watch mode (auto-regenerate every 20s):

```bash
/Users/sebbo/Desktop/Newsletter/generate_newsletter.sh --watch
```

Watch mode with custom interval:

```bash
/Users/sebbo/Desktop/Newsletter/generate_newsletter.sh --watch --interval 10
```

Optional browser auto-refresh for preview:
- Open this URL in browser: `file:///Users/sebbo/Desktop/Newsletter/newsletter.html?refresh=10`
- The page will reload every 10 seconds (minimum supported is 5).

## Online Auto-Update (Cloudflare Worker)

- See `/Users/sebbo/Desktop/Newsletter/CLOUDFLARE_WORKER_SETUP.md`.
- Worker URL: `https://crimson-bar-9107.sebastiano-noziglia.workers.dev/`
- Online mode updates automatically from Google Sheets with a 2-minute edge cache.

### From local Excel

```bash
python3 /Users/sebbo/Desktop/Newsletter/build_newsletter.py
```

### From Google Sheets (recommended)

```bash
python3 /Users/sebbo/Desktop/Newsletter/build_newsletter.py \
  --google-sheet "https://docs.google.com/spreadsheets/d/YOUR_SHEET_ID/edit"
```

Optional tab-name overrides:

```bash
python3 /Users/sebbo/Desktop/Newsletter/build_newsletter.py \
  --google-sheet "YOUR_SHEET_ID" \
  --google-meta-tab "meta" \
  --google-points-tab "points" \
  --google-distribution-tab "distribution"
```

Google Sheets access requirement:
- Share as at least `Anyone with the link can view`, or use a sheet the current machine can fetch publicly.

Important behavior:
- `newsletter.html` is a static snapshot file.
- After changing Google Sheets, regenerate (`generate_newsletter.sh`) and then refresh browser to see updates.

## PDF Download (from newsletter.html)

When you open `/Users/sebbo/Desktop/Newsletter/newsletter.html` in a browser, there is now a `Download PDF` button at the top.

- Click `Download PDF`
- Browser print dialog opens
- Choose `Save as PDF`

## Automatic History

- Excel mode: saves `/Users/sebbo/Desktop/Newsletter/history/newsletter_YYYY-MM-DD_HHMM.xlsx`
- Google Sheets mode: saves one workbook snapshot:
  - `/Users/sebbo/Desktop/Newsletter/history/newsletter_YYYY-MM-DD_HHMM.xlsx`
  - with tabs `meta`, `points`, and `distribution` (when present)

## New Template File (if needed)

```bash
python3 /Users/sebbo/Desktop/Newsletter/build_newsletter.py --init-template --force
```
