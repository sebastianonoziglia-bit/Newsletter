# Cloudflare Worker Setup (Auto-Updated Newsletter)

This project now includes a Worker that renders the newsletter directly from Google Sheets and caches it for 2 minutes.

## Included config

- Worker name: `crimson-bar-9107`
- Worker entry: `/Users/sebbo/Desktop/Newsletter/src/worker.js`
- Wrangler config: `/Users/sebbo/Desktop/Newsletter/wrangler.toml`
- Static assets: `/Users/sebbo/Desktop/Newsletter/public/` (logo + point images)
- Google Sheet source: `1ukXTu8PXWHGe4Fzg5rA424BmN6Bi3FPAWwA12QybNpI`
- Tab names: `live_prices`, `meta`, `points`, `distribution`
- Refresh window: `120` seconds (`CACHE_TTL_SECONDS`)

## Deploy

```bash
cd /Users/sebbo/Desktop/Newsletter
npx wrangler login
npx wrangler deploy
```

If you are already logged in to Wrangler, only run `npx wrangler deploy`.

## Verify

- Main page: `https://crimson-bar-9107.sebastiano-noziglia.workers.dev/`
- Health endpoint: `https://crimson-bar-9107.sebastiano-noziglia.workers.dev/health`

## How updates work

- Worker fetches Google Sheets tabs and renders HTML on request.
- Response is cached at edge for 2 minutes.
- New sheet changes show up automatically after cache expiry.

Force-refresh one request immediately:

```text
https://crimson-bar-9107.sebastiano-noziglia.workers.dev/?force=1
```

## Change update interval

Edit in `/Users/sebbo/Desktop/Newsletter/wrangler.toml`:

```toml
CACHE_TTL_SECONDS = "120"
```

Examples:
- `120` = 2 minutes
- `300` = 5 minutes

After changing it, redeploy:

```bash
cd /Users/sebbo/Desktop/Newsletter
npx wrangler deploy
```
