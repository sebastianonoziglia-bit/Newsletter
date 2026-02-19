# GitHub + Cloudflare Push Flow

## 1) Create GitHub repo (one time)

In GitHub web UI, create an empty repo (no README), for example:
- `newsletter-worker`

## 2) Connect local folder to that repo (one time)

```bash
cd /Users/sebbo/Desktop/Newsletter
git add .
git commit -m "Initial newsletter worker setup"
git remote add origin https://github.com/<YOUR_USER>/<YOUR_REPO>.git
git push -u origin main
```

If `origin` already exists:

```bash
git remote set-url origin https://github.com/<YOUR_USER>/<YOUR_REPO>.git
git push -u origin main
```

## 3) Daily workflow (just push)

```bash
cd /Users/sebbo/Desktop/Newsletter
git add .
git commit -m "Update newsletter"
git push
```

## 4) Connect repo to Cloudflare (Workers Builds)

Cloudflare Dashboard:
- Workers & Pages -> Overview -> `crimson-bar-9107`
- Settings -> Builds / Git integration
- Connect GitHub repo
- Branch: `main`

After that, each push to `main` triggers a deploy.

## 5) GitHub Actions auto-deploy (already included)

This repo includes:
- `/Users/sebbo/Desktop/Newsletter/.github/workflows/deploy-worker.yml`

In your GitHub repo settings, add these repository secrets:
- `CLOUDFLARE_API_TOKEN`
- `CLOUDFLARE_ACCOUNT_ID`

Then every push to `main` deploys automatically.

## Notes

- Worker config is in `wrangler.toml`.
- If you need API-based deploy from CI, set secrets:
  - `CLOUDFLARE_API_TOKEN`
  - `CLOUDFLARE_ACCOUNT_ID`
