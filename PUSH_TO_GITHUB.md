# Push this project to GitHub

Takes about 10 minutes. Do it once; the dashboard and the daily refresh run themselves after that.

## 1. Create the repo

1. Go to https://github.com/new
2. Repository name: `sector-teardown-20` (or anything you like)
3. Visibility: **Public** (Pages + free Actions minutes work best public)
4. Do NOT initialize with a README — the folder already has everything
5. Click **Create repository**, then copy the repo URL

## 2. Push from this machine

```bash
cd ~/workspace/sector-teardown-20
git init -b main
git add -A
git commit -m "20-company AI-augmented sector teardown + live dashboard"
git remote add origin https://github.com/<your-username>/sector-teardown-20.git
git push -u origin main
```

(Replace `<your-username>` with your GitHub username.)

## 3. Enable GitHub Pages

1. In the repo, go to **Settings → Pages**
2. Under **Build and deployment → Source**, select **GitHub Actions**
3. Push any small change (or re-run the workflow) — the **Deploy dashboard to Pages** workflow publishes `site/`
4. Your dashboard URL appears under Settings → Pages and on each deploy run

## 4. Enable the daily price refresh

1. Go to the **Actions** tab → click **Daily price refresh** → **Run workflow** (verifies it works on day one)
2. It then runs automatically weekdays at 17:30 ET, updating `site/data/companies.json` and redeploying
3. If the push step fails with a permissions error: **Settings → Actions → General → Workflow permissions** → select **Read and write permissions** → Save

## What stays fresh automatically

- **Prices:** refreshed every weekday after close by the Action; the dashboard also pulls a live quote (~15 min delayed) in the visitor's browser
- **DCF valuations:** fixed at the model date by design — fair values only change when you rebuild a company's model on new fundamentals (rerun that company's `pull_data.py` + `build_model.py`, then `python3 build_site_data.py` and push)

## Repo layout

| Path | What it is |
|---|---|
| `site/` | The dashboard (deployed to Pages) |
| `site/data/companies.json` | Pipeline-generated data contract for the dashboard |
| `companies/<TICKER>/` | Per-company model (.xlsx), memo (.md), raw data, rerun scripts |
| `SECTOR_REPORT.md` / `.pdf` | The flagship written sector report |
| `build_site_data.py` | Rebuilds `companies.json` from the 20 `summary.json` files |
| `refresh_pipeline.py` | Daily price refresh used by the Action |
| `.github/workflows/` | `pages.yml` (deploy) and `refresh.yml` (daily prices) |
