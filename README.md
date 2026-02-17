# WASDE Dashboard — Innovative Grain

Interactive visualization of the USDA WASDE (World Agricultural Supply and Demand Estimates) report for Corn, Soybeans, and Wheat.

## 🌐 Live Dashboard

Hosted on GitHub Pages and embedded on [innovativegrain.ai](https://www.innovativegrain.ai).

## 🔄 Auto-Update

A GitHub Actions workflow runs on the 11th-12th of each month (the day after WASDE typically releases) to:
1. Download the latest WASDE Excel file from USDA
2. Extract Corn, Soybeans, and Wheat data
3. Update the dashboard HTML
4. Deploy to GitHub Pages

You can also trigger manually: Actions → "Update WASDE Dashboard" → Run workflow.

## 🛠 Manual Update

```bash
pip install xlrd
python update_wasde.py --year 2026 --month 3
```

## 📋 Embedding on GoDaddy

Add an **HTML embed section** on your GoDaddy site with:

```html
<iframe src="https://YOUR-GITHUB-USERNAME.github.io/wasde-dashboard/"
        width="100%" height="2400" frameborder="0"
        style="border:none; border-radius:8px;">
</iframe>
```

## Data Source

[USDA WASDE Report](https://www.usda.gov/oce/commodity/wasde) — Released monthly by the World Agricultural Outlook Board.
