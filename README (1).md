# 📊 Company Dashboard Tracker

Automated financial dashboard generator for your watchlist companies from Screener.in

## ✨ Features

- 📈 Scrapes financial data from your Screener.in watchlists
- 📊 Creates interactive Excel dashboard with dropdown menu
- 🎨 Auto-updating charts and metrics
- ☁️ Runs automatically via GitHub Actions (daily at 6 PM IST)
- 📧 Optional email notifications
- 💾 Backup data to Google Sheets

## 🎯 What It Creates

An Excel file with:
- **Dashboard Sheet**: Interactive dropdown to select any company
- **Auto-updating metrics**: Price, Market Cap, P/E, Sector
- **Raw Data Sheet**: All companies data
- **Company List**: For dropdown (hidden)

### How It Works

```
Select Company: [Indiamart ▼]

Key Metrics:
  Current Price: ₹2,157
  Market Cap: ₹12,956 Cr
  P/E Ratio: 24.13
  Sector: Trading Companies and Distributors
```

## 🚀 Setup Instructions

### 1. Prerequisites

- GitHub account
- Screener.in account (with watchlists created)
- Google Cloud service account (same as Concall Tracker)

### 2. Create Repository

1. Create new repository: `company-dashboard-tracker`
2. Public repository
3. Add README

### 3. Add GitHub Secrets

Go to Settings → Secrets and variables → Actions

Add these secrets:

```
SCREENER_USERNAME = your.email@gmail.com
SCREENER_PASSWORD = your_screener_password
GOOGLE_SHEET_ID = your_sheet_id
GOOGLE_CREDENTIALS_BASE64 = your_base64_credentials
MY_STONKS_WATCHLIST_URL = https://www.screener.in/watchlist/YOUR_ID/
CORE_WATCHLIST_URL = https://www.screener.in/watchlist/YOUR_ID/
```

### 4. Upload Files

Upload these files to your repository:
- `company_dashboard_tracker.py`
- `requirements.txt`
- `.github/workflows/company-scraper.yml`

### 5. Create Google Sheet

1. Create new Google Sheet named "Company Dashboard Data"
2. Share with service account (Editor permission)
3. Copy Sheet ID from URL
4. Add as GitHub secret

### 6. Run!

- Go to Actions tab
- Click "Company Dashboard Scraper"
- Click "Run workflow"
- Wait 5-10 minutes
- Download Excel file from Artifacts

## 📊 How to Use the Excel Dashboard

1. Download the Excel file
2. Open it
3. Go to "Dashboard" sheet
4. Click the dropdown in cell B3
5. Select any company
6. All metrics update automatically!

## ⚙️ Customization

### Change Schedule

Edit `.github/workflows/company-scraper.yml`:

```yaml
schedule:
  - cron: '30 12 * * *'  # Daily at 6 PM IST
```

### Add More Watchlists

Add more secrets:
```
WATCHLIST_3_URL = https://www.screener.in/watchlist/...
```

Then update the code to include them.

## 📁 Files

- `company_dashboard_tracker.py` - Main scraper script
- `requirements.txt` - Python dependencies
- `.github/workflows/company-scraper.yml` - GitHub Actions workflow

## 🎨 Features Coming Soon

- [ ] Revenue trend charts
- [ ] Profit margin charts
- [ ] Quarterly performance charts
- [ ] Comparison between companies
- [ ] Email delivery of Excel file
- [ ] Historical data tracking

## 📧 Support

Create an issue if you need help!

## 📝 License

MIT

---

Built with ❤️ using Python, Selenium, and openpyxl
