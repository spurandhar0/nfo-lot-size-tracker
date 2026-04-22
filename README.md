# 📊 NSE NFO Lot Size Tracker

Automatically fetches **all NSE F&O stocks + indices** with **lot sizes, LTP, and contract values** daily via GitHub Actions and saves the data into a formatted Excel file committed back to this repo.

---

## 📁 Repository Structure

```
nfo-lot-size-tracker/
├── .github/
│   └── workflows/
│       └── nfo_daily_update.yml   ← GitHub Actions workflow
├── scripts/
│   └── fetch_nfo_data.py          ← Main Python fetcher
├── data/
│   ├── nfo_lot_sizes.xlsx         ← 📊 Auto-updated Excel (output)
│   └── nfo_lot_sizes.json         ← Raw JSON snapshot (output)
├── requirements.txt
└── README.md
```

---

## 📊 Excel Output — Sheet Overview

| Sheet | Contents |
|---|---|
| 📊 Summary | Quick stats + Index reference table |
| 📈 F&O Stocks | All ~200 F&O stocks with lot size, LTP, contract value |
| 🏛️ Indices | Nifty 50, BankNifty, FinNifty, MidCap, Sensex etc. |
| 📋 Change Log | Timestamped history of every run |

**Columns in Stock/Index sheets:**

| Column | Description |
|---|---|
| Symbol | NSE ticker symbol |
| Name | Full company / index name |
| Lot Size | Minimum contract quantity |
| LTP (₹) | Last traded price |
| Change (%) | Daily % change (green = positive, orange = negative) |
| Contract Value (₹) | Auto-calculated: `Lot Size × LTP` |
| Type | Stock or Index |

---

## ⚙️ How It Works

1. **GitHub Actions** runs every weekday at **6:30 AM IST** (Mon–Fri)
2. Python script hits NSE APIs (no API key needed — public endpoints)
3. Builds / updates the Excel with fresh data
4. Commits `data/nfo_lot_sizes.xlsx` and `data/nfo_lot_sizes.json` back to the repo

---

## 🚀 Setup Instructions

### Step 1 — Fork / Clone this repo

```bash
git clone https://github.com/<your-username>/nfo-lot-size-tracker.git
cd nfo-lot-size-tracker
```

### Step 2 — Enable GitHub Actions

Go to your repo → **Actions** tab → click **"I understand my workflows, go ahead and enable them"**

> ✅ No secrets needed — the workflow uses the built-in `GITHUB_TOKEN`.

### Step 3 — Give Actions write permission

Go to **Settings → Actions → General → Workflow permissions** → select **"Read and write permissions"** → Save.

### Step 4 — Run manually (first time)

Go to **Actions → NSE NFO Daily Update → Run workflow** → click **Run workflow**.

---

## 🖥️ Run Locally

```bash
# Install dependencies
pip install -r requirements.txt

# Run the script
python scripts/fetch_nfo_data.py
```

Output will be saved to `data/nfo_lot_sizes.xlsx`.

---

## 🔄 Schedule

The workflow runs automatically:
- **Every weekday (Mon–Fri) at 6:30 AM IST**
- **Manually** anytime via Actions → Run workflow

To change the schedule, edit the `cron` line in `.github/workflows/nfo_daily_update.yml`:

```yaml
- cron: "0 1 * * 1-5"   # 01:00 UTC = 06:30 IST, Mon–Fri
```

---

## 📌 Notes

- NSE blocks non-browser requests — the script warms up a session via the NSE homepage before making API calls.
- Lot sizes are updated by NSE/SEBI periodically; the script always fetches the latest.
- If NSE changes their API structure, update the endpoint URLs in `scripts/fetch_nfo_data.py`.

---

## 📜 License

MIT — free to use and modify.
