"""
NSE NFO Lot Size & Price Fetcher
Fetches all F&O stocks + indices with lot sizes and LTP from NSE,
then writes/updates an Excel file in data/nfo_lot_sizes.xlsx
"""

import requests
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime
import json
import os
import time

# ── NSE session headers (required to avoid 403) ─────────────────────────────
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Accept-Language": "en-US,en;q=0.9",
    "Accept-Encoding": "gzip, deflate, br",
    "Referer": "https://www.nseindia.com/",
    "Connection": "keep-alive",
}

NSE_BASE = "https://www.nseindia.com"
OUTPUT_PATH = os.path.join(os.path.dirname(__file__), "..", "data", "nfo_lot_sizes.xlsx")


def get_nse_session() -> requests.Session:
    """Create a warmed-up NSE session (hits homepage first to get cookies)."""
    session = requests.Session()
    session.headers.update(HEADERS)
    session.get(NSE_BASE, timeout=15)
    time.sleep(1)
    return session


def fetch_fno_stocks(session: requests.Session) -> pd.DataFrame:
    """Fetch all F&O equity stocks with lot size and LTP."""
    url = f"{NSE_BASE}/api/equity-stockIndices?index=SECURITIES%20IN%20F%26O"
    resp = session.get(url, timeout=15)
    resp.raise_for_status()
    data = resp.json()

    rows = []
    for item in data.get("data", []):
        symbol = item.get("symbol", "")
        if not symbol or symbol in ("NIFTY 50", "NIFTY BANK"):
            continue
        rows.append({
            "Symbol": symbol,
            "Company Name": item.get("meta", {}).get("companyName", item.get("symbol", "")),
            "LTP (₹)": item.get("lastPrice", 0),
            "Change (%)": item.get("pChange", 0),
            "52W High": item.get("yearHigh", 0),
            "52W Low": item.get("yearLow", 0),
        })

    df_ltp = pd.DataFrame(rows)

    # Fetch lot sizes from F&O participant data
    lot_url = f"{NSE_BASE}/api/master-quote"
    # Use the derivative market watch for lot sizes
    lot_data = fetch_lot_sizes(session)

    if not lot_data.empty and not df_ltp.empty:
        df = pd.merge(df_ltp, lot_data, on="Symbol", how="left")
    elif not df_ltp.empty:
        df = df_ltp
        df["Lot Size"] = "-"
    else:
        df = lot_data

    df["Type"] = "Stock"
    return df


def fetch_lot_sizes(session: requests.Session) -> pd.DataFrame:
    """Fetch lot sizes from NSE F&O market watch."""
    url = f"{NSE_BASE}/api/option-chain-equities?symbol=NIFTY"
    # We use the derivatives snapshot endpoint for lot sizes
    snap_url = f"{NSE_BASE}/api/equity-derivatives-snapshot"
    try:
        resp = session.get(snap_url, timeout=15)
        if resp.status_code == 200:
            data = resp.json()
            rows = []
            for item in data.get("data", []):
                symbol = item.get("underlying", "")
                lot_size = item.get("marketLot", item.get("lotSize", "-"))
                if symbol:
                    rows.append({"Symbol": symbol, "Lot Size": lot_size})
            if rows:
                return pd.DataFrame(rows).drop_duplicates("Symbol")
    except Exception:
        pass

    # Fallback: parse from F&O bhavcopy metadata
    try:
        bhavcopy_url = f"{NSE_BASE}/api/equity-derivatives-bhavcopy-download"
        resp = session.get(f"{NSE_BASE}/api/fo-mktlots", timeout=15)
        if resp.status_code == 200:
            data = resp.json()
            rows = []
            for item in data:
                symbol = item.get("symbol", "").strip()
                lot = item.get("lot_size", item.get("lotSize", "-"))
                if symbol:
                    rows.append({"Symbol": symbol, "Lot Size": lot})
            return pd.DataFrame(rows).drop_duplicates("Symbol")
    except Exception:
        pass

    return pd.DataFrame(columns=["Symbol", "Lot Size"])


def fetch_index_data(session: requests.Session) -> pd.DataFrame:
    """Fetch index F&O instruments (Nifty, BankNifty, FinNifty, etc.)."""
    indices = [
        ("NIFTY 50",      "NIFTY",     65),
        ("NIFTY BANK",    "BANKNIFTY", 30),
        ("NIFTY FIN SVC", "FINNIFTY",  40),
        ("NIFTY MID 50",  "MIDCPNIFTY",120),
        ("NIFTY NEXT 50", "NIFTYNXT50",25),
        ("SENSEX",        "SENSEX",    20),
        ("BANKEX",        "BANKEX",    15),
    ]

    rows = []
    for display_name, symbol, default_lot in indices:
        url = f"{NSE_BASE}/api/allIndices"
        try:
            resp = session.get(url, timeout=15)
            resp.raise_for_status()
            all_indices = resp.json().get("data", [])
            matched = next((i for i in all_indices if display_name in i.get("index", "")), None)
            ltp = matched.get("last", 0) if matched else 0
            change = matched.get("percentChange", 0) if matched else 0
        except Exception:
            ltp, change = 0, 0

        rows.append({
            "Symbol": symbol,
            "Company Name": display_name,
            "LTP (₹)": ltp,
            "Change (%)": change,
            "52W High": 0,
            "52W Low": 0,
            "Lot Size": default_lot,
            "Type": "Index",
        })

    return pd.DataFrame(rows)


def build_excel(df_stocks: pd.DataFrame, df_indices: pd.DataFrame, output_path: str):
    """Build a formatted Excel workbook with Summary, Stocks, and Indices sheets."""
    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    wb = Workbook()
    wb.remove(wb.active)  # remove default sheet

    today = datetime.now().strftime("%d-%b-%Y %I:%M %p")

    # ── colour palette ───────────────────────────────────────────────────────
    NAVY    = "1F3864"
    GOLD    = "C9A84C"
    LIGHT   = "EEF2FF"
    WHITE   = "FFFFFF"
    GREEN   = "E2EFDA"
    ORANGE  = "FCE4D6"

    def header_style(cell, bg=NAVY, fg=WHITE, bold=True, size=10):
        cell.font = Font(name="Arial", bold=bold, color=fg, size=size)
        cell.fill = PatternFill("solid", start_color=bg)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    def thin_border():
        s = Side(style="thin", color="CCCCCC")
        return Border(left=s, right=s, top=s, bottom=s)

    def write_sheet(ws, df, title, sheet_type="stock"):
        # Title row
        ws.merge_cells("A1:H1")
        ws["A1"] = f"NSE F&O {title} — Last Updated: {today}"
        ws["A1"].font = Font(name="Arial", bold=True, size=13, color=WHITE)
        ws["A1"].fill = PatternFill("solid", start_color=NAVY)
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 30

        # Column headers
        cols = ["#", "Symbol", "Name", "Lot Size", "LTP (₹)", "Change (%)", "Contract Value (₹)", "Type"]
        for ci, col in enumerate(cols, 1):
            cell = ws.cell(row=2, column=ci, value=col)
            header_style(cell, bg=GOLD, fg=NAVY)
        ws.row_dimensions[2].height = 22

        # Data rows
        for ri, (_, row) in enumerate(df.iterrows(), 3):
            bg = WHITE if ri % 2 == 0 else LIGHT
            ws.cell(row=ri, column=1, value=ri - 2)
            ws.cell(row=ri, column=2, value=row.get("Symbol", ""))
            ws.cell(row=ri, column=3, value=row.get("Company Name", ""))
            ws.cell(row=ri, column=4, value=row.get("Lot Size", "-"))
            ws.cell(row=ri, column=5, value=row.get("LTP (₹)", 0))
            ws.cell(row=ri, column=6, value=row.get("Change (%)", 0))

            # Contract value formula = LotSize * LTP
            lot_col = get_column_letter(4)
            ltp_col = get_column_letter(5)
            lot_cell = f"{lot_col}{ri}"
            ltp_cell = f"{ltp_col}{ri}"
            ws.cell(row=ri, column=7, value=f'=IF(ISNUMBER({lot_cell}),{lot_cell}*{ltp_cell},"-")')

            ws.cell(row=ri, column=8, value=row.get("Type", "Stock"))

            for ci in range(1, 9):
                c = ws.cell(row=ri, column=ci)
                c.fill = PatternFill("solid", start_color=bg)
                c.border = thin_border()
                c.font = Font(name="Arial", size=9)
                c.alignment = Alignment(horizontal="center", vertical="center")

            # Colour-code Change %
            chg = row.get("Change (%)", 0)
            try:
                chg = float(chg)
                ws.cell(row=ri, column=6).fill = PatternFill(
                    "solid", start_color=GREEN if chg >= 0 else ORANGE
                )
            except (ValueError, TypeError):
                pass

        # Column widths
        widths = [5, 16, 34, 10, 12, 12, 20, 10]
        for ci, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(ci)].width = w

        # Freeze top 2 rows
        ws.freeze_panes = "A3"

        # Auto-filter on header row
        ws.auto_filter.ref = f"A2:{get_column_letter(len(cols))}2"

    # ── Sheet 1: Summary ─────────────────────────────────────────────────────
    ws_sum = wb.create_sheet("📊 Summary")
    ws_sum.sheet_view.showGridLines = False

    summary_data = [
        ("Total F&O Stocks", len(df_stocks)),
        ("Total Indices", len(df_indices)),
        ("Total Instruments", len(df_stocks) + len(df_indices)),
        ("Last Updated", today),
        ("Data Source", "NSE India (www.nseindia.com)"),
    ]

    ws_sum.merge_cells("A1:D1")
    ws_sum["A1"] = "NSE NFO — Daily Snapshot"
    ws_sum["A1"].font = Font(name="Arial", bold=True, size=16, color=WHITE)
    ws_sum["A1"].fill = PatternFill("solid", start_color=NAVY)
    ws_sum["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws_sum.row_dimensions[1].height = 36

    for ri, (label, value) in enumerate(summary_data, 3):
        ws_sum.cell(row=ri, column=1, value=label).font = Font(name="Arial", bold=True, size=10)
        ws_sum.cell(row=ri, column=2, value=value).font = Font(name="Arial", size=10)
        ws_sum.cell(row=ri, column=1).fill = PatternFill("solid", start_color=LIGHT)
        ws_sum.cell(row=ri, column=2).fill = PatternFill("solid", start_color=WHITE)
        ws_sum.row_dimensions[ri].height = 18

    ws_sum.column_dimensions["A"].width = 25
    ws_sum.column_dimensions["B"].width = 35

    # Index quick-ref table
    ws_sum.cell(row=9, column=1, value="Index Quick Reference").font = Font(
        name="Arial", bold=True, size=11, color=WHITE
    )
    ws_sum.cell(row=9, column=1).fill = PatternFill("solid", start_color=NAVY)
    ws_sum.merge_cells("A9:D9")

    idx_headers = ["Index", "Lot Size", "LTP (₹)", "Change (%)"]
    for ci, h in enumerate(idx_headers, 1):
        c = ws_sum.cell(row=10, column=ci, value=h)
        header_style(c, bg=GOLD, fg=NAVY)

    for ri, (_, row) in enumerate(df_indices.iterrows(), 11):
        ws_sum.cell(row=ri, column=1, value=row.get("Symbol", ""))
        ws_sum.cell(row=ri, column=2, value=row.get("Lot Size", "-"))
        ws_sum.cell(row=ri, column=3, value=row.get("LTP (₹)", 0))
        ws_sum.cell(row=ri, column=4, value=row.get("Change (%)", 0))
        for ci in range(1, 5):
            c = ws_sum.cell(row=ri, column=ci)
            c.font = Font(name="Arial", size=9)
            c.alignment = Alignment(horizontal="center")
            c.border = thin_border()

    # ── Sheet 2: Stocks ───────────────────────────────────────────────────────
    ws_stocks = wb.create_sheet("📈 F&O Stocks")
    ws_stocks.sheet_view.showGridLines = False
    write_sheet(ws_stocks, df_stocks, "Stocks", "stock")

    # ── Sheet 3: Indices ──────────────────────────────────────────────────────
    ws_idx = wb.create_sheet("🏛️ Indices")
    ws_idx.sheet_view.showGridLines = False
    write_sheet(ws_idx, df_indices, "Indices", "index")

    # ── Sheet 4: Change Log ───────────────────────────────────────────────────
    ws_log = wb.create_sheet("📋 Change Log")
    ws_log["A1"] = "Date"
    ws_log["B1"] = "Total Stocks"
    ws_log["C1"] = "Total Indices"
    ws_log["D1"] = "Notes"
    for ci in range(1, 5):
        header_style(ws_log.cell(row=1, column=ci), bg=NAVY)
    ws_log.append([today, len(df_stocks), len(df_indices), "Auto-updated by GitHub Actions"])

    wb.save(output_path)
    print(f"✅ Excel saved → {output_path}")
    return output_path


def update_change_log(output_path: str, df_stocks: pd.DataFrame, df_indices: pd.DataFrame):
    """Append a row to the Change Log sheet on subsequent runs."""
    today = datetime.now().strftime("%d-%b-%Y %I:%M %p")
    try:
        wb = load_workbook(output_path)
        if "📋 Change Log" in wb.sheetnames:
            ws = wb["📋 Change Log"]
            ws.append([today, len(df_stocks), len(df_indices), "Auto-updated by GitHub Actions"])
            wb.save(output_path)
    except Exception as e:
        print(f"⚠️  Could not update change log: {e}")


def main():
    print("🚀 Starting NSE NFO data fetch...")
    session = get_nse_session()

    print("📦 Fetching F&O stocks...")
    df_stocks = fetch_fno_stocks(session)
    print(f"   → {len(df_stocks)} stocks found")

    print("🏛️  Fetching index data...")
    df_indices = fetch_index_data(session)
    print(f"   → {len(df_indices)} indices found")

    output_path = os.path.abspath(OUTPUT_PATH)

    if os.path.exists(output_path):
        print("🔄 Updating existing Excel file...")
        update_change_log(output_path, df_stocks, df_indices)
        # Rebuild the data sheets with fresh data
        wb = load_workbook(output_path)
        for sname in ["📈 F&O Stocks", "🏛️ Indices", "📊 Summary"]:
            if sname in wb.sheetnames:
                del wb[sname]
        wb.save(output_path)

    print("📊 Building Excel report...")
    build_excel(df_stocks, df_indices, output_path)

    # Save a JSON snapshot alongside
    snapshot = {
        "updated_at": datetime.now().isoformat(),
        "total_stocks": len(df_stocks),
        "total_indices": len(df_indices),
        "stocks": df_stocks.to_dict(orient="records"),
        "indices": df_indices.to_dict(orient="records"),
    }
    json_path = output_path.replace(".xlsx", ".json")
    with open(json_path, "w") as f:
        json.dump(snapshot, f, indent=2, default=str)
    print(f"✅ JSON snapshot saved → {json_path}")
    print("🎉 Done!")


if __name__ == "__main__":
    main()
