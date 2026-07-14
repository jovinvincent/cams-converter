"""
CAMS CAS → Portfolio Excel Converter
One-click web app: upload PDF, get Excel automatically.
Free hosting on Streamlit Community Cloud.
"""

import io
import re
import pdfplumber
import streamlit as st
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ============================================================================
# CONFIG
# ============================================================================

QUANT_SIF_ISINS = {
    "INF966L30019": "Equity Long-Short Fund",
    "INF966L30159": "Equity Ex-Top 100 Long-Short",
    "INF966L30274": "Sector Rotation LongShort Fu",
    "INF966L30241": "Active Asset Allocator Long-",
    "INF966L30076": "Hybrid Long-Short Fund",
}
LIQUID_FUND_ISIN = "INF966L01820"

LATEST_NAV = {
    "INF966L30019": 10.9871,
    "INF966L30159": 10.7319,
    "INF966L30274": 10.3595,
    "INF966L30241": 10.7997,
    "INF966L30076": 10.7412,
}

STAMP_DUTY_RATE = 0.00005
STT_RATE = 0.00001
XIRR_VALUATION_DATE = datetime(2026, 7, 10)


# ============================================================================
# PDF PARSING
# ============================================================================

def extract_pdf_text(pdf_bytes: bytes, password: str = "") -> str:
    with pdfplumber.open(io.BytesIO(pdf_bytes), password=password or None) as pdf:
        return "\n".join(p.extract_text() or "" for p in pdf.pages)


def parse_num(s: str) -> float:
    s = s.replace(",", "").replace("(", "-").replace(")", "").strip()
    try:
        return float(s)
    except Exception:
        return 0.0


def parse_date(s: str):
    for fmt in ("%d-%b-%Y", "%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(s.strip(), fmt)
        except Exception:
            continue
    return None


def parse_cas_pdf(text: str) -> dict:
    funds = {isin: [] for isin in QUANT_SIF_ISINS}
    lines = text.split("\n")
    current_isin = None

    txn_re = re.compile(
        r"^(\d{2}-[A-Za-z]{3}-\d{4})\s+(.+?)\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{3})\s+([\d,]+\.\d{4})\s+([\d,]+\.\d{3})\s*$"
    )
    isin_re = re.compile(r"ISIN:\s*(INF966L\d{5})")

    for line in lines:
        m_isin = isin_re.search(line)
        if m_isin:
            isin = m_isin.group(1)
            if isin in QUANT_SIF_ISINS:
                current_isin = isin
            elif isin == LIQUID_FUND_ISIN:
                current_isin = None
            continue

        if not current_isin:
            continue

        m_txn = txn_re.match(line.strip())
        if not m_txn:
            continue

        date_s, desc, amount_s, units_s, nav_s, balance_s = m_txn.groups()
        d = parse_date(date_s)
        if not d:
            continue
        amount = parse_num(amount_s)
        units = parse_num(units_s)
        nav = parse_num(nav_s)
        balance = parse_num(balance_s)

        if amount > 0:
            txn_type = "NFO Purchase" if "NFO" in desc else "New Purchase"
            if "Additional" in desc:
                txn_type = "Additional Purchase"
            if "Lateral" in desc and "In" in desc:
                txn_type = "Lateral Shift In"
            stamp = round(amount * STAMP_DUTY_RATE, 2)
            stt = 0.0
        else:
            txn_type = "Redemption, STT" if "STT" in desc else "Redemption"
            if "Lateral" in desc and "Out" in desc:
                txn_type = "Lateral Shift Out, STT"
            stamp = 0.0
            stt = round(abs(amount) * STT_RATE, 2) if "STT" in desc else 0.0

        funds[current_isin].append({
            "date": d.strftime("%d-%b-%Y"),
            "fund": f"Quant SIF - {QUANT_SIF_ISINS[current_isin]} - Direct Plan",
            "isin": current_isin,
            "type": txn_type,
            "amount": amount,
            "stt": stt,
            "stamp_duty": stamp,
            "tds": 0.0,
            "units": units,
            "nav": nav,
            "balance": balance,
            "notes": "Purchase" if amount > 0 else "Redemption net of TDS/STT",
        })
    return funds


# ============================================================================
# EXCEL WRITER (locked-in format)
# ============================================================================

def build_excel(funds: dict) -> bytes:
    wb = Workbook()
    wb.remove(wb.active)

    title_font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    title_fill = PatternFill("solid", fgColor="2E75B6")
    header_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="1F4E79")
    summary_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    summary_fill = PatternFill("solid", fgColor="2E75B6")
    purchase_font = Font(name="Calibri", size=10, color="006100")
    redeem_font = Font(name="Calibri", size=10, color="9C0006")
    thin = Side(border_style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    headers = [
        "#", "Date", "Fund", "ISIN", "Type", "Amount (₹)", "STT 0.001% (₹)",
        "Stamp Duty (₹)", "TDS (₹)", "Total Invested (₹)", "Units", "NAV (₹)",
        "Value (₹)", "Balance Units", "Total Investor Cost (₹)", "Cost/Unit (₹)",
        "Holding Days", "Period", "Exit Load (₹)", "Notes",
    ]

    for isin, sheet_name in QUANT_SIF_ISINS.items():
        ws = wb.create_sheet(sheet_name)
        txns = funds.get(isin, [])

        # Row 1: Title
        ws.merge_cells("A1:T1")
        c = ws.cell(row=1, column=1, value=QUANT_SIF_ISINS[isin])
        c.font = title_font
        c.fill = title_fill
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[1].height = 28

        # Row 3: Headers
        for col, h in enumerate(headers, start=1):
            c = ws.cell(row=3, column=col, value=h)
            c.font = header_font
            c.fill = header_fill
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = border
        ws.row_dimensions[3].height = 32

        # Rows 4+: Transactions
        lots = []
        for i, t in enumerate(txns):
            r = 4 + i
            units = t["units"]
            if units > 0:
                stamp = t["stamp_duty"]
                cpu = (t["amount"] + stamp) / units
                lots.append({"units": units, "cpu": cpu, "date": t["date"]})
            elif units < 0:
                to_sell = abs(units)
                new_lots = []
                for lot in lots:
                    if to_sell <= 0:
                        new_lots.append(lot)
                        continue
                    take = min(lot["units"], to_sell)
                    if lot["units"] - take > 1e-9:
                        new_lots.append({"units": lot["units"] - take, "cpu": lot["cpu"], "date": lot["date"]})
                    to_sell -= take
                lots = new_lots

            balance = sum(l["units"] for l in lots)
            total_cost = sum(l["units"] * l["cpu"] for l in lots)

            row_vals = [
                i + 1, t["date"], t["fund"], t["isin"], t["type"],
                t["amount"], t["stt"], t["stamp_duty"], t["tds"],
                round(t["units"] * t["nav"], 3) if t["nav"] else None,
                t["units"], t["nav"], None,
                round(balance, 3), round(total_cost, 2),
                None, None, None, t["notes"],
            ]
            for col, v in enumerate(row_vals, start=1):
                c = ws.cell(row=r, column=col, value=v)
                c.font = purchase_font if (t["amount"] or 0) > 0 else redeem_font
                c.alignment = Alignment(horizontal="center" if col in (1, 4, 6, 7, 8, 9, 10, 11, 12, 14, 15, 16, 17, 18, 19) else "left", vertical="center")
                c.border = border
            ws.row_dimensions[r].height = 18

        # SUMMARY block (last 4 rows)
        if txns:
            nav = LATEST_NAV.get(isin, 0)
            final_balance = sum(l["units"] for l in lots)
            final_cost = sum(l["units"] * l["cpu"] for l in lots)

            total_lot_units = sum(l["units"] for l in lots)
            if total_lot_units > 0:
                wd = sum(
                    l["units"] * (XIRR_VALUATION_DATE - parse_date(l["date"])).days
                    for l in lots
                ) / total_lot_units
            else:
                wd = 0
            xirr = ((nav * final_balance / final_cost) ** (365 / wd) - 1) * 100 if (wd > 0 and final_cost > 0) else 0

            n_txn = len(txns)
            sr_summary = 4 + n_txn + 1
            sr_cost = sr_summary + 1
            sr_bal = sr_summary + 2
            sr_xirr = sr_summary + 3

            for c in range(1, 21):
                ws.cell(row=sr_summary, column=c).fill = summary_fill
                ws.cell(row=sr_cost, column=c).fill = summary_fill
                ws.cell(row=sr_bal, column=c).fill = summary_fill
                ws.cell(row=sr_xirr, column=c).fill = summary_fill

            ws.cell(row=sr_summary, column=1, value="  SUMMARY").font = summary_font
            ws.cell(row=sr_summary, column=1).alignment = Alignment(horizontal="left", vertical="center", indent=1)

            ws.cell(row=sr_cost, column=1, value="  Total Investor Cost (₹)").font = summary_font
            ws.cell(row=sr_cost, column=1).alignment = Alignment(horizontal="left", vertical="center", indent=1)
            cc = ws.cell(row=sr_cost, column=5, value=round(final_cost, 2))
            cc.font = summary_font
            cc.alignment = Alignment(horizontal="right", vertical="center", indent=1)
            cc.number_format = '#,##0.00'

            ws.cell(row=sr_bal, column=1, value="  Balance Units").font = summary_font
            ws.cell(row=sr_bal, column=1).alignment = Alignment(horizontal="left", vertical="center", indent=1)
            cc = ws.cell(row=sr_bal, column=5, value=round(final_balance, 3))
            cc.font = summary_font
            cc.alignment = Alignment(horizontal="right", vertical="center", indent=1)
            cc.number_format = '#,##0.000'

            ws.cell(row=sr_xirr, column=1, value="  XIRR (Annualized)").font = summary_font
            ws.cell(row=sr_xirr, column=1).alignment = Alignment(horizontal="left", vertical="center", indent=1)
            cc = ws.cell(row=sr_xirr, column=5, value=round(xirr, 2))
            cc.font = summary_font
            cc.alignment = Alignment(horizontal="right", vertical="center", indent=1)
            cc.number_format = '0.00"%"'

            for r in (sr_summary, sr_cost, sr_bal, sr_xirr):
                ws.row_dimensions[r].height = 22

        widths = [5, 13, 35, 14, 22, 13, 11, 12, 9, 14, 11, 10, 11, 13, 16, 11, 11, 9, 11, 24]
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[chr(64 + i)].width = w
        ws.freeze_panes = "A4"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ============================================================================
# STREAMLIT UI - One-click
# ============================================================================

st.set_page_config(
    page_title="CAMS → Excel",
    page_icon="📊",
    layout="centered",
)

st.title("📊 CAMS CAS → Portfolio Excel")
st.write("Upload your CAMS PDF. Excel downloads automatically.")

uploaded = st.file_uploader(
    "Choose CAMS PDF",
    type=["pdf"],
    label_visibility="visible",
)

if uploaded is not None:
    with st.spinner("Processing PDF..."):
        try:
            pdf_bytes = uploaded.read()
            text = extract_pdf_text(pdf_bytes, "")
            funds = parse_cas_pdf(text)
            xlsx_bytes = build_excel(funds)
        except Exception as e:
            st.error(f"❌ Failed: {e}")
            st.info("If PDF is password-protected, edit app.py to set DEFAULT_PASSWORD.")
            st.stop()

    total_txn = sum(len(v) for v in funds.values())
    if total_txn == 0:
        st.warning("No transactions found. Check password or PDF format.")
        st.stop()

    st.success(f"✅ {total_txn} transactions parsed")
    st.download_button(
        label="📥 Download Excel",
        data=xlsx_bytes,
        file_name="Portfolio_Transactions.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
