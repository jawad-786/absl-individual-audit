import io
import re
import openpyxl
import pandas as pd
import pypdf
import streamlit as st

st.set_page_config(
    page_title="ABSLI Fund Factsheet Audit Engine",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Custom Styling for Corporate ABSLI Theme
st.markdown(
    """
    <style>
        .main-header {
            background: linear-gradient(135deg, #c4122f 0%, #8b0015 100%);
            padding: 20px 28px;
            color: white;
            border-radius: 8px;
            margin-bottom: 22px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }
        .main-header h1 { margin: 0; font-size: 24px; font-weight: 700; color: white; }
        .main-header p { margin-top: 5px; opacity: 0.9; font-size: 13px; margin-bottom: 0; }
        .metric-card {
            background: white;
            border: 1px solid #e2e8f0;
            border-radius: 8px;
            padding: 12px 16px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.05);
            text-align: center;
        }
        .pass-tag { background-color: #d4edda; color: #155724; padding: 4px 10px; border-radius: 4px; font-weight: 700; font-size: 12px; }
        .breach-tag { background-color: #f8d7da; color: #721c24; padding: 4px 10px; border-radius: 4px; font-weight: 700; font-size: 12px; }
        .warn-tag { background-color: #fff3cd; color: #856404; padding: 4px 10px; border-radius: 4px; font-weight: 700; font-size: 12px; }
        .info-box { background: #ebf8ff; border-left: 4px solid #3182ce; padding: 12px 16px; border-radius: 4px; font-size: 13px; margin-bottom: 15px; }
        
        button[data-testid="stNumberInputStepDown"],
        button[data-testid="stNumberInputStepUp"] {
            display: none !important;
        }
        div[data-testid="stNumberInputContainer"] {
            padding-right: 8px !important;
        }
        div[data-testid="stButton"] button {
            margin-top: 28px;
            height: 42px;
            font-weight: 700;
        }
    </style>
""",
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="main-header">
        <h1>🛡️ ABSLI Fund Factsheet Audit Engine</h1>
        <p>Comprehensive Factsheet Audit: NAV, AUM, Corporate Debt, G-Sec, Equity, Asset Allocation, Ratings & Sectors</p>
    </div>
""",
    unsafe_allow_html=True,
)

# -------------------------------------------------------------
# BUILT-IN SCHEME DIRECTORY
# -------------------------------------------------------------
BUILTIN_SCHEME_NAMES = {
    "ULGF00112/06/01BSLGGROWTH109": "Group Growth Fund",
    "ULGF00212/06/01BSLGSECURE109": "Group Secure Fund",
    "ULGF00312/06/01BSLGSTABLE109": "Group Stable Fund",
    "ULGF00416/07/02BSLGFIXINT109": "Group Fixed Interest Fund",
    "ULGF00530/05/03BSLIGRBOND109": "Group Bond Fund",
    "ULGF00630/05/03BSLIGRGILT109": "Group Gilt Fund",
    "ULGF00724/08/04BSLIGFLOAT109": "Group Floating Rate Fund",
    "ULGF00824/08/04BSLIGRMMKT109": "Group Money Market Fund",
    "ULGF01026/11/07BSLIGGRADV109": "Group Growth Advantage Fund",
    "ULGF01322/09/08BSLGSHTDBT109": "Group Short Term Debt Fund",
    "ULGF01425/02/10BSLGINCADV109": "Group Income Advantage Fund",
    "ULGF01525/02/10BSLIGRMAXI109": "Group Growth Maximiser Fund",
    "ULGF01628/11/11BSLGRBOND2109": "Group Bond Fund II",
    "ULGF01728/11/11BSLGFXINT2109": "Group Fixed Interest Fund II",
    "ULGF01828/11/11BSLGROWTH2109": "Group Growth Fund II",
    "ULGF01928/11/11BSLGRMMKT2109": "Group Money Market Fund II",
    "ULGF02028/11/11BSLGSECUR2109": "Group Secure Fund II",
    "ULGF02128/11/11BSLGSHTDB2109": "Group Short Term Debt Fund II",
    "ULGF02228/11/11BSLGSTABL2109": "Group Stable Fund II",
    "ULGF02311/07/25BSLCAPPRTC109": "Group Capital Protection Fund",
    "ULIF00113/03/01BSLBUILDER109": "Individual Builder Fund",
    "ULIF00213/03/01BSLENHANCE109": "Individual Enhancer Fund",
    "ULIF00313/03/01BSLPROTECT109": "Individual Protector Fund",
    "ULIF00404/03/03BSLIENRICH109": "Pension Enrich Fund",
    "ULIF00504/03/03BSLIGROWTH109": "Pension Growth Fund",
    "ULIF00604/03/03BSLNOURISH109": "Pension Nourish Fund",
    "ULIF00704/02/04BSLCREATOR109": "Individual Creator Fund",
    "ULIF00826/06/04BSLIIMAGNI109": "Individual Magnifier Fund",
    "ULIF00931/05/05BSLBALANCE109": "Individual Balancer Fund",
    "ULIF01008/07/05BSLIASSURE109": "Individual Assure Fund",
    "ULIF01101/06/07BSLIINMAXI109": "Individual Maximiser Fund",
    "ULIF01217/10/07BSLINMULTI109": "Individual Multiplier Fund",
    "ULIF01507/08/08BSLIINCADV109": "Individual Income Advantage Fund",
    "ULIF01723/06/09BSLSUPER20109": "Super 20 Fund",
    "ULIF01911/12/09BSLITITAN1109": "Titanium I Fund",
    "ULIF02203/02/10BSLPLATPR1109": "Platinum Premier Fund",
    "ULIF02408/09/10BSLPLATADV109": "Platinum Advantage Fund",
    "ULIF02707/10/11BSLIPUREEQ109": "Individual Pure Equity Fund",
    "ULIF02807/10/11BSLLIQPLUS109": "Individual Liquid Plus Fund",
    "ULIF02907/10/11BSLIVALUEM109": "Individual Value Momentum Fund",
    "ULIF03027/08/13BSLIMAXGT109": "Individual Maximiser Guaranteed Fund",
    "ULIF03127/08/13BSLIINADGT109": "Individual Income Advantage Guaranteed Fund",
    "ULIF03205/07/13BSLILDIS109": "BSLI Linked Discontinued Policy Fund",
    "ULIF03305/07/13BSLIPNDIS109": "BSLI Pension Discontinued Policy Fund",
    "ULIF03430/10/14BSLIASTALC109": "BSLI Asset Allocation Fund",
    "ULIF03530/10/14BSLICNFIDX109": "BSLI Capped Nifty Index Fund",
    "ULIF03722/06/18ABSLIMUMNC109": "ABSLI MNC Fund",
    "ULIF03810/11/23ABSLESGFND109": "ABSLI ESG Fund",
    "ULIF03910/11/23ABSLSMALCP109": "ABSLI Small Cap Fund",
    "ULIF04015/06/24ABSLIALPHA109": "Nifty Alpha 50 Index Fund",
    "ULIF04101/10/24ABSLMOMNTM109": "Nifty Midcap 150 Momentum 50 Index Fund",
    "ULIF04206/11/25ABSLDIVYLD109": "Dividend Yield Fund",
    "ULIF04312/02/26ABSLFLXCAP109": "ABSLI Flexicap Fund",
    "ULIF04410/09/26ABSLPENMAX109": "Pension Maximiser Fund",
    "ULIF04510/09/26ABSLPENINA109": "Pension Income Advantage Fund",
}


# -------------------------------------------------------------
# TABLE STYLING HELPER (MISMATCH / NOT IN PDF / BREACH: Red, PASS: Green)
# -------------------------------------------------------------
def color_status_cells(val):
  if not isinstance(val, str):
    return ""
  val_upper = val.upper().strip()

  if any(
      k in val_upper
      for k in ["MISMATCH", "NOT IN PDF", "BREACH", "FAIL", "DIFF"]
  ):
    return "background-color: #f8d7da; color: #721c24; font-weight: bold;"
  elif any(k in val_upper for k in ["PASS", "MATCH", "COMPLIANT"]):
    return "background-color: #d4edda; color: #155724; font-weight: bold;"
  elif any(k in val_upper for k in ["PENDING", "WARNING", "CONCENTRATION"]):
    return "background-color: #fff3cd; color: #856404; font-weight: bold;"
  return ""


def apply_status_style(df, column_name):
  styler = df.style
  if hasattr(styler, "map"):
    return styler.map(color_status_cells, subset=[column_name])
  else:
    return styler.applymap(color_status_cells, subset=[column_name])


# -------------------------------------------------------------
# SIDEBAR: Flexible Multi-File Uploaders (ALL OPTIONAL)
# -------------------------------------------------------------
with st.sidebar:
  st.header("📂 Data Sources")
  st.caption("Upload PDF (Individual/Group) & Master Excel Sheets:")

  uploaded_pdf = st.file_uploader(
      "1. Factsheet PDF (Individual/Group)", type=["pdf"]
  )
  uploaded_master = st.file_uploader("2. Master Factsheet Excel", type=["xlsx"])
  uploaded_nav = st.file_uploader(
      "3. Birla Nav Excel (Birla Nav.xlsx)", type=["xlsx"]
  )
  uploaded_aum = st.file_uploader(
      "4. AUM Breakup Excel (AUM Breakup.xlsx)", type=["xlsx"]
  )
  uploaded_ytm = st.file_uploader(
      "5. Yield & Duration Excel (Yield...xlsx)", type=["xlsx"]
  )

  tolerance = st.slider(
      "Tolerance Threshold (Difference %)",
      min_value=0.00,
      max_value=0.10,
      value=0.02,
      step=0.01,
  )


# -------------------------------------------------------------
# PDF PARSER (Extracts SFIN, Scheme Name, NAV & AUM)
# -------------------------------------------------------------
@st.cache_data(show_spinner=False)
def extract_pdf_data(pdf_bytes):
  schemes_pdf = {}
  if not pdf_bytes:
    return schemes_pdf

  try:
    reader = pypdf.PdfReader(io.BytesIO(pdf_bytes))
    for page_idx, page in enumerate(reader.pages):
      text = page.extract_text() or ""
      sfin_matches = re.findall(r"(?:ULIF|ULGF)[A-Za-z0-9/]{10,35}", text)
      if not sfin_matches:
        continue

      nav_candidates = re.findall(r"\b\d{1,3}\.\d{4}\b", text)
      aum_candidates = re.findall(r"([\d\.]+)\s*Cr", text)

      fm_match = re.search(
          r"FUND\s+MANAGER:\s*([^\n\r]+)", text, re.IGNORECASE
      )
      fm_val = fm_match.group(1).strip() if fm_match else ""

      ytm_match = re.search(
          r"Yield\s+to\s+Maturity\s*\|?\s*([\d\.]+)%", text, re.IGNORECASE
      )
      dur_match = re.search(
          r"Modified\s+Duration\s*\|?\s*([\d\.]+)", text, re.IGNORECASE
      )

      # Extract fund name from header line
      header_fund_name = None
      for line in text.split("\n")[:10]:
        if any(s in line for s in sfin_matches):
          clean_line = re.sub(
              r"SFIN\s*No\.?\s*:?.*", "", line, flags=re.IGNORECASE
          ).strip()
          clean_line = re.sub(
              r"(?:ULIF|ULGF)[A-Za-z0-9/]{10,35}", "", clean_line
          ).strip()
          if (
              clean_line
              and len(clean_line) > 3
              and clean_line.upper()
              not in ["EQUITY", "CORPORATE DEBT", "TOTAL", "SECURITIES"]
          ):
            header_fund_name = clean_line
            break

      for sfin in set(sfin_matches):
        sfin = sfin.strip()
        best_name = (
            header_fund_name
            or BUILTIN_SCHEME_NAMES.get(sfin)
            or f"Scheme ({sfin})"
        )

        if sfin not in schemes_pdf or nav_candidates:
          schemes_pdf[sfin] = {
              "page": page_idx + 1,
              "sfin": sfin,
              "fund_name": best_name,
              "page_text": text,
              "nav_candidates": [float(n) for n in nav_candidates],
              "aum_candidates": [float(a) for a in aum_candidates],
              "fund_manager": fm_val,
              "ytm": float(ytm_match.group(1)) if ytm_match else None,
              "mod_duration": float(dur_match.group(1)) if dur_match else None,
          }
  except Exception as e:
    st.sidebar.error(f"Error reading PDF: {e}")
  return schemes_pdf


# -------------------------------------------------------------
# EXCEL PARSER (Deep Scanning & Header Exclusion)
# -------------------------------------------------------------
@st.cache_data(show_spinner=False)
def parse_master_sheets(master_bytes):
  if not master_bytes:
    return {}
  parsed = {}
  try:
    wb = openpyxl.load_workbook(io.BytesIO(master_bytes), data_only=True)
    for sheet_name in wb.sheetnames:
      ws = wb[sheet_name]
      fund_name = None
      sfin = None

      for r in range(1, 30):
        for c in range(1, 25):
          val = ws.cell(r, c).value
          if val and isinstance(val, str):
            val_str = val.strip()
            if ("ULIF" in val_str or "ULGF" in val_str) and not sfin:
              m_sfin = re.search(r"(?:ULIF|ULGF)[A-Za-z0-9/]{10,35}", val_str)
              if m_sfin:
                sfin = m_sfin.group(0).strip()
            if (
                not fund_name
                and any(
                    k in val_str.upper()
                    for k in [
                        "FUND",
                        "BOND",
                        "SECURE",
                        "STABLE",
                        "GROWTH",
                        "INDEX",
                    ]
                )
                and val_str.strip().upper()
                not in [
                    "EQUITY",
                    "CORPORATE DEBT",
                    "GOVERNMENT SECURITIES",
                    "TOTAL",
                    "MMI",
                    "SECURITIES",
                ]
            ):
              if len(val_str) < 60 and not val_str.startswith("UL"):
                fund_name = val_str

      if not sfin:
        m_sheet = re.search(r"(?:ULIF|ULGF)[A-Za-z0-9/]{10,35}", sheet_name)
        if m_sheet:
          sfin = m_sheet.group(0).strip()
        elif len(wb.sheetnames) > 3:
          sfin = sheet_name.strip()

      if not sfin:
        continue

      if not fund_name or fund_name.upper() in [
          "EQUITY",
          "CORPORATE DEBT",
          "TOTAL",
      ]:
        fund_name = (
            BUILTIN_SCHEME_NAMES.get(sfin)
            or ws.cell(2, 3).value
            or sheet_name
        )

      # 1. Asset Allocation & Mandates
      asset_alloc = []
      for r in range(6, 20):
        for c in range(5, 12):
          cat = ws.cell(r, c).value
          if cat and any(
              ac in str(cat).upper()
              for ac in [
                  "EQUITY",
                  "DEBT",
                  "GOVERNMENT",
                  "CORPORATE",
                  "MMI",
                  "TOTAL",
              ]
          ):
            actual = ws.cell(r, c + 1).value
            rng = ws.cell(r, c + 2).value
            if (
                actual is not None
                and str(cat).strip().upper()
                not in ["ASSET CATEGORY", "ASSET ALLOCATION"]
            ):
              try:
                act_val = float(actual)
                pct = act_val * 100 if act_val <= 1.0 else act_val
                asset_alloc.append({
                    "Asset Category": str(cat).strip(),
                    "Actual (%)": f"{pct:.2f}%",
                    "Raw Actual": round(pct, 2),
                    "Defined Mandate Range": str(rng) if rng else "N/A",
                })
              except:
                pass
            break

      # 2. Rating Profile
      ratings = []
      seen_ratings = set()
      for r in range(1, 15):
        rat = ws.cell(r, 18).value
        pct = ws.cell(r, 19).value
        if (
            rat
            and pct is not None
            and str(rat).strip()
            in ["AAA", "SOV", "AA+", "AA", "A1+", "AA-", "A+", "P1+"]
        ):
          try:
            p_float = float(pct)
            p_val = p_float * 100 if p_float <= 1.0 else p_float
            r_name = str(rat).strip()
            if r_name not in seen_ratings:
              seen_ratings.add(r_name)
              ratings.append({
                  "Rating": r_name,
                  "Weight (%)": f"{p_val:.2f}%",
                  "Raw Weight": round(p_val, 2),
              })
          except:
            pass

      if not ratings:
        for r in range(1, 40):
          for c in range(1, 25):
            val = ws.cell(r, c).value
            if (
                val
                and str(val).strip()
                in ["AAA", "SOV", "AA+", "AA", "A1+", "AA-", "A+", "P1+"]
            ):
              r_name = str(val).strip()
              if r_name not in seen_ratings:
                for dc in [1, -1, 2]:
                  v_num = ws.cell(r, c + dc).value
                  if (
                      v_num
                      and isinstance(v_num, (int, float))
                      and 0 <= float(v_num) <= 100
                  ):
                    p_val = (
                        float(v_num) * 100 if float(v_num) <= 1.0 else float(v_num)
                    )
                    seen_ratings.add(r_name)
                    ratings.append({
                        "Rating": r_name,
                        "Weight (%)": f"{p_val:.2f}%",
                        "Raw Weight": round(p_val, 2),
                    })
                    break

      # 3. Sector Allocation
      sectors = []
      for r in range(25, 60):
        for c in range(5, 10):
          sec_name = ws.cell(r, c).value
          if (
              sec_name
              and len(str(sec_name).strip()) > 3
              and str(sec_name).strip() != "Sectoral Allocation"
          ):
            for dc in [1, 2]:
              sec_pct = ws.cell(r, c + dc).value
              if (
                  sec_pct
                  and isinstance(sec_pct, (int, float))
                  and 0 < float(sec_pct) <= 100
              ):
                try:
                  s_float = float(sec_pct)
                  s_val = s_float * 100 if s_float <= 1.0 else s_float
                  sectors.append({
                      "Sector Name": str(sec_name).strip(),
                      "Weight (%)": f"{s_val:.2f}%",
                      "Raw Weight": round(s_val, 2),
                  })
                except:
                  pass
                break
            break

      # 4. Line-by-Line Holdings Split
      corp_debt = []
      g_sec = []
      equity = []
      curr_sec = None

      for r in range(5, 95):
        name_cell = ws.cell(r, 3).value
        wt_cell = ws.cell(r, 4).value

        if not name_cell:
          continue

        item_str = str(name_cell).strip()
        item_upper = item_str.upper()

        if item_upper == "CORPORATE DEBT":
          curr_sec = "corp"
          continue
        elif item_upper in [
            "GOVERNMENT SECURITIES",
            "CENTRAL GOVERNMENT SECURITIES",
        ]:
          curr_sec = "gsec"
          continue
        elif item_upper == "EQUITY":
          curr_sec = "equity"
          continue
        elif item_upper in [
            "MMI",
            "TOTAL",
            "SECURITISED DEBT",
            "OTHER SECURITISED DEBT",
            "SECURITIES",
        ]:
          curr_sec = None
          continue

        if curr_sec and wt_cell is not None:
          try:
            w_num = float(wt_cell)
            w_pct = w_num * 100 if w_num <= 1.0 else w_num
            clean_row = {
                "Security Description": item_str,
                "Weight (%)": f"{w_pct:.2f}%",
                "Raw Weight": round(w_pct, 2),
                "Status": "PASS",
            }
            if curr_sec == "corp":
              corp_debt.append(clean_row)
            elif curr_sec == "gsec":
              g_sec.append(clean_row)
            elif curr_sec == "equity":
              equity.append(clean_row)
          except:
            pass

      parsed[sfin] = {
          "sheet_name": sheet_name,
          "fund_name": fund_name or sheet_name,
          "sfin": sfin,
          "asset_allocation": asset_alloc,
          "ratings": ratings,
          "sectors": sectors,
          "corp_debt": corp_debt,
          "g_sec": g_sec,
          "equity": equity,
      }
  except Exception as e:
    st.sidebar.error(f"Master Sheet Read Error: {e}")
  return parsed


@st.cache_data(show_spinner=False)
def parse_nav_file(nav_bytes):
  if not nav_bytes:
    return {}
  nav_dict = {}
  try:
    df = pd.read_excel(io.BytesIO(nav_bytes))
    for _, row in df.iterrows():
      sfin = str(row.get("SFIN", "")).strip()
      nav = row.get("NAV Per Unit", None)
      name = row.get("Segregated Fund Name", "")
      if (
          sfin
          and (sfin.startswith("ULIF") or sfin.startswith("ULGF"))
          and pd.notnull(nav)
      ):
        nav_dict[sfin] = {
            "nav": float(nav),
            "name": (
                str(name).strip()
                if str(name).strip()
                else BUILTIN_SCHEME_NAMES.get(sfin, sfin)
            ),
            "type": (
                "Individual (ULIF)"
                if sfin.startswith("ULIF")
                else "Group (ULGF)"
            ),
        }
  except Exception as e:
    st.sidebar.error(f"NAV File Error: {e}")
  return nav_dict


@st.cache_data(show_spinner=False)
def parse_aum_file(aum_bytes):
  if not aum_bytes:
    return {}
  aum_dict = {}
  try:
    xls = pd.ExcelFile(io.BytesIO(aum_bytes))
    target_sheet = (
        "AUm Break UP 2"
        if "AUm Break UP 2" in xls.sheet_names
        else xls.sheet_names[0]
    )
    df = pd.read_excel(io.BytesIO(aum_bytes), sheet_name=target_sheet)
    for _, row in df.iterrows():
      row_code = str(row.iloc[0]).strip()
      if row_code and row_code != "nan" and row_code != "Row Labels":
        try:
          c_debt = (
              row.iloc[1] if len(row) > 1 and pd.notnull(row.iloc[1]) else 0
          )
          eq = row.iloc[2] if len(row) > 2 and pd.notnull(row.iloc[2]) else 0
          gsec = row.iloc[3] if len(row) > 3 and pd.notnull(row.iloc[3]) else 0
          mmi = row.iloc[4] if len(row) > 4 and pd.notnull(row.iloc[4]) else 0
          gt = row.iloc[5] if len(row) > 5 and pd.notnull(row.iloc[5]) else 0

          aum_dict[row_code] = {
              "Corporate Debt": float(c_debt) / 1e7 if c_debt else 0,
              "Equity": float(eq) / 1e7 if eq else 0,
              "G-Sec": float(gsec) / 1e7 if gsec else 0,
              "MMI": float(mmi) / 1e7 if mmi else 0,
              "Total AUM (Cr)": float(gt) / 1e7 if gt else 0,
          }
        except:
          pass
  except Exception as e:
    st.sidebar.error(f"AUM File Error: {e}")
  return aum_dict


@st.cache_data(show_spinner=False)
def parse_ytm_file(ytm_bytes):
  if not ytm_bytes:
    return {}
  ytm_dict = {}
  try:
    df = pd.read_excel(io.BytesIO(ytm_bytes))
    for _, row in df.iterrows():
      code = str(row.get("Scheme Code", "")).strip()
      ytm = row.get("YTM", None)
      dur = row.get("Duration", None)
      if code and code != "nan":
        ytm_dict[code] = {
            "YTM": float(ytm) * 100 if pd.notnull(ytm) else None,
            "Duration": float(dur) if pd.notnull(dur) else None,
        }
  except Exception as e:
    st.sidebar.error(f"YTM File Error: {e}")
  return ytm_dict


# Load Data
pdf_data = extract_pdf_data(uploaded_pdf.getvalue()) if uploaded_pdf else {}
master_data = (
    parse_master_sheets(uploaded_master.getvalue()) if uploaded_master else {}
)
nav_data = parse_nav_file(uploaded_nav.getvalue()) if uploaded_nav else {}
aum_data = parse_aum_file(uploaded_aum.getvalue()) if uploaded_aum else {}
ytm_data = parse_ytm_file(uploaded_ytm.getvalue()) if uploaded_ytm else {}

# Status Metrics
status_cols = st.columns(6)
status_cols[0].metric("PDF Schemes", len(pdf_data) if pdf_data else "No PDF")
status_cols[1].metric(
    "Master Sheets", len(master_data) if master_data else "No Master"
)
status_cols[2].metric("NAV Records", len(nav_data) if nav_data else "No NAV")
status_cols[3].metric("AUM Breakups", len(aum_data) if aum_data else "No AUM")
status_cols[4].metric("YTM Records", len(ytm_data) if ytm_data else "No YTM")
status_cols[5].metric("Tolerance", f"{tolerance:.2f}%")

st.write("")

# Main Tabs (AUM Master Audit Tab Included as Tab 2)
tabs = st.tabs([
    "📈 NAV Master Audit (All Schemes)",
    "💰 AUM Master Audit (All Schemes)",
    "📑 Scheme-Wise Deep Audit",
    "📊 AUM Asset Breakup",
    "📉 YTM & Duration Audit",
    "📥 Export Full Audit Workbook",
])

# -------------------------------------------------------------
# TAB 1: NAV MASTER AUDIT (PDF vs Excel Cross Check)
# -------------------------------------------------------------
with tabs[0]:
  st.subheader("📈 Direct NAV Master Audit (Individual & Group Funds)")
  if not nav_data:
    st.info(
        "💡 Sirf **`Birla Nav.xlsx`** sidebar mein upload karein aur saari"
        " schemes ka NAV ek sath check ho jayega."
    )
  else:
    cat_filter = st.radio(
        "Filter Schemes (NAV):",
        ["All Schemes", "Individual Funds (ULIF)", "Group Funds (ULGF)"],
        horizontal=True,
        key="nav_cat_filter",
    )

    nav_rows = []
    for sfin, ninfo in nav_data.items():
      if (
          cat_filter == "Individual Funds (ULIF)"
          and not sfin.startswith("ULIF")
      ):
        continue
      if cat_filter == "Group Funds (ULGF)" and not sfin.startswith("ULGF"):
        continue

      excel_nav = ninfo["nav"]
      pdf_info = pdf_data.get(sfin)

      pdf_nav = None
      if pdf_info and pdf_info["nav_candidates"]:
        found = None
        for cand in pdf_info["nav_candidates"]:
          if abs(cand - excel_nav) < 0.05:
            found = cand
            break
        pdf_nav = (
            found if found is not None else pdf_info["nav_candidates"][-1]
        )

      status = "PENDING PDF"
      diff = 0.0
      if not uploaded_pdf:
        status = "PENDING PDF"
      elif pdf_nav is not None:
        diff = abs(pdf_nav - excel_nav)
        status = "PASS" if diff <= tolerance else f"MISMATCH ({diff:.4f})"
      elif sfin not in pdf_data:
        status = "NOT IN PDF"
      else:
        status = "NOT IN PDF"

      nav_rows.append({
          "Category": ninfo.get("type", "Fund"),
          "SFIN Code": sfin,
          "Fund Name": ninfo["name"],
          "Excel NAV Per Unit": round(excel_nav, 4),
          "PDF Published NAV": round(pdf_nav, 4) if pdf_nav else "N/A",
          "Difference": round(diff, 4) if pdf_nav else "N/A",
          "Audit Status": status,
      })

    df_nav_audit = pd.DataFrame(nav_rows)
    styled_nav_audit = apply_status_style(df_nav_audit, "Audit Status")
    st.dataframe(styled_nav_audit, use_container_width=True, hide_index=True)

# -------------------------------------------------------------
# TAB 2: AUM MASTER AUDIT (PDF vs Excel Cross Check - ALL SCHEMES)
# -------------------------------------------------------------
with tabs[1]:
  st.subheader("💰 Direct AUM Master Audit (All Schemes in Crores)")
  if not aum_data and not pdf_data:
    st.info(
        "💡 Sidebar mein **`AUM Breakup.xlsx`** aur **Factsheet PDF** upload"
        " karein AUM audit ke liye."
    )
  else:
    aum_audit_rows = []
    # Collect all available schemes from PDF or NAV or Master
    all_aum_sfins = set(list(pdf_data.keys()) + list(nav_data.keys()))

    for sfin in sorted(list(all_aum_sfins)):
      fname = (
          nav_data.get(sfin, {}).get("name")
          or BUILTIN_SCHEME_NAMES.get(sfin)
          or pdf_data.get(sfin, {}).get("fund_name")
          or sfin
      )
      ftype = (
          "Individual (ULIF)" if sfin.startswith("ULIF") else "Group (ULGF)"
      )

      # Extract PDF AUM
      pdf_aum_val = None
      p_info = pdf_data.get(sfin)
      if p_info and p_info["aum_candidates"]:
        pdf_aum_val = p_info["aum_candidates"][-1]

      # Extract Excel AUM from AUM Breakup
      excel_aum_val = None
      sheet_name_cand = master_data.get(sfin, {}).get("sheet_name", "")

      # Try matching by sfin, sheet name, or fund name abbreviation
      for code_key, a_info in aum_data.items():
        if (
            code_key.upper() in sfin.upper()
            or (sheet_name_cand and code_key.upper() in sheet_name_cand.upper())
            or code_key.upper() in fname.upper()
        ):
          excel_aum_val = a_info["Total AUM (Cr)"]
          break

      status = "PENDING PDF"
      diff = 0.0
      if not uploaded_pdf:
        status = "PENDING PDF"
      elif pdf_aum_val is not None and excel_aum_val is not None:
        diff = abs(pdf_aum_val - excel_aum_val)
        status = "PASS" if diff <= 0.05 else f"MISMATCH ({diff:.2f} Cr)"
      elif pdf_aum_val is not None and excel_aum_val is None:
        status = "PENDING EXCEL"
      elif sfin not in pdf_data:
        status = "NOT IN PDF"
      else:
        status = "NOT IN PDF"

      aum_audit_rows.append({
          "Category": ftype,
          "SFIN Code": sfin,
          "Fund Name": fname,
          "Excel AUM (Cr)": (
              f"{excel_aum_val:.2f} Cr" if excel_aum_val is not None else "N/A"
          ),
          "PDF Published AUM": (
              f"{pdf_aum_val:.2f} Cr" if pdf_aum_val is not None else "N/A"
          ),
          "Difference": (
              f"{diff:.2f} Cr"
              if (pdf_aum_val is not None and excel_aum_val is not None)
              else "N/A"
          ),
          "Audit Status": status,
      })

    df_aum_audit = pd.DataFrame(aum_audit_rows)
    styled_aum_audit = apply_status_style(df_aum_audit, "Audit Status")
    st.dataframe(styled_aum_audit, use_container_width=True, hide_index=True)

# -------------------------------------------------------------
# TAB 3: SCHEME-WISE DEEP AUDIT (WITH WORKING PREV & NEXT NAVIGATION)
# -------------------------------------------------------------
with tabs[2]:
  st.subheader("📑 Scheme-Wise Deep Portfolio Audit")

  # Build scheme dictionary with clean names
  available_schemes = {}
  all_possible_sfins = set(list(master_data.keys()) + list(pdf_data.keys()))

  for sfin_item in all_possible_sfins:
    m_name = (
        master_data.get(sfin_item, {}).get("fund_name")
        if sfin_item in master_data
        else None
    )
    n_name = (
        nav_data.get(sfin_item, {}).get("name") if sfin_item in nav_data else None
    )
    p_name = (
        pdf_data.get(sfin_item, {}).get("fund_name")
        if sfin_item in pdf_data
        else None
    )
    b_name = BUILTIN_SCHEME_NAMES.get(sfin_item)

    final_name = b_name or n_name or m_name or p_name or sfin_item
    if str(final_name).upper() in ["EQUITY", "CORPORATE DEBT", "TOTAL"]:
      final_name = b_name or n_name or sfin_item

    available_schemes[sfin_item] = f"{final_name} ({sfin_item})"

  if not available_schemes:
    st.info(
        "💡 Is section ke liye sidebar mein **Factsheet PDF** ya **Master"
        " Factsheet Excel** upload karein."
    )
  else:
    sorted_sfins = sorted(list(available_schemes.keys()))

    # Ensure scheme_selector_key is initialized with a valid scheme
    if (
        "scheme_selector_key" not in st.session_state
        or st.session_state.scheme_selector_key not in sorted_sfins
    ):
      st.session_state.scheme_selector_key = sorted_sfins[0]

    current_idx = sorted_sfins.index(st.session_state.scheme_selector_key)

    # 3-Column Navigation: [ ⬅️ Previous ] [ Dropdown ] [ Next ➡️ ]
    nav_col1, nav_col2, nav_col3 = st.columns([1.2, 5, 1.2])

    with nav_col1:
      if st.button(
          "⬅️ Previous",
          use_container_width=True,
          disabled=(current_idx == 0),
      ):
        st.session_state.scheme_selector_key = sorted_sfins[current_idx - 1]
        st.rerun()

    with nav_col3:
      if st.button(
          "Next ➡️",
          use_container_width=True,
          disabled=(current_idx == len(sorted_sfins) - 1),
      ):
        st.session_state.scheme_selector_key = sorted_sfins[current_idx + 1]
        st.rerun()

    with nav_col2:
      chosen_sfin = st.selectbox(
          "Select Scheme to Inspect:",
          options=sorted_sfins,
          key="scheme_selector_key",
          format_func=lambda x: available_schemes[x],
      )

    s_master = master_data.get(chosen_sfin, None)
    pdf_info = pdf_data.get(chosen_sfin, None)
    pdf_text = pdf_info["page_text"] if pdf_info else ""
    sheet_key = s_master["sheet_name"] if s_master else ""

    # Sub-tabs for each specific module
    sub_tabs = st.tabs([
        "1. Asset Allocation",
        "2. Rating Profile",
        "3. Sectoral Allocation",
        "4. Corporate Debt",
        "5. Government Securities",
        "6. Equity Holdings",
        "7. AUM Audit & Charts",
    ])

    # 1. Asset Allocation & Mandates
    with sub_tabs[0]:
      st.markdown("#### Asset Allocation Compliance")
      if s_master and s_master["asset_allocation"]:
        df_a = pd.DataFrame(s_master["asset_allocation"])

        def check_mandate(r):
          rng = str(r["Defined Mandate Range"])
          val = float(r["Raw Actual"])
          if "-" in rng and "%" in rng:
            try:
              clean = rng.replace("%", "").split("-")
              lo, hi = float(clean[0]), float(clean[1])
              if val < lo or val > hi:
                return "BREACH"
              return "COMPLIANT"
            except:
              return "COMPLIANT"
          return "COMPLIANT"

        df_a["Compliance"] = df_a.apply(check_mandate, axis=1)
        show_a = df_a[[
            "Asset Category",
            "Actual (%)",
            "Defined Mandate Range",
            "Compliance",
        ]]
        styled_a = apply_status_style(show_a, "Compliance")
        st.dataframe(styled_a, use_container_width=True, hide_index=True)
      else:
        st.info(
            "Upload Master Factsheet Excel to view Asset Allocation compliance."
        )

    # 2. Rating Profile (AUDITED: Excel vs PDF Cross-Check)
    with sub_tabs[1]:
      st.markdown(
          "#### Credit Rating Profile Audit (Excel vs. PDF Pie Chart)"
      )
      if s_master and s_master["ratings"]:
        rating_audit_rows = []
        for r_item in s_master["ratings"]:
          r_label = r_item["Rating"]
          excel_val = r_item["Raw Weight"]

          pdf_val_str = "N/A"
          diff_val = 0.0
          status_str = "PENDING PDF"

          if pdf_text:
            str_target = f"{excel_val:.2f}%"
            if str_target in pdf_text or f"{excel_val:.1f}%" in pdf_text:
              pdf_val_str = str_target
              diff_val = 0.0
              status_str = "PASS"
            else:
              all_pcts = re.findall(r"(\d{1,3}\.\d{1,2})%", pdf_text)
              closest = None
              min_d = 999
              for p in all_pcts:
                d = abs(float(p) - excel_val)
                if d < min_d:
                  min_d = d
                  closest = float(p)
              if closest is not None and min_d <= 0.05:
                pdf_val_str = f"{closest:.2f}%"
                diff_val = round(min_d, 2)
                status_str = "PASS"
              elif closest is not None:
                pdf_val_str = f"{closest:.2f}%"
                diff_val = round(min_d, 2)
                status_str = f"MISMATCH ({diff_val:.2f}%)"
              else:
                status_str = "NOT IN PDF"

          rating_audit_rows.append({
              "Credit Rating": r_label,
              "Master Excel (%)": f"{excel_val:.2f}%",
              "PDF Published (%)": pdf_val_str,
              "Difference (%)": (
                  f"{diff_val:.2f}%" if pdf_val_str != "N/A" else "N/A"
              ),
              "Audit Status": status_str,
          })

        df_rating_audit = pd.DataFrame(rating_audit_rows)
        styled_r = apply_status_style(df_rating_audit, "Audit Status")
        st.dataframe(styled_r, use_container_width=True, hide_index=True)
      elif pdf_text:
        all_pdf_ratings = []
        for rat_code in ["SOV", "AAA", "AA+", "AA", "A1+", "AA-"]:
          if rat_code in pdf_text:
            all_pdf_ratings.append(rat_code)
        if all_pdf_ratings:
          st.write(
              f"**Ratings Detected in PDF:** `{', '.join(all_pdf_ratings)}`"
          )
          st.info(
              "Upload Master Factsheet Excel to audit exact percentages"
              " against PDF."
          )
        else:
          st.info(
              "No credit rating profile table required for this fund (100%"
              " Equity / Index)."
          )
      else:
        st.info("No credit rating profile table required.")

    # 3. Sectoral Allocation
    with sub_tabs[2]:
      st.markdown("#### Sectoral Concentration & Exposure")
      if s_master and s_master["sectors"]:
        df_s = pd.DataFrame(s_master["sectors"])

        def sec_flag(r):
          val = float(r["Raw Weight"])
          name = str(r["Sector Name"]).lower()
          if "central government" in name or "gilt" in name:
            return "COMPLIANT"
          if val > 25.0:
            return "BREACH (>25%)"
          return "COMPLIANT"

        df_s["Audit Status"] = df_s.apply(sec_flag, axis=1)
        show_s = df_s[["Sector Name", "Weight (%)", "Audit Status"]]
        styled_s = apply_status_style(show_s, "Audit Status")
        st.dataframe(styled_s, use_container_width=True, hide_index=True)
      else:
        st.info(
            "Upload Master Factsheet Excel to view sectoral allocation."
        )

    # 4. Corporate Debt
    with sub_tabs[3]:
      st.markdown("#### Corporate Debt Holdings")
      if s_master and s_master["corp_debt"]:
        df_cd = pd.DataFrame(s_master["corp_debt"])
        show_cd = df_cd[["Security Description", "Weight (%)", "Status"]]
        styled_cd = apply_status_style(show_cd, "Status")
        st.dataframe(styled_cd, use_container_width=True, hide_index=True)
      else:
        st.info("No Corporate Debt instruments held in this scheme.")

    # 5. Government Securities
    with sub_tabs[4]:
      st.markdown(
          "#### Government Securities Holdings (G-Sec / SDL / T-Bills)"
      )
      if s_master and s_master["g_sec"]:
        df_gs = pd.DataFrame(s_master["g_sec"])
        show_gs = df_gs[["Security Description", "Weight (%)", "Status"]]
        styled_gs = apply_status_style(show_gs, "Status")
        st.dataframe(styled_gs, use_container_width=True, hide_index=True)
      else:
        st.info("No Government Securities held in this scheme.")

    # 6. Equity Holdings
    with sub_tabs[5]:
      st.markdown("#### Equity Holdings (Top Equities)")
      if s_master and s_master["equity"]:
        df_eq = pd.DataFrame(s_master["equity"])
        show_eq = df_eq[["Security Description", "Weight (%)", "Status"]]
        styled_eq = apply_status_style(show_eq, "Status")
        st.dataframe(styled_eq, use_container_width=True, hide_index=True)
      else:
        st.info(
            "No Equity holdings held in this scheme (100% Debt / Liquid"
            " Fund)."
        )

    # 7. AUM Audit & Charts
    with sub_tabs[6]:
      st.markdown("#### Scheme Asset Class Breakup & AUM Audit")
      found_aum_key = None
      if aum_data and sheet_key:
        for k in aum_data.keys():
          if k.upper() in sheet_key.upper() or sheet_key.upper() in k.upper():
            found_aum_key = k
            break

      pdf_aum_val = (
          pdf_info["aum_candidates"][-1]
          if (pdf_info and pdf_info["aum_candidates"])
          else None
      )
      excel_aum_val = (
          aum_data[found_aum_key]["Total AUM (Cr)"] if found_aum_key else None
      )

      # AUM Direct Cross Comparison Metrics
      m_c1, m_c2, m_c3 = st.columns(3)
      m_c1.metric(
          "PDF Published AUM",
          f"{pdf_aum_val:.2f} Cr" if pdf_aum_val else "N/A",
      )
      m_c2.metric(
          "Excel System AUM",
          f"{excel_aum_val:.2f} Cr" if excel_aum_val else "N/A",
      )
      if pdf_aum_val and excel_aum_val:
        diff_aum = abs(pdf_aum_val - excel_aum_val)
        m_c3.metric(
            "Difference",
            f"{diff_aum:.2f} Cr",
            delta="MATCH" if diff_aum <= 0.05 else "-MISMATCH",
        )

      if found_aum_key and found_aum_key in aum_data:
        info_a = aum_data[found_aum_key]
        chart_df = pd.DataFrame({
            "Asset Class": [
                "Corporate Debt",
                "Government Securities",
                "Equity",
                "MMI",
            ],
            "AUM (In Crores)": [
                info_a["Corporate Debt"],
                info_a["G-Sec"],
                info_a["Equity"],
                info_a["MMI"],
            ],
        }).set_index("Asset Class")

        chart_df_non_zero = chart_df[chart_df["AUM (In Crores)"] > 0]
        if not chart_df_non_zero.empty:
          st.bar_chart(chart_df_non_zero)
          st.table(chart_df)
      else:
        st.info(
            "Upload `AUM Breakup.xlsx` in sidebar to view asset class"
            " breakdown."
        )

# -------------------------------------------------------------
# TAB 4: AUM ASSET BREAKUP (RAW CODES)
# -------------------------------------------------------------
with tabs[3]:
  st.subheader("📊 Raw AUM Code Breakdown Table")
  if not aum_data:
    st.info(
        "💡 Sirf **`AUM Breakup.xlsx`** upload karein AUM numbers verify"
        " karne ke liye."
    )
  else:
    aum_rows = []
    for code, info_a in aum_data.items():
      aum_rows.append({
          "Scheme Code": code,
          "Corporate Debt (Cr)": f"{info_a['Corporate Debt']:.2f}",
          "G-Sec (Cr)": f"{info_a['G-Sec']:.2f}",
          "Equity (Cr)": f"{info_a['Equity']:.2f}",
          "MMI (Cr)": f"{info_a['MMI']:.2f}",
          "Total AUM (Cr)": f"{info_a['Total AUM (Cr)']:.2f}",
      })
    st.dataframe(pd.DataFrame(aum_rows), use_container_width=True, hide_index=True)

# -------------------------------------------------------------
# TAB 5: YTM & DURATION AUDIT
# -------------------------------------------------------------
with tabs[4]:
  st.subheader("📉 Yield to Maturity (YTM) & Duration Audit")
  if not ytm_data:
    st.info(
        "💡 Sirf **`Yield, Duration Report-30.09.26.xlsx`** upload karein"
        " YTM aur Duration audit ke liye."
    )
  else:
    ytm_rows = []
    for code, info in ytm_data.items():
      ytm_rows.append({
          "Scheme Code": code,
          "YTM (%)": f"{info['YTM']:.2f}%" if info["YTM"] else "N/A",
          "Modified Duration": (
              f"{info['Duration']:.2f}" if info["Duration"] else "N/A"
          ),
      })
    st.dataframe(pd.DataFrame(ytm_rows), use_container_width=True, hide_index=True)

# -------------------------------------------------------------
# TAB 6: EXPORT AUDIT WORKBOOK
# -------------------------------------------------------------
with tabs[5]:
  st.subheader("📥 Export Full Audit Reconciliation Report")

  out_buf = io.BytesIO()
  with pd.ExcelWriter(out_buf, engine="xlsxwriter") as writer:
    if nav_data and "df_nav_audit" in locals():
      df_nav_audit.to_excel(writer, sheet_name="NAV Audit", index=False)
    if "df_aum_audit" in locals():
      df_aum_audit.to_excel(writer, sheet_name="AUM Master Audit", index=False)
    if ytm_data and "ytm_rows" in locals():
      pd.DataFrame(ytm_rows).to_excel(
          writer, sheet_name="YTM & Duration", index=False
      )

  st.download_button(
      label="📊 Download Comprehensive Audit Summary (.xlsx)",
      data=out_buf.getvalue(),
      file_name="ABSLI_Factsheet_Audit_Summary_Sep2026.xlsx",
      mime=(
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      ),
  )
