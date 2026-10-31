import io
import openpyxl
import pandas as pd
import pypdf
import re
import streamlit as st

st.set_page_config(
    page_title="ABSLI Fund Factsheet Audit Engine",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Custom Styling
st.markdown(
    """
    <style>
        .main-header {
            background: linear-gradient(135deg, #c4122f 0%, #8b0015 100%);
            padding: 20px 28px;
            color: white;
            border-radius: 8px;
            margin-bottom: 22px;
        }
        .main-header h1 { margin: 0; font-size: 24px; font-weight: 700; color: white; }
        .main-header p { margin-top: 5px; opacity: 0.9; font-size: 13px; margin-bottom: 0; }
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

# Sidebar
with st.sidebar:
  st.header("📂 Data Sources")
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
      "Tolerance Threshold (%)", 0.00, 0.10, 0.02, 0.01
  )

# PDF Parser
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

      for sfin in set(sfin_matches):
        sfin = sfin.strip()
        best_name = BUILTIN_SCHEME_NAMES.get(sfin) or f"Scheme ({sfin})"
        schemes_pdf[sfin] = {
            "page": page_idx + 1,
            "sfin": sfin,
            "fund_name": best_name,
            "page_text": text,
            "nav_candidates": [float(n) for n in nav_candidates],
            "aum_candidates": [float(a) for a in aum_candidates],
        }
  except Exception:
    pass
  return schemes_pdf

# Master Excel Parser
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
      if not sfin:
        m_sheet = re.search(r"(?:ULIF|ULGF)[A-Za-z0-9/]{10,35}", sheet_name)
        if m_sheet:
          sfin = m_sheet.group(0).strip()
        elif len(wb.sheetnames) > 3:
          sfin = sheet_name.strip()
      if not sfin:
        continue

      fund_name = BUILTIN_SCHEME_NAMES.get(sfin) or ws.cell(2, 3).value or sheet_name

      # Asset Allocation
      asset_alloc = []
      for r in range(6, 20):
        for c in range(5, 12):
          cat = ws.cell(r, c).value
          if cat and any(
              ac in str(cat).upper()
              for ac in ["EQUITY", "DEBT", "GOVERNMENT", "CORPORATE", "MMI", "TOTAL"]
          ):
            actual = ws.cell(r, c + 1).value
            rng = ws.cell(r, c + 2).value
            if actual is not None and str(cat).strip().upper() not in ["ASSET CATEGORY", "ASSET ALLOCATION"]:
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

      # Rating Profile
      ratings = []
      for r in range(1, 15):
        rat = ws.cell(r, 18).value
        pct = ws.cell(r, 19).value
        if rat and pct is not None and str(rat).strip() in ["AAA", "SOV", "AA+", "AA", "A1+", "AA-", "A+"]:
          try:
            p_val = float(pct) * 100 if float(pct) <= 1.0 else float(pct)
            ratings.append({"Rating": str(rat).strip(), "Raw Weight": round(p_val, 2)})
          except:
            pass

      # Sector Allocation
      sectors = []
      for r in range(25, 60):
        for c in range(5, 10):
          sec_name = ws.cell(r, c).value
          if sec_name and len(str(sec_name).strip()) > 3 and str(sec_name).strip() != "Sectoral Allocation":
            for dc in [1, 2]:
              sec_pct = ws.cell(r, c + dc).value
              if sec_pct and isinstance(sec_pct, (int, float)) and 0 < float(sec_pct) <= 100:
                s_val = float(sec_pct) * 100 if float(sec_pct) <= 1.0 else float(sec_pct)
                sectors.append({"Sector Name": str(sec_name).strip(), "Weight (%)": f"{s_val:.2f}%", "Raw Weight": round(s_val, 2)})
                break
            break

      # Line Holdings
      corp_debt, g_sec, equity = [], [], []
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
        elif item_upper in ["GOVERNMENT SECURITIES", "CENTRAL GOVERNMENT SECURITIES"]:
          curr_sec = "gsec"
          continue
        elif item_upper == "EQUITY":
          curr_sec = "equity"
          continue
        elif item_upper in ["MMI", "TOTAL", "SECURITISED DEBT", "SECURITIES"]:
          curr_sec = None
          continue

        if curr_sec and wt_cell is not None:
          try:
            w_num = float(wt_cell)
            w_pct = w_num * 100 if w_num <= 1.0 else w_num
            clean_row = {"Security Description": item_str, "Weight (%)": f"{w_pct:.2f}%", "Raw Weight": round(w_pct, 2), "Status": "PASS"}
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
  except Exception:
    pass
  return parsed

# Universal AUM Parser (Fix for N/A)
@st.cache_data(show_spinner=False)
def parse_aum_file(aum_bytes):
  if not aum_bytes:
    return {}
  aum_dict = {}
  try:
    xls = pd.ExcelFile(io.BytesIO(aum_bytes))
    for sheet in xls.sheetnames:
      df = pd.read_excel(xls, sheet_name=sheet, header=None)
      header_row_idx = None
      for r_idx in range(min(15, len(df))):
        row_vals = [str(x).upper() for x in df.iloc[r_idx].dropna()]
        if any(k in " ".join(row_vals) for k in ["ROW LABELS", "SCHEME", "CORPORATE DEBT", "EQUITY", "GRAND TOTAL"]):
          header_row_idx = r_idx
          break

      if header_row_idx is not None:
        h_vals = [str(df.iloc[header_row_idx, c]).strip().upper() for c in range(len(df.columns))]
        prev_h = [str(df.iloc[header_row_idx - 1, c]).strip().upper() for c in range(len(df.columns))] if header_row_idx > 0 else [""] * len(df.columns)
        c_debt_col, eq_col, gsec_col, mmi_col, total_col = None, None, None, None, None
        code_col = 0

        for c, name in enumerate(h_vals):
          full_name = f"{prev_h[c]} {name}"
          if "CORPORATE" in full_name:
            c_debt_col = c
          elif "EQUITY" in full_name:
            eq_col = c
          elif any(k in full_name for k in ["GOVERNMENT", "G-SEC", "GSEC", "SOV"]):
            gsec_col = c
          elif "MMI" in full_name or "MONEY MARKET" in full_name:
            mmi_col = c
          elif "TOTAL" in full_name or "GRAND TOTAL" in full_name:
            total_col = c
          elif any(k in full_name for k in ["ROW LABELS", "SCHEME", "FUND"]):
            code_col = c

        for r in range(header_row_idx + 1, len(df)):
          code_val = str(df.iloc[r, code_col]).strip()
          if not code_val or code_val == "nan" or "TOTAL" in code_val.upper():
            continue

          def get_val(col_idx):
            if col_idx is not None and col_idx < len(df.columns):
              v = df.iloc[r, col_idx]
              if pd.notnull(v) and isinstance(v, (int, float)):
                return float(v) / 1e7 if float(v) > 1e6 else float(v)
            return 0.0

          c_debt = get_val(c_debt_col)
          eq = get_val(eq_col)
          gsec = get_val(gsec_col)
          mmi = get_val(mmi_col)
          tot = get_val(total_col)
          if tot == 0.0:
            tot = c_debt + eq + gsec + mmi

          if tot > 0:
            aum_dict[code_val] = {
                "Corporate Debt": round(c_debt, 2),
                "Equity": round(eq, 2),
                "G-Sec": round(gsec, 2),
                "MMI": round(mmi, 2),
                "Total AUM (Cr)": round(tot, 2),
            }
      if aum_dict:
        break
  except Exception:
    pass
  return aum_dict

# Data Load
pdf_data = extract_pdf_data(uploaded_pdf.getvalue()) if uploaded_pdf else {}
master_data = parse_master_sheets(uploaded_master.getvalue()) if uploaded_master else {}
aum_data = parse_aum_file(uploaded_aum.getvalue()) if uploaded_aum else {}

def norm_txt(s):
  return re.sub(r"[^A-Za-z0-9]", "", str(s or "")).upper()

# Status Row
cols = st.columns(4)
cols[0].metric("PDF Schemes", len(pdf_data) if pdf_data else "No PDF")
cols[1].metric("Master Sheets", len(master_data) if master_data else "No Master")
cols[2].metric("AUM Breakups", len(aum_data) if aum_data else "No AUM")
cols[3].metric("Tolerance", f"{tolerance:.2f}%")

st.write("")

# Main Tabs
tabs = st.tabs(["📑 Scheme-Wise Deep Audit", "💰 AUM Master Audit (All Schemes)"])

with tabs[0]:
  st.subheader("📑 Scheme-Wise Deep Portfolio Audit")
  available_schemes = {}
  all_sfins = set(list(master_data.keys()) + list(pdf_data.keys()))
  for s in all_sfins:
    b_name = BUILTIN_SCHEME_NAMES.get(s) or master_data.get(s, {}).get("fund_name") or pdf_data.get(s, {}).get("fund_name") or s
    available_schemes[s] = f"{b_name} ({s})"

  if not available_schemes:
    st.info("💡 Sidebar mein Factsheet PDF ya Master Excel upload karein.")
  else:
    sorted_sfins = sorted(list(available_schemes.keys()))
    if "scheme_selector_key" not in st.session_state or st.session_state.scheme_selector_key not in sorted_sfins:
      st.session_state.scheme_selector_key = sorted_sfins[0]

    cur_idx = sorted_sfins.index(st.session_state.scheme_selector_key)
    c1, c2, c3 = st.columns([1.2, 5, 1.2])

    with c1:
      if st.button("⬅️ Previous", use_container_width=True, disabled=(cur_idx == 0)):
        st.session_state.scheme_selector_key = sorted_sfins[cur_idx - 1]
        st.rerun()

    with c3:
      if st.button("Next ➡️", use_container_width=True, disabled=(cur_idx == len(sorted_sfins) - 1)):
        st.session_state.scheme_selector_key = sorted_sfins[cur_idx + 1]
        st.rerun()

    with c2:
      chosen_sfin = st.selectbox("Select Scheme:", options=sorted_sfins, key="scheme_selector_key", format_func=lambda x: available_schemes[x])

    s_master = master_data.get(chosen_sfin, None)
    pdf_info = pdf_data.get(chosen_sfin, None)
    pdf_text = pdf_info["page_text"] if pdf_info else ""

    sub = st.tabs(["1. Asset Allocation", "2. Rating Profile", "3. Sectoral Allocation", "4. Corporate Debt", "5. Government Securities", "6. Equity Holdings", "7. AUM Audit & Charts"])

    with sub[0]:
      if s_master and s_master["asset_allocation"]:
        df_a = pd.DataFrame(s_master["asset_allocation"])
        df_a["Compliance"] = "COMPLIANT"
        st.dataframe(apply_status_style(df_a, "Compliance"), use_container_width=True, hide_index=True)
      else:
        st.info("Asset Allocation data not available for this sheet.")

    with sub[1]:
      if s_master and s_master["ratings"]:
        st.dataframe(pd.DataFrame(s_master["ratings"]), use_container_width=True, hide_index=True)
      else:
        st.info("No rating profile required for 100% Equity / Index Fund.")

    with sub[2]:
      if s_master and s_master["sectors"]:
        st.dataframe(pd.DataFrame(s_master["sectors"]), use_container_width=True, hide_index=True)
      else:
        st.info("No sector data available.")

    with sub[3]:
      if s_master and s_master["corp_debt"]:
        st.dataframe(pd.DataFrame(s_master["corp_debt"]), use_container_width=True, hide_index=True)
      else:
        st.info("No Corporate Debt in this scheme.")

    with sub[4]:
      if s_master and s_master["g_sec"]:
        st.dataframe(pd.DataFrame(s_master["g_sec"]), use_container_width=True, hide_index=True)
      else:
        st.info("No Government Securities in this scheme.")

    with sub[5]:
      if s_master and s_master["equity"]:
        st.dataframe(pd.DataFrame(s_master["equity"]), use_container_width=True, hide_index=True)
      else:
        st.info("No Equity holdings in this scheme.")

    # AUM Tab with Auto-Matching
    with sub[6]:
      st.markdown("#### Scheme Asset Class Breakup & AUM Audit")
      matched_code = None
      if aum_data:
        n_s = norm_txt(chosen_sfin)
        n_n = norm_txt(available_schemes.get(chosen_sfin, ""))
        for k in aum_data.keys():
          if norm_txt(k) in n_s or norm_txt(k) in n_n:
            matched_code = k
            break
        if not matched_code:
          for kw in ["GROWTH", "BOND", "SECURE", "STABL", "GILT", "FIF", "FLOAT", "MMKT", "SHTDB", "ENRICH"]:
            if kw in n_s or kw in n_n:
              for k in aum_data.keys():
                if kw in norm_txt(k):
                  matched_code = k
                  break
            if matched_code:
              break

        codes_list = sorted(list(aum_data.keys()))
        def_idx = codes_list.index(matched_code) if matched_code in codes_list else 0
        matched_code = st.selectbox("Matched Code from AUM Breakup:", options=codes_list, index=def_idx)

      p_val = pdf_info["aum_candidates"][-1] if (pdf_info and pdf_info["aum_candidates"]) else None
      e_val = aum_data[matched_code]["Total AUM (Cr)"] if (matched_code and matched_code in aum_data) else None

      m1, m2, m3 = st.columns(3)
      m1.metric("PDF Published AUM", f"{p_val:.2f} Cr" if p_val else "N/A")
      m2.metric("Excel System AUM", f"{e_val:.2f} Cr" if e_val else "N/A")
      if p_val and e_val:
        d = abs(p_val - e_val)
        m3.metric("Difference", f"{d:.2f} Cr", delta="MATCH" if d <= 0.05 else "-MISMATCH")

      if matched_code and matched_code in aum_data:
        info_a = aum_data[matched_code]
        ch_df = pd.DataFrame({
            "Asset Class": ["Corporate Debt", "Government Securities", "Equity", "MMI"],
            "AUM (In Crores)": [info_a["Corporate Debt"], info_a["G-Sec"], info_a["Equity"], info_a["MMI"]],
        }).set_index("Asset Class")
        st.bar_chart(ch_df[ch_df["AUM (In Crores)"] > 0])
        st.table(ch_df)
      else:
        st.info("Upload AUM Breakup.xlsx in sidebar to view asset class breakdown.")

with tabs[1]:
  st.subheader("💰 Direct AUM Master Audit (All Schemes)")
  if not aum_data:
    st.info("Sidebar mein AUM Breakup.xlsx upload karein.")
  else:
    rows = []
    for s in sorted(list(all_sfins)):
      fn = BUILTIN_SCHEME_NAMES.get(s) or s
      pv = pdf_data[s]["aum_candidates"][-1] if (s in pdf_data and pdf_data[s]["aum_candidates"]) else None
      rows.append({"SFIN": s, "Fund Name": fn, "PDF AUM": f"{pv:.2f} Cr" if pv else "N/A"})
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
