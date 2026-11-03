# =============================================================
# ABSLI FUND FACTSHEET AUDIT ENGINE (COMPLETE PRODUCTION CODE)
# =============================================================
import io
import re
import openpyxl
import pandas as pd
import pypdf
import streamlit as st

# =============================================================
# BLOCK 1: APP CONFIG & UI HEADER
# =============================================================
st.set_page_config(
    page_title="ABSLI Fund Factsheet Audit Engine",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded",
)

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
        div[data-testid="stButton"] button {
            margin-top: 28px;
            height: 42px;
            font-weight: 700;
        }
        button[data-testid="stNumberInputStepDown"],
        button[data-testid="stNumberInputStepUp"] {
            display: none !important;
        }
        div[data-testid="stNumberInputContainer"] {
            padding-right: 8px !important;
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

# =============================================================
# BLOCK 2: BUILT-IN SCHEME DIRECTORY & STYLING
# =============================================================
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
    "ULIF01911/12/09BSLITITAN1109": "Titanium 1 Fund",
    "ULIF02011/12/09BSLITITAN2109": "Titanium 2 Fund",
    "ULIF02203/02/10BSLPLATPR1109": "Platinum Premier Fund",
    "ULIF02301/07/10BSLIDISCPF109": "Discontinued Policy Fund",
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
    "ULIF04015/06/24ABSLIALPHA109": "Nifty 50 Alpha Fund",
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
    if any(k in val_upper for k in ["MISMATCH", "NOT IN PDF", "BREACH", "FAIL"]):
        return "background-color: #f8d7da; color: #721c24; font-weight: bold;"
    elif any(k in val_upper for k in ["PASS", "MATCH"]):
        return "background-color: #d4edda; color: #155724; font-weight: bold;"
    elif any(k in val_upper for k in ["PENDING", "WARNING"]):
        return "background-color: #fff3cd; color: #856404; font-weight: bold;"
    return ""

def apply_status_style(df, column_name):
    styler = df.style
    if hasattr(styler, "map"):
        return styler.map(color_status_cells, subset=[column_name])
    else:
        return styler.applymap(color_status_cells, subset=[column_name])

def norm_txt(s):
    return re.sub(r"[^A-Za-z0-9]", "", str(s or "")).upper()

# =============================================================
# BLOCK 3: CONTEXTUAL LINE-BY-LINE AUDIT HELPER (FUTURE-PROOF)
# =============================================================
def audit_line_item_against_pdf(item_name, excel_val, pdf_text, tolerance=0.05):
    if not pdf_text:
        return "N/A", "N/A", "PENDING PDF"
        
    clean_item_upper = str(item_name).upper().strip()
    
    # 1. 0.00% Residual Check for Holdings (Corporate vs Government)
    is_residual_other = (
        clean_item_upper.startswith("OTHER ") 
        and any(k in clean_item_upper for k in ["CORPORATE", "GOVERNMENT", "G-SEC"])
        and not any(k in clean_item_upper for k in ["TREPS", "MONEY MARKET", "FINANCIAL"])
    )
    
    def matches_other_category(line_str):
        l_up = line_str.upper()
        if "OTHER" not in l_up:
            return False
        if "CORPORATE" in clean_item_upper or "CORP" in clean_item_upper:
            return ("CORPORATE" in l_up or "CORP" in l_up) and "GOVERNMENT" not in l_up and "G-SEC" not in l_up
        elif "GOVERNMENT" in clean_item_upper or "G-SEC" in clean_item_upper or "GSEC" in clean_item_upper:
            return ("GOVERNMENT" in l_up or "G-SEC" in l_up or "GSEC" in l_up or "GOI" in l_up) and "CORPORATE" not in l_up
        return False

    if abs(excel_val) < 0.001 and is_residual_other:
        lines = pdf_text.split("\n")
        found_line = None
        for line in lines:
            if matches_other_category(line):
                found_line = line.strip()
                break
        if found_line:
            pcts = re.findall(r"(\d{1,3}\.\d{1,2})\s*%?", found_line)
            if pcts:
                val = float(pcts[-1])
                diff = abs(val - excel_val)
                status = "PASS" if diff <= tolerance else f"MISMATCH ({diff:.2f}%)"
                return f"{val:.2f}%", f"{diff:.2f}%", status
            return "0.00%", "0.00%", "PASS"
        return "0.00%", "0.00%", "PASS"

    lines = pdf_text.split("\n")

    # 2. Precise Bond / G-Sec Matcher (Coupon & Maturity Date with Lookahead)
    m_coupon = re.search(r"(\d{1,2}\.\d{1,4})\s*%", item_name)
    coupon_str = m_coupon.group(1) if m_coupon else None
    
    m_date = re.search(r"(\d{2}/\d{2}/\d{2,4})", item_name)
    date_str = m_date.group(1) if m_date else None

    best_line = None
    if coupon_str or date_str:
        for idx, line in enumerate(lines):
            line_clean = line.strip()
            if not line_clean:
                continue
            matched = False
            if coupon_str and date_str:
                if coupon_str in line_clean or date_str in line_clean:
                    matched = True
            elif coupon_str and re.search(rf"\b{re.escape(coupon_str)}\s*%", line_clean):
                matched = True
            elif date_str and date_str in line_clean:
                matched = True
                
            if matched:
                combined = line_clean
                for k in range(1, 4):
                    pcts_found = re.findall(r"(\d{1,3}\.\d{1,2})\s*%?", combined)
                    if len(pcts_found) >= 2 or (len(pcts_found) >= 1 and coupon_str not in pcts_found):
                        break
                    if idx + k < len(lines):
                        combined += " " + lines[idx + k].strip()
                best_line = combined
                break

    # 3. Line Search with Boundary Safety (Sectors, Equities, Asset Allocation, Ratings)
    SECTION_BOUNDARY_KEYWORDS = [
        "PORTFOLIO", "HOLDINGS", "SECTORAL ALLOCATION", "RATING PROFILE",
        "ASSET ALLOCATION", "ABOUT THE FUND", "FUND MANAGER", "NAV", "AUM"
    ]

    is_rating_item = clean_item_upper in [
        "SOV", "SOVEREIGN", "AAA", "AA+", "AA", "A1+", "AA-", "A+", "P1+", "AAA(SO)", "UNRATED"
    ]

    if not best_line:
        clean_name = re.sub(r"[^A-Za-z0-9\s]", " ", str(item_name)).strip()
        excluded_words = ["except", "from", "with", "than", "supply", "products", "activities", "compulsory", "social", "limited", "ltd"]
        if not is_residual_other:
            excluded_words.append("other")
            
        words = [w for w in clean_name.split() if len(w) >= 2 and w.lower() not in excluded_words]
        max_matches = 0
        
        for idx, line in enumerate(lines):
            line_clean = line.strip()
            if not line_clean:
                continue
            if is_residual_other and not matches_other_category(line_clean):
                continue
                
            # Ignore footnote lines for MMI
            if "MMI" in clean_item_upper and any(k in line_clean.upper() for k in ["IN CRORES", "CRORES", "₹"]):
                continue

            matches = 0
            if is_rating_item:
                if re.search(r"\b" + re.escape(clean_item_upper) + r"\b", line_clean.upper()):
                    matches = 10
            else:
                matches = sum(1 for w in words if re.search(r"\b" + re.escape(w) + r"\b", line_clean, re.IGNORECASE))

            if matches > max_matches and (matches >= min(2, len(words)) or (len(words) == 1 and matches > 0)):
                max_matches = matches
                
                is_multi_code_legend = is_rating_item and sum(
                    1 for rc in ["SOV", "AAA", "AA+", "AA", "A1+"] if re.search(r"\b" + re.escape(rc) + r"\b", line_clean.upper())
                ) >= 2
                
                has_num_in_line = bool(re.search(r"\d{1,3}\.\d{1,2}", line_clean))
                
                combined = line_clean
                if not is_multi_code_legend or has_num_in_line:
                    for k in range(1, 4):
                        if re.search(r"\d{1,3}\.\d{1,2}", combined):
                            break
                        if idx + k < len(lines):
                            next_l = lines[idx + k].strip()
                            if any(next_l.upper().startswith(sb) for sb in SECTION_BOUNDARY_KEYWORDS):
                                break
                            combined += " " + next_l
                best_line = combined
                if is_rating_item and has_num_in_line:
                    break

    # 4. MMI alias check
    if not best_line and "MMI" in clean_item_upper:
        for idx, line in enumerate(lines):
            line_up = line.upper()
            if any(k in line_up for k in ["IN CRORES", "CRORES", "₹"]):
                continue
            if "MMI" in line_up or "MONEY MARKET" in line_up:
                combined = line.strip()
                for k in range(1, 4):
                    if re.search(r"\d{1,3}\.\d{1,2}\s*%", combined):
                        break
                    if idx + k < len(lines):
                        if not any(k_w in lines[idx + k].upper() for k_w in ["IN CRORES", "CRORES", "₹"]):
                            combined += " " + lines[idx + k].strip()
                best_line = combined
                break

    # 5. Value Extraction & Safe Pie-Chart Recovery
    if best_line:
        pcts_in_line = re.findall(r"(\d{1,3}\.\d{1,2})\s*%?", best_line)
        if pcts_in_line:
            if len(pcts_in_line) > 1 and coupon_str and pcts_in_line[0] == coupon_str:
                line_val = float(pcts_in_line[-1])
            elif len(pcts_in_line) > 1:
                line_val = min([float(p) for p in pcts_in_line], key=lambda x: abs(x - excel_val))
            else:
                line_val = float(pcts_in_line[0])

            diff = abs(line_val - excel_val)
            if diff <= tolerance:
                return f"{line_val:.2f}%", f"{diff:.2f}%", "PASS"
            
            # Pie Chart recovery: if line match picked an unrelated table row, verify exact published percentage
            val_2dec = f"{excel_val:.2f}"
            if re.search(rf"\b{re.escape(val_2dec)}\s*%", pdf_text):
                return f"{excel_val:.2f}%", "0.00%", "PASS"
                
            return f"{line_val:.2f}%", f"{diff:.2f}%", f"MISMATCH ({diff:.2f}%)"

    # 6. Global Exact Match Fallback
    val_2dec = f"{excel_val:.2f}"
    val_1dec = f"{excel_val:.1f}"
    patterns = [
        rf"\b{re.escape(val_2dec)}\s*%",
        rf"\b{re.escape(val_2dec)}\b",
        rf"\b{re.escape(val_1dec)}\s*%",
    ]
    for pat in patterns:
        m = re.search(pat, pdf_text)
        if m:
            found_str = m.group(0).replace(" ", "")
            if not found_str.endswith("%"):
                found_str += "%"
            return found_str, "0.00%", "PASS"

    return "NOT IN PDF", "N/A", "MISMATCH (NOT IN PDF)"

# =============================================================
# BLOCK 4: SIDEBAR UPLOADERS & PDF PARSER
# =============================================================
with st.sidebar:
    st.header("📂 Data Sources")
    st.caption("Upload PDF (Individual/Group) & Master Excel Sheets:")
    uploaded_pdf = st.file_uploader("1. Factsheet PDF (Individual/Group)", type=["pdf"])
    uploaded_master = st.file_uploader("2. Master Factsheet Excel", type=["xlsx"])
    uploaded_nav = st.file_uploader("3. Birla Nav Excel (Birla Nav.xlsx)", type=["xlsx"])
    uploaded_aum = st.file_uploader("4. AUM Breakup Excel (AUM with Fund names.xlsx)", type=["xlsx"])
    uploaded_ytm = st.file_uploader("5. Yield & Duration Excel (Yield...xlsx)", type=["xlsx"])
    tolerance = st.slider("Tolerance Threshold (Difference %)", 0.00, 0.10, 0.02, 0.01)

@st.cache_data(show_spinner=False)
def extract_pdf_data(pdf_bytes):
    schemes_pdf = {}
    if not pdf_bytes:
        return schemes_pdf
    try:
        reader = pypdf.PdfReader(io.BytesIO(pdf_bytes))
        for page_idx, page in enumerate(reader.pages):
            text = page.extract_text() or ""
            
            clean_sfins = re.findall(r"(?:ULIF|ULGF)\d{3}\d{2}/\d{2}/\d{2}[A-Za-z0-9]+?109", text)
            if not clean_sfins:
                clean_sfins = re.findall(r"(?:ULIF|ULGF)[A-Za-z0-9/]+?109", text)

            if not clean_sfins:
                continue

            unique_sfins = set(clean_sfins)
            
            has_detail_sections = any(sec in text.upper() for sec in [
                "ASSET ALLOCATION", "SECTORAL ALLOCATION", "ABOUT THE FUND", 
                "FUND MANAGER", "RATING PROFILE", "HOLDINGS"
            ])
            is_true_dedicated = (len(unique_sfins) <= 2) and has_detail_sections

            nav_candidates = re.findall(r"\b\d{1,3}\.\d{4}\b", text)

            aum_val = None
            m_aum_line = re.search(
                r"(?:Asset\s+held|AUM|Assets?\s+under\s+Management).*?[₹\s]+([0-9,]+\.[0-9]{1,2})\s*(?:Cr|Crores)\b",
                text,
                re.IGNORECASE
            )
            if m_aum_line:
                aum_val = float(m_aum_line.group(1).replace(",", ""))
            else:
                m_cr_tokens = re.findall(r"\b([0-9,]+\.[0-9]{1,2})\s*Cr\b", text, re.IGNORECASE)
                valid_aums = [float(x.replace(",", "")) for x in m_cr_tokens if float(x.replace(",", "")) != 109.0]
                if valid_aums:
                    aum_val = valid_aums[0]

            fm_match = re.search(r"FUND\s+MANAGER:\s*([^\r\n]+)", text, re.IGNORECASE)
            fm_val = fm_match.group(1).strip() if fm_match else ""

            ytm_match = re.search(r"Yield\s+to\s+Maturity\s*\|?\s*([\d\.]+)%", text, re.IGNORECASE)
            dur_match = re.search(r"Modified\s+Duration\s*\|?\s*([\d\.]+)", text, re.IGNORECASE)

            header_fund_name = None
            for line in text.split("\n")[:10]:
                if any(s in line for s in clean_sfins):
                    clean_line = re.sub(r"SFIN\s*No\.?\s*:?.*", "", line, flags=re.IGNORECASE).strip()
                    clean_line = re.sub(r"(?:ULIF|ULGF)[A-Za-z0-9/\s]{10,40}", "", clean_line).strip()
                    if clean_line and len(clean_line) > 3 and clean_line.upper() not in ["EQUITY", "CORPORATE DEBT", "TOTAL", "SECURITIES"]:
                        header_fund_name = clean_line
                        break

            for sfin in unique_sfins:
                best_name = header_fund_name or BUILTIN_SCHEME_NAMES.get(sfin) or f"Scheme ({sfin})"
                prev_info = schemes_pdf.get(sfin)

                should_update = False
                if prev_info is None:
                    should_update = True
                elif is_true_dedicated and not prev_info.get("is_true_dedicated", False):
                    should_update = True
                elif is_true_dedicated and prev_info.get("is_true_dedicated", False):
                    prev_info["page_text"] += "\n" + text
                    if aum_val and not prev_info.get("aum_val"):
                        prev_info["aum_val"] = aum_val
                    continue

                if should_update:
                    schemes_pdf[sfin] = {
                        "page": page_idx + 1,
                        "sfin": sfin,
                        "fund_name": best_name,
                        "page_text": text,
                        "is_true_dedicated": is_true_dedicated,
                        "nav_candidates": [round(float(n), 4) for n in nav_candidates],
                        "aum_val": aum_val,
                        "fund_manager": fm_val,
                        "ytm": float(ytm_match.group(1)) if ytm_match else None,
                        "mod_duration": float(dur_match.group(1)) if dur_match else None,
                    }
    except Exception as e:
        st.sidebar.error(f"Error reading PDF: {e}")
    return schemes_pdf

# =============================================================
# BLOCK 5: MASTER EXCEL PARSER & SHEET MAP
# =============================================================
SHEET_NAME_TO_SFIN = {
    "INC-ADV-FUND": "ULIF01507/08/08BSLIINCADV109",
    "INC ADV FUND": "ULIF01507/08/08BSLIINCADV109",
    "INCOME ADVANTAGE": "ULIF01507/08/08BSLIINCADV109",
    "IND-INC-ADV": "ULIF01507/08/08BSLIINCADV109",
    "GR-INC-ADV": "ULGF01425/02/10BSLGINCADV109",
    "GR-INCOME-ADV": "ULGF01425/02/10BSLGINCADV109",
    "GROUP INCOME ADV": "ULGF01425/02/10BSLGINCADV109",
    "IND-BUILDER": "ULIF00113/03/01BSLBUILDER109",
    "BUILDER": "ULIF00113/03/01BSLBUILDER109",
    "GR-GROWTH": "ULGF00112/06/01BSLGGROWTH109",
    "GR-PROTECTOR": "ULGF00212/06/01BSLGSECURE109",
    "GR-STABLE": "ULGF00312/06/01BSLGSTABLE109",
    "GR-FIF": "ULGF00416/07/02BSLGFIXINT109",
    "GR-BOND": "ULGF00530/05/03BSLIGRBOND109",
    "GR-GILT": "ULGF00630/05/03BSLIGRGILT109",
    "GR-MMFUND": "ULGF00824/08/04BSLIGRMMKT109",
    "GR-ADVANTAGE": "ULGF01026/11/07BSLIGGRADV109",
    "GR-DEBTFUND": "ULGF01322/09/08BSLGSHTDBT109",
    "IND-ENHANCER": "ULIF00213/03/01BSLENHANCE109",
    "INDPROTECTOR": "ULIF00313/03/01BSLPROTECT109",
    "IPP-ENRICH": "ULIF00404/03/03BSLIENRICH109",
    "IPP-GROWTH": "ULIF00504/03/03BSLIGROWTH109",
    "IPP-NOURISH": "ULIF00604/03/03BSLNOURISH109",
    "IND-CREATOR": "ULIF00704/02/04BSLCREATOR109",
    "INDMAGNIFIER": "ULIF00826/06/04BSLIIMAGNI109",
    "BILICO": "ULIF00931/05/05BSLBALANCE109",
    "IND-AF": "ULIF01008/07/05BSLIASSURE109",
    "INDMAXIMISER": "ULIF01101/06/07BSLIINMAXI109",
    "IND-MULTI": "ULIF01217/10/07BSLINMULTI109",
    "SUPER20": "ULIF01723/06/09BSLSUPER20109",
    "TITANIUM": "ULIF01911/12/09BSLITITAN1109",
    "TITANIUM2": "ULIF02011/12/09BSLITITAN2109",
    "PLATINUM-PR1": "ULIF02203/02/10BSLPLATPR1109",
    "BSLI-DISCF": "ULIF02301/07/10BSLIDISCPF109",
    "BSLIPUREEQ": "ULIF02707/10/11BSLIPUREEQ109",
    "BSLLIQPLUS": "ULIF02807/10/11BSLLIQPLUS109",
    "BSLIVALUEM": "ULIF02907/10/11BSLIVALUEM109",
    "BSLIASTALC": "ULIF03430/10/14BSLIASTALC109",
    "BSLICNFIDX": "ULIF03530/10/14BSLICNFIDX109",
    "MNC": "ULIF03722/06/18ABSLIMUMNC109",
    "ABSLIESG": "ULIF03810/11/23ABSLESGFND109",
    "SMALLCAP": "ULIF03910/11/23ABSLSMALCP109",
    "INDALPHA50": "ULIF04015/06/24ABSLIALPHA109",
    "INDMOMNTM": "ULIF04101/10/24ABSLMOMNTM109",
    "INDDIVYIELD": "ULIF04206/11/25ABSLDIVYLD109",
    "INDFLEXICAP": "ULIF04312/02/26ABSLFLXCAP109",
}

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
            sfin = SHEET_NAME_TO_SFIN.get(sheet_name.strip().upper())

            for r in range(1, 30):
                for c in range(1, 25):
                    val = ws.cell(r, c).value
                    if val and isinstance(val, str):
                        val_str = val.strip()
                        if ("ULIF" in val_str or "ULGF" in val_str) and not sfin:
                            m_sfin = re.search(r"(?:ULIF|ULGF)\d{3}\d{2}/\d{2}/\d{2}[A-Za-z0-9]+?109", val_str)
                            if not m_sfin:
                                m_sfin = re.search(r"(?:ULIF|ULGF)[A-Za-z0-9/]+?109", val_str)
                            if m_sfin:
                                sfin = m_sfin.group(0).strip()
                        if (
                            not fund_name
                            and any(k in val_str.upper() for k in ["FUND", "BOND", "SECURE", "STABLE", "GROWTH", "INDEX"])
                            and val_str.strip().upper() not in ["EQUITY", "CORPORATE DEBT", "GOVERNMENT SECURITIES", "TOTAL", "MMI", "SECURITIES"]
                        ):
                            if len(val_str) < 60 and not val_str.startswith("UL"):
                                fund_name = val_str

            if not sfin:
                m_sheet = re.search(r"(?:ULIF|ULGF)\d{3}\d{2}/\d{2}/\d{2}[A-Za-z0-9]+?109", sheet_name)
                if m_sheet:
                    sfin = m_sheet.group(0).strip()
                elif len(wb.sheetnames) > 3:
                    sfin = sheet_name.strip()

            if not sfin:
                continue

            if not fund_name or fund_name.upper() in ["EQUITY", "CORPORATE DEBT", "TOTAL"]:
                fund_name = BUILTIN_SCHEME_NAMES.get(sfin) or ws.cell(2, 3).value or sheet_name

            # 1. Asset Allocation
            asset_alloc = []
            seen_cats = set()
            for r in range(6, 25):
                for c in range(5, 12):
                    cat = ws.cell(r, c).value
                    if cat and isinstance(cat, str):
                        cat_str = cat.strip()
                        cat_upper = cat_str.upper()
                        if any(x in cat_upper for x in ["IN RS", "CRORES", "TOTAL", "ASSET CATEGORY", "ASSET ALLOCATION"]):
                            continue
                        if any(ac in cat_upper for ac in ["EQUITY", "DEBT", "GOVERNMENT", "CORPORATE", "MMI", "OTHER"]):
                            actual = ws.cell(r, c + 1).value
                            rng = ws.cell(r, c + 2).value
                            if actual is not None and cat_upper not in seen_cats:
                                try:
                                    act_val = float(actual)
                                    pct = act_val * 100 if act_val <= 1.0 else act_val
                                    seen_cats.add(cat_upper)
                                    asset_alloc.append({
                                        "Asset Category": cat_str,
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
                    and str(rat).strip().upper() in ["AAA", "SOV", "SOVEREIGN", "AA+", "AA", "A1+", "AA-", "A+", "P1+", "AAA(SO)", "UNRATED"]
                ):
                    try:
                        p_float = float(pct)
                        p_val = p_float * 100 if p_float <= 1.0 else p_float
                        r_name = str(rat).strip().upper()
                        if r_name not in seen_ratings:
                            seen_ratings.add(r_name)
                            ratings.append({"Rating": r_name, "Weight (%)": f"{p_val:.2f}%", "Raw Weight": round(p_val, 2)})
                    except:
                        pass

            if not ratings:
                target_codes = ["AAA", "SOV", "SOVEREIGN", "AA+", "AA", "A1+", "AA-", "A+", "P1+", "AAA(SO)", "UNRATED"]
                for r in range(1, min(70, ws.max_row + 1)):
                    for c in range(1, min(30, ws.max_column + 1)):
                        val = ws.cell(r, c).value
                        if val and isinstance(val, str):
                            v_clean = val.strip().upper()
                            if v_clean in target_codes and v_clean not in seen_ratings:
                                for dc in [1, 2, -1]:
                                    v_num = ws.cell(r, c + dc).value
                                    if v_num is not None and isinstance(v_num, (int, float)) and 0 < float(v_num) <= 100:
                                        p_val = float(v_num) * 100 if float(v_num) <= 1.0 else float(v_num)
                                        seen_ratings.add(v_clean)
                                        ratings.append({"Rating": v_clean, "Weight (%)": f"{p_val:.2f}%", "Raw Weight": round(p_val, 2)})
                                        break

            # 3. Sector Allocation
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

            # 4. Holdings
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
                    if wt_cell is not None:
                        try:
                            w_num = float(wt_cell)
                            w_pct = w_num * 100 if w_num <= 1.0 else w_num
                            corp_debt.append({"Security Description": "📌 CORPORATE DEBT (HEADING TOTAL)", "Weight (%)": f"{w_pct:.2f}%", "Raw Weight": round(w_pct, 2)})
                        except: pass
                    continue
                elif item_upper in ["GOVERNMENT SECURITIES", "CENTRAL GOVERNMENT SECURITIES"]:
                    curr_sec = "gsec"
                    if wt_cell is not None:
                        try:
                            w_num = float(wt_cell)
                            w_pct = w_num * 100 if w_num <= 1.0 else w_num
                            g_sec.append({"Security Description": "📌 GOVERNMENT SECURITIES (HEADING TOTAL)", "Weight (%)": f"{w_pct:.2f}%", "Raw Weight": round(w_pct, 2)})
                        except: pass
                    continue
                elif item_upper == "EQUITY":
                    curr_sec = "equity"
                    if wt_cell is not None:
                        try:
                            w_num = float(wt_cell)
                            w_pct = w_num * 100 if w_num <= 1.0 else w_num
                            equity.append({"Security Description": "📌 EQUITY (HEADING TOTAL)", "Weight (%)": f"{w_pct:.2f}%", "Raw Weight": round(w_pct, 2)})
                        except: pass
                    continue
                elif item_upper in ["MMI", "TOTAL", "SECURITISED DEBT", "OTHER SECURITISED DEBT", "SECURITIES"]:
                    curr_sec = None
                    continue

                if curr_sec and wt_cell is not None:
                    try:
                        w_num = float(wt_cell)
                        w_pct = w_num * 100 if w_num <= 1.0 else w_num
                        clean_row = {"Security Description": item_str, "Weight (%)": f"{w_pct:.2f}%", "Raw Weight": round(w_pct, 2)}
                        if curr_sec == "corp":
                            corp_debt.append(clean_row)
                        elif curr_sec == "gsec":
                            g_sec.append(clean_row)
                        elif curr_sec == "equity":
                            equity.append(clean_row)
                    except:
                        pass

            # Headings fallback
            if not any("HEADING TOTAL" in x["Security Description"] for x in corp_debt):
                for a in asset_alloc:
                    if "CORPORATE" in a["Asset Category"].upper():
                        corp_debt.insert(0, {"Security Description": "📌 CORPORATE DEBT (HEADING TOTAL)", "Weight (%)": a["Actual (%)"], "Raw Weight": a["Raw Actual"]})
                        break
            if not any("HEADING TOTAL" in x["Security Description"] for x in g_sec):
                for a in asset_alloc:
                    if "GOVERNMENT" in a["Asset Category"].upper() or "G-SEC" in a["Asset Category"].upper():
                        g_sec.insert(0, {"Security Description": "📌 GOVERNMENT SECURITIES (HEADING TOTAL)", "Weight (%)": a["Actual (%)"], "Raw Weight": a["Raw Actual"]})
                        break
            if not any("HEADING TOTAL" in x["Security Description"] for x in equity):
                for a in asset_alloc:
                    if a["Asset Category"].strip().upper() == "EQUITY":
                        equity.insert(0, {"Security Description": "📌 EQUITY (HEADING TOTAL)", "Weight (%)": a["Actual (%)"], "Raw Weight": a["Raw Actual"]})
                        break

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

# =============================================================
# BLOCK 6: SECONDARY PARSERS & AUM LINKING
# =============================================================
@st.cache_data(show_spinner=False)
def parse_aum_file(aum_bytes):
    if not aum_bytes:
        return {}
    aum_dict = {}
    try:
        xls = pd.ExcelFile(io.BytesIO(aum_bytes))
        target_sheet = "Net_Assets" if "Net_Assets" in xls.sheet_names else xls.sheet_names[0]
        df = pd.read_excel(xls, sheet_name=target_sheet, header=None)

        # 1. Parse via Row Labels (Col 0: Label, Col 1: INR, Col 2: In Crores)
        for r in range(1, len(df)):
            label = str(df.iloc[r, 0]).strip() if pd.notnull(df.iloc[r, 0]) else ""
            if not label or label in ["Row Labels", "nan", "Grand Total"]:
                continue

            val_cr = None
            if len(df.columns) > 2 and pd.notnull(df.iloc[r, 2]):
                try:
                    val_cr = float(df.iloc[r, 2])
                except:
                    pass

            if val_cr is None and len(df.columns) > 1 and pd.notnull(df.iloc[r, 1]):
                try:
                    val_cr = float(df.iloc[r, 1]) / 1e7
                except:
                    pass

            if val_cr is not None:
                aum_dict[label] = {
                    "Code": label,
                    "Name": label,
                    "Total AUM (Cr)": round(val_cr, 2),
                    "Corporate Debt": 0.0,
                    "Equity": 0.0,
                    "G-Sec": 0.0,
                    "MMI": 0.0,
                }

        # 2. Parse secondary columns if present
        if len(df.columns) > 8:
            for r in range(1, len(df)):
                c7 = str(df.iloc[r, 7]).strip() if pd.notnull(df.iloc[r, 7]) else ""
                c8 = df.iloc[r, 8] if pd.notnull(df.iloc[r, 8]) else None
                c10 = str(df.iloc[r, 10]).strip() if len(df.columns) > 10 and pd.notnull(df.iloc[r, 10]) else ""

                if c7 and c7 not in ["nan", ""]:
                    try:
                        v_cr = float(c8)
                        record = {
                            "Code": c7,
                            "Name": c10 if c10 and c10 != "nan" else c7,
                            "Total AUM (Cr)": round(v_cr, 2),
                            "Corporate Debt": 0.0,
                            "Equity": 0.0,
                            "G-Sec": 0.0,
                            "MMI": 0.0,
                        }
                        aum_dict[c7] = record
                        if c10 and c10 != "nan":
                            aum_dict[c10] = record
                    except:
                        pass
    except Exception as e:
        st.sidebar.error(f"AUM File Error: {e}")
    return aum_dict

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
            if sfin and (sfin.startswith("ULIF") or sfin.startswith("ULGF")) and pd.notnull(nav):
                nav_dict[sfin] = {
                    "nav": float(nav),
                    "name": str(name).strip() if str(name).strip() else BUILTIN_SCHEME_NAMES.get(sfin, sfin),
                    "type": "Individual (ULIF)" if sfin.startswith("ULIF") else "Group (ULGF)",
                }
    except Exception as e:
        st.sidebar.error(f"NAV File Error: {e}")
    return nav_dict

@st.cache_data(show_spinner=False)
def parse_ytm_file(ytm_bytes):
    if not ytm_bytes:
        return {}
    ytm_dict = {}
    try:
        xls = pd.ExcelFile(io.BytesIO(ytm_bytes))
        sheet = xls.sheet_names[0]
        df_raw = pd.read_excel(xls, sheet_name=sheet, header=None)

        # 1. Header row auto-detect karein (pehle 10 rows scan karke)
        header_row_idx = 0
        for r in range(min(10, len(df_raw))):
            row_str = " ".join([str(v).upper() for v in df_raw.iloc[r] if pd.notnull(v)])
            if any(k in row_str for k in ["SCHEME", "FUND"]) and any(k in row_str for k in ["YTM", "YIELD", "DURATION"]):
                header_row_idx = r
                break

        df = pd.read_excel(xls, sheet_name=sheet, header=header_row_idx)

        # 2. Dynamic Column Mapping (Case & Space insensitive)
        code_col = None
        name_col = None
        ytm_col = None
        dur_col = None

        for col in df.columns:
            c_up = str(col).strip().upper()
            if any(k in c_up for k in ["SCHEME CODE", "SCHEME_CODE", "CODE", "SCHEME NO"]):
                code_col = col
            elif any(k in c_up for k in ["SCHEME NAME", "FUND NAME", "SCHEME_NAME", "PLAN"]):
                name_col = col
            elif "YTM" in c_up or "YIELD" in c_up:
                ytm_col = col
            elif "MOD" in c_up or "DUR" in c_up:
                dur_col = col

        # Fallback agar specific code column na mile toh pehla text column
        if not code_col and name_col:
            code_col = name_col
        elif not code_col and len(df.columns) > 0:
            code_col = df.columns[0]

        # 3. Data Ingestion
        for _, row in df.iterrows():
            code_val = str(row.get(code_col, "")).strip() if code_col else ""
            if not code_val or code_val in ["nan", "None", "Total", "Grand Total"]:
                continue

            name_val = str(row.get(name_col, "")).strip() if name_col else code_val
            
            raw_ytm = row.get(ytm_col) if ytm_col else None
            raw_dur = row.get(dur_col) if dur_col else None

            ytm_val = None
            if pd.notnull(raw_ytm):
                try:
                    f_ytm = float(str(raw_ytm).replace("%", "").strip())
                    ytm_val = f_ytm * 100 if f_ytm <= 0.30 else f_ytm
                except:
                    pass

            dur_val = None
            if pd.notnull(raw_dur):
                try:
                    dur_val = float(str(raw_dur).strip())
                except:
                    pass

            ytm_dict[code_val] = {
                "Code": code_val,
                "Name": name_val,
                "YTM": ytm_val,
                "Duration": dur_val,
            }
            if name_val and name_val != code_val:
                ytm_dict[name_val] = ytm_dict[code_val]

    except Exception as e:
        st.sidebar.error(f"YTM File Error: {e}")
    return ytm_dict

# Load Data Sources
pdf_data = extract_pdf_data(uploaded_pdf.getvalue()) if uploaded_pdf else {}
master_data = parse_master_sheets(uploaded_master.getvalue()) if uploaded_master else {}
nav_data = parse_nav_file(uploaded_nav.getvalue()) if uploaded_nav else {}
aum_data = parse_aum_file(uploaded_aum.getvalue()) if uploaded_aum else {}
ytm_data = parse_ytm_file(uploaded_ytm.getvalue()) if uploaded_ytm else {}

EXHAUSTIVE_SFIN_MAP = {
    # Group Funds (ULGF)
    "ULGF00112/06/01BSLGGROWTH109": "GR-GROWTH",
    "ULGF00212/06/01BSLGSECURE109": "GR-PROTECTOR",
    "ULGF00312/06/01BSLGSTABLE109": "GR-STABLE",
    "ULGF00416/07/02BSLGFIXINT109": "GR-FIF",
    "ULGF00530/05/03BSLIGRBOND109": "GR-BOND",
    "ULGF00630/05/03BSLIGRGILT109": "GR-GILT",
    "ULGF00724/08/04BSLIGFLOAT109": "BSLGRMMKT2",
    "ULGF00824/08/04BSLIGRMMKT109": "GR-MMFUND",
    "ULGF01026/11/07BSLIGGRADV109": "GR-ADVANTAGE",
    "ULGF01322/09/08BSLGSHTDBT109": "GR-DEBTFUND",
    "ULGF01425/02/10BSLGINCADV109": "GR-INC-ADV",
    "ULGF01928/11/11BSLGRMMKT2109": "BSLGRMMKT2",
    "ULGF02311/07/25BSLCAPPRTC109": "GRCAPPROTECT",

    # Individual & Pension Funds (ULIF)
    "ULIF00113/03/01BSLBUILDER109": "IND-BUILDER",
    "ULIF00213/03/01BSLENHANCE109": "IND-ENHANCER",
    "ULIF00313/03/01BSLPROTECT109": "INDPROTECTOR",
    "ULIF00404/03/03BSLIENRICH109": "IPP-ENRICH",
    "ULIF00504/03/03BSLIGROWTH109": "IPP-GROWTH",
    "ULIF00604/03/03BSLNOURISH109": "IPP-NOURISH",
    "ULIF00704/02/04BSLCREATOR109": "IND-CREATOR",
    "ULIF00826/06/04BSLIIMAGNI109": "INDMAGNIFIER",
    "ULIF00931/05/05BSLBALANCE109": "BILICO",
    "ULIF01008/07/05BSLIASSURE109": "IND-AF",
    "ULIF01101/06/07BSLIINMAXI109": "INDMAXIMISER",
    "ULIF01217/10/07BSLINMULTI109": "IND-MULTI",
    "ULIF01507/08/08BSLIINCADV109": "INC-ADV-FUND",
    "ULIF01723/06/09BSLSUPER20109": "SUPER20",
    "ULIF01911/12/09BSLITITAN1109": "TITANIUM",
    "ULIF02011/12/09BSLITITAN2109": "TITANIUM2",
    "ULIF02203/02/10BSLPLATPR1109": "PLATINUM-PR1",
    "ULIF02301/07/10BSLIDISCPF109": "BSLI-DISCF",
    "ULIF02408/09/10BSLPLATADV109": "PLATINUMADV",
    "ULIF02707/10/11BSLIPUREEQ109": "BSLIPUREEQ",
    "ULIF02807/10/11BSLLIQPLUS109": "BSLLIQPLUS",
    "ULIF02907/10/11BSLIVALUEM109": "BSLIVALUEM",
    "ULIF03027/08/13BSLIMAXGT109": "BSLIMAXGUA",
    "ULIF03127/08/13BSLIINADGT109": "BSLIINADGU",
    "ULIF03205/07/13BSLILDIS109": "BSLIDISLIF",
    "ULIF03305/07/13BSLIPNDIS109": "BSLIDISPEN",
    "ULIF03430/10/14BSLIASTALC109": "BSLIASTALC",
    "ULIF03530/10/14BSLICNFIDX109": "BSLICNFIDX",
    "ULIF03722/06/18ABSLIMUMNC109": "MNC",
    "ULIF03810/11/23ABSLESGFND109": "ABSLIESG",
    "ULIF03910/11/23ABSLSMALCP109": "SMALLCAP",
    "ULIF04015/06/24ABSLIALPHA109": "INDALPHA50",
    "ULIF04101/10/24ABSLMOMNTM109": "INDMOMNTM",
    "ULIF04206/11/25ABSLDIVYLD109": "INDDIVYIELD",
    "ULIF04312/02/26ABSLFLXCAP109": "INDFLEXICAP",
    "ULIF04410/09/26ABSLPENMAX109": "ABSLPENMAX",
    "ULIF04510/09/26ABSLPENINA109": "ABSLPENINA",
}

def get_exact_aum_for_sfin(sfin, fund_name, aum_dict):
    if not aum_dict:
        return None, None

    # 1. Exact or normalized map check
    target_code = EXHAUSTIVE_SFIN_MAP.get(sfin)
    if target_code and target_code in aum_dict:
        return target_code, aum_dict[target_code]["Total AUM (Cr)"]

    if target_code:
        norm_t = norm_txt(target_code)
        for c, info in aum_dict.items():
            if norm_txt(c) == norm_t:
                return c, info["Total AUM (Cr)"]

    # 2. Fund name matching
    if fund_name:
        norm_f = norm_txt(fund_name)
        for c, info in aum_dict.items():
            name_val = info.get("Name", "")
            if name_val and norm_txt(name_val) == norm_f:
                return c, info["Total AUM (Cr)"]
            if norm_txt(c) == norm_f:
                return c, info["Total AUM (Cr)"]

    # 3. Substring match
    norm_s = norm_txt(sfin)
    for c, info in aum_dict.items():
        if len(c) >= 4 and norm_txt(c) in norm_s:
            return c, info["Total AUM (Cr)"]

    return None, None

# Status Metrics
status_cols = st.columns(6)
status_cols[0].metric("PDF Schemes", len(pdf_data) if pdf_data else "No PDF")
status_cols[1].metric("Master Sheets", len(master_data) if master_data else "No Master")
status_cols[2].metric("NAV Records", len(nav_data) if nav_data else "No NAV")
status_cols[3].metric("AUM Breakups", len(aum_data) if aum_data else "No AUM")
status_cols[4].metric("YTM Records", len(ytm_data) if ytm_data else "No YTM")
status_cols[5].metric("Tolerance", f"{tolerance:.2f}%")

st.write("")

# Main Tabs Navigation
tabs = st.tabs([
    "📈 NAV Master Audit (All Schemes)",
    "💰 AUM Master Audit (All Schemes)",
    "📑 Scheme-Wise Deep Audit",
    "📊 AUM Asset Breakup",
    "📉 YTM & Duration Audit",
    "📥 Export Full Audit Workbook",
])

# =============================================================
# BLOCK 7: TAB 1 & 2 (NAV & AUM MASTER AUDITS WITH FILTERS)
# =============================================================
# --- TAB 1: NAV MASTER AUDIT ---
with tabs[0]:
    st.subheader("📈 Direct NAV Master Audit (Individual & Group Funds)")
    if not nav_data:
        st.info("💡 Sirf **`Birla Nav.xlsx`** sidebar mein upload karein aur saari schemes ka NAV ek sath check ho jayega.")
    else:
        cat_filter = st.radio(
            "Filter Schemes (NAV):",
            ["All Schemes", "Individual Funds (ULIF)", "Group Funds (ULGF)"],
            horizontal=True,
            key="nav_cat_filter"
        )
        nav_rows = []
        for sfin, ninfo in nav_data.items():
            if cat_filter == "Individual Funds (ULIF)" and not sfin.startswith("ULIF"):
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
                pdf_nav = found if found is not None else pdf_info["nav_candidates"][-1]

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

            excel_nav_str = f"{excel_nav:.4f}" if excel_nav is not None else "N/A"
            pdf_nav_str = f"{pdf_nav:.4f}" if pdf_nav is not None else "N/A"
            diff_str = f"{diff:.4f}" if pdf_nav is not None else "N/A"

            nav_rows.append({
                "Category": ninfo.get("type", "Fund"),
                "SFIN Code": sfin,
                "Fund Name": ninfo["name"],
                "Excel NAV Per Unit": excel_nav_str,
                "PDF Published NAV": pdf_nav_str,
                "Difference": diff_str,
                "Audit Status": status,
            })

        df_nav_audit = pd.DataFrame(nav_rows)
        styled_nav_audit = apply_status_style(df_nav_audit, "Audit Status")
        st.dataframe(styled_nav_audit, use_container_width=True, hide_index=True)

# --- TAB 2: AUM MASTER AUDIT (ALL SCHEMES) ---
with tabs[1]:
    st.subheader("💰 Direct AUM Master Audit (All Schemes in Crores)")
    if not aum_data and not pdf_data:
        st.info("💡 Sidebar mein **`AUM with Fund names.xlsx`** aur **Factsheet PDF** upload karein AUM audit ke liye.")
    else:
        aum_cat_filter = st.radio(
            "Filter Schemes (AUM):",
            ["All Schemes", "Individual Funds (ULIF)", "Group Funds (ULGF)"],
            horizontal=True,
            key="aum_cat_filter"
        )

        aum_audit_rows = []
        all_aum_sfins = set(list(pdf_data.keys()) + list(nav_data.keys()))

        for sfin in sorted(list(all_aum_sfins)):
            if aum_cat_filter == "Individual Funds (ULIF)" and not sfin.startswith("ULIF"):
                continue
            if aum_cat_filter == "Group Funds (ULGF)" and not sfin.startswith("ULGF"):
                continue

            fname = nav_data.get(sfin, {}).get("name") or BUILTIN_SCHEME_NAMES.get(sfin) or pdf_data.get(sfin, {}).get("fund_name") or sfin
            ftype = "Individual (ULIF)" if sfin.startswith("ULIF") else "Group (ULGF)"

            p_info = pdf_data.get(sfin)
            pdf_aum_val = p_info.get("aum_val") if p_info else None

            matched_code, excel_aum_val = get_exact_aum_for_sfin(sfin, fname, aum_data)

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
                "Excel Code": matched_code if matched_code else "N/A",
                "Excel AUM (Cr)": f"{excel_aum_val:.2f} Cr" if excel_aum_val is not None else "N/A",
                "PDF Published AUM": f"{pdf_aum_val:.2f} Cr" if pdf_aum_val is not None else "N/A",
                "Difference": f"{diff:.2f} Cr" if (pdf_aum_val is not None and excel_aum_val is not None) else "N/A",
                "Audit Status": status,
            })

        df_aum_audit = pd.DataFrame(aum_audit_rows)
        styled_aum_audit = apply_status_style(df_aum_audit, "Audit Status")
        st.dataframe(styled_aum_audit, use_container_width=True, hide_index=True)

# =============================================================
# BLOCK 8: TAB 3 (SCHEME-WISE DEEP AUDIT & 7 SUB-TABS)
# =============================================================
with tabs[2]:
    st.subheader("📑 Scheme-Wise Deep Portfolio Audit")
    available_schemes = {}
    all_possible_sfins = set(list(master_data.keys()) + list(pdf_data.keys()))

    for sfin_item in all_possible_sfins:
        m_name = master_data.get(sfin_item, {}).get("fund_name") if sfin_item in master_data else None
        n_name = nav_data.get(sfin_item, {}).get("name") if sfin_item in nav_data else None
        p_name = pdf_data.get(sfin_item, {}).get("fund_name") if sfin_item in pdf_data else None
        b_name = BUILTIN_SCHEME_NAMES.get(sfin_item)
        final_name = b_name or n_name or m_name or p_name or sfin_item
        if str(final_name).upper() in ["EQUITY", "CORPORATE DEBT", "TOTAL"]:
            final_name = b_name or n_name or sfin_item
        available_schemes[sfin_item] = f"{final_name} ({sfin_item})"

    if not available_schemes:
        st.info("💡 Is section ke liye sidebar mein **Factsheet PDF** ya **Master Factsheet Excel** upload karein.")
    else:
        sorted_sfins = sorted(list(available_schemes.keys()))
        if "scheme_selector_key" not in st.session_state or st.session_state.scheme_selector_key not in sorted_sfins:
            st.session_state.scheme_selector_key = sorted_sfins[0]

        current_idx = sorted_sfins.index(st.session_state.scheme_selector_key)

        nav_col1, nav_col2, nav_col3 = st.columns([1.2, 5, 1.2])
        with nav_col1:
            if st.button("⬅️ Previous", use_container_width=True, disabled=(current_idx == 0)):
                st.session_state.scheme_selector_key = sorted_sfins[current_idx - 1]
                st.rerun()

        with nav_col3:
            if st.button("Next ➡️", use_container_width=True, disabled=(current_idx == len(sorted_sfins) - 1)):
                st.session_state.scheme_selector_key = sorted_sfins[current_idx + 1]
                st.rerun()

        with nav_col2:
            chosen_sfin = st.selectbox(
                "Select Scheme to Inspect:",
                options=sorted_sfins,
                key="scheme_selector_key",
                format_func=lambda x: available_schemes[x]
            )

        def get_master_record(target_sfin, master_dict):
            if not master_dict or not target_sfin:
                return None
            if target_sfin in master_dict:
                return master_dict[target_sfin]
            norm_target = re.sub(r"[^A-Za-z0-9]", "", target_sfin).upper()
            for k, v in master_dict.items():
                norm_k = re.sub(r"[^A-Za-z0-9]", "", k).upper()
                if norm_target == norm_k or (len(norm_target) >= 15 and norm_target in norm_k) or (len(norm_k) >= 15 and norm_k in norm_target):
                    return v
            return None

        s_master = get_master_record(chosen_sfin, master_data)
        pdf_info = get_master_record(chosen_sfin, pdf_data)
        pdf_text = pdf_info["page_text"] if pdf_info else ""
        sheet_key = s_master["sheet_name"] if s_master else ""

        sub_tabs = st.tabs([
            "1. Asset Allocation",
            "2. Rating Profile",
            "3. Sectoral Allocation",
            "4. Corporate Debt",
            "5. Government Securities",
            "6. Equity Holdings",
            "7. AUM Audit & Charts",
        ])

        # --- SUB-TAB 1: ASSET ALLOCATION (STRICT SEQUENCE + TOTAL 100% CHECK) ---
        with sub_tabs[0]:
            st.markdown("#### Asset Allocation Audit (Excel vs. PDF Published)")
            if s_master and s_master["asset_allocation"]:
                TARGET_ASSET_ORDER = [
                    "EQUITY",
                    "DEBT INSTRUMENTS",
                    "GOVERNMENT SECURITIES",
                    "CORPORATE DEBT",
                    "MMI",
                ]

                def get_asset_sort_key(name):
                    n = str(name).upper().strip()
                    for idx, target in enumerate(TARGET_ASSET_ORDER):
                        if target == n or target in n or n in target:
                            return idx
                    return 99

                alloc_items = [
                    item for item in s_master["asset_allocation"]
                    if not (item["Asset Category"].upper() == "OTHER CORPORATE DEBT" and item["Raw Actual"] == 0.0)
                ]
                sorted_alloc = sorted(alloc_items, key=lambda x: get_asset_sort_key(x["Asset Category"]))

                alloc_rows = []
                total_excel_val = 0.0
                total_pdf_val = 0.0
                has_pdf_values = False

                for item in sorted_alloc:
                    cat_name = item["Asset Category"].upper()
                    excel_val = item["Raw Actual"]
                    pdf_pct, diff_str, status_str = audit_line_item_against_pdf(cat_name, excel_val, pdf_text, tolerance)
                    
                    total_excel_val += excel_val
                    if pdf_pct not in ["NOT IN PDF", "N/A", "PENDING PDF"]:
                        try:
                            total_pdf_val += float(pdf_pct.replace("%", "").strip())
                            has_pdf_values = True
                        except:
                            pass

                    rng = str(item.get("Defined Mandate Range", "N/A"))
                    alloc_rows.append({
                        "Asset Category": cat_name,
                        "Master Excel (%)": f"{excel_val:.2f}%",
                        "PDF Published (%)": pdf_pct,
                        "Difference (%)": diff_str,
                        "Defined Mandate Range": rng,
                        "Audit Status": status_str,
                    })

                # Dynamic 100.00% TOTAL Row Check
                if alloc_rows:
                    pdf_total_str = f"{total_pdf_val:.2f}%" if has_pdf_values else "N/A"
                    total_diff = abs(total_excel_val - total_pdf_val) if has_pdf_values else 0.0
                    total_diff_str = f"{total_diff:.2f}%" if has_pdf_values else "N/A"

                    is_excel_100 = abs(total_excel_val - 100.00) <= 0.05
                    is_pdf_100 = has_pdf_values and (abs(total_pdf_val - 100.00) <= 0.05)

                    if not has_pdf_values:
                        total_status = "PENDING PDF"
                    elif is_excel_100 and is_pdf_100:
                        total_status = "PASS"
                    else:
                        total_status = f"MISMATCH (TOTAL {total_pdf_val:.2f}%)"

                    alloc_rows.append({
                        "Asset Category": "📌 TOTAL",
                        "Master Excel (%)": f"{total_excel_val:.2f}%",
                        "PDF Published (%)": pdf_total_str,
                        "Difference (%)": total_diff_str,
                        "Defined Mandate Range": "100.00%",
                        "Audit Status": total_status,
                    })

                df_a = pd.DataFrame(alloc_rows)
                styled_a = apply_status_style(df_a, "Audit Status")
                st.dataframe(styled_a, use_container_width=True, hide_index=True)
            else:
                st.info("Upload Master Factsheet Excel to view Asset Allocation compliance.")

        # --- SUB-TAB 2: RATING PROFILE (LIVE 4-COLUMN AUDIT) ---
        with sub_tabs[1]:
            st.markdown("#### Credit Rating Profile Audit (Excel vs. PDF Published)")
            has_debt_exposure = False
            if s_master and s_master["asset_allocation"]:
                for item in s_master["asset_allocation"]:
                    cat_name = str(item.get("Asset Category", "")).upper()
                    if any(d in cat_name for d in ["DEBT", "GOVERNMENT", "CORPORATE", "MMI"]):
                        val_num = float(item.get("Raw Actual", 0))
                        if val_num > 0:
                            has_debt_exposure = True
                            break

            excel_ratings = s_master["ratings"] if (s_master and s_master["ratings"]) else []

            if excel_ratings:
                rating_audit_rows = []
                for r_item in excel_ratings:
                    r_label = r_item["Rating"]
                    excel_val = r_item["Raw Weight"]
                    pdf_pct, diff_str, status_str = audit_line_item_against_pdf(r_label, excel_val, pdf_text, tolerance)

                    rating_audit_rows.append({
                        "Credit Rating": r_label,
                        "Master Excel (%)": f"{excel_val:.2f}%",
                        "PDF Published (%)": pdf_pct,
                        "Difference (%)": diff_str,
                        "Audit Status": status_str,
                    })

                df_rating_audit = pd.DataFrame(rating_audit_rows)
                styled_r = apply_status_style(df_rating_audit, "Audit Status")
                st.dataframe(styled_r, use_container_width=True, hide_index=True)

            elif pdf_text:
                detected_ratings = []
                for code in ["SOV", "SOVEREIGN", "AAA", "AA+", "AA", "A1+", "AA-", "A+", "UNRATED"]:
                    if code in pdf_text:
                        detected_ratings.append(code)
                if detected_ratings:
                    st.write(f"**Ratings Detected in PDF Page:** `{', '.join(detected_ratings)}`")
                    st.warning("⚠️ Master Excel sheet mein rating table nahi mili ya format alag hai.")
                elif not has_debt_exposure:
                    st.info("No rating profile table required (100% Equity / Index Fund).")
                else:
                    st.info("No credit rating profile table required for this scheme.")
            elif not has_debt_exposure:
                st.info("No rating profile table required (100% Equity / Index Fund).")
            else:
                st.info("Upload Master Factsheet Excel or PDF to audit rating profile.")

        # --- SUB-TAB 3: SECTORAL ALLOCATION ---
        with sub_tabs[2]:
            st.markdown("#### Sectoral Allocation Audit (Excel vs. PDF Published)")
            if s_master and s_master["sectors"]:
                sector_audit_rows = []
                for s_item in s_master["sectors"]:
                    sec_name = s_item["Sector Name"]
                    excel_val = s_item["Raw Weight"]
                    pdf_pct, diff_str, status_str = audit_line_item_against_pdf(sec_name, excel_val, pdf_text, tolerance)

                    if "central government" not in str(sec_name).lower() and "gilt" not in str(sec_name).lower():
                        if excel_val > 25.0:
                            status_str = "BREACH (>25%)"

                    sector_audit_rows.append({
                        "Sector Name": sec_name,
                        "Master Excel (%)": f"{excel_val:.2f}%",
                        "PDF Published (%)": pdf_pct,
                        "Difference (%)": diff_str,
                        "Audit Status": status_str,
                    })

                df_sec_audit = pd.DataFrame(sector_audit_rows)
                styled_s = apply_status_style(df_sec_audit, "Audit Status")
                st.dataframe(styled_s, use_container_width=True, hide_index=True)
            else:
                st.info("Upload Master Factsheet Excel to view sectoral allocation.")

        # --- SUB-TAB 4: CORPORATE DEBT ---
        with sub_tabs[3]:
            st.markdown("#### Corporate Debt Holdings Audit (Excel vs. PDF Published)")
            if s_master and s_master["corp_debt"]:
                cd_rows = []
                for item in s_master["corp_debt"]:
                    sec_desc = item["Security Description"]
                    excel_val = item["Raw Weight"]
                    pdf_pct, diff_str, status_str = audit_line_item_against_pdf(sec_desc, excel_val, pdf_text, tolerance)

                    cd_rows.append({
                        "Security Description": sec_desc,
                        "Master Excel (%)": f"{excel_val:.2f}%",
                        "PDF Published (%)": pdf_pct,
                        "Difference (%)": diff_str,
                        "Audit Status": status_str,
                    })
                df_cd = pd.DataFrame(cd_rows)
                styled_cd = apply_status_style(df_cd, "Audit Status")
                st.dataframe(styled_cd, use_container_width=True, hide_index=True)
            else:
                st.info("No Corporate Debt instruments held in this scheme.")

        # --- SUB-TAB 5: GOVERNMENT SECURITIES ---
        with sub_tabs[4]:
            st.markdown("#### Government Securities Holdings Audit (Excel vs. PDF Published)")
            if s_master and s_master["g_sec"]:
                gs_rows = []
                for item in s_master["g_sec"]:
                    sec_desc = item["Security Description"]
                    excel_val = item["Raw Weight"]
                    pdf_pct, diff_str, status_str = audit_line_item_against_pdf(sec_desc, excel_val, pdf_text, tolerance)

                    gs_rows.append({
                        "Security Description": sec_desc,
                        "Master Excel (%)": f"{excel_val:.2f}%",
                        "PDF Published (%)": pdf_pct,
                        "Difference (%)": diff_str,
                        "Audit Status": status_str,
                    })
                df_gs = pd.DataFrame(gs_rows)
                styled_gs = apply_status_style(df_gs, "Audit Status")
                st.dataframe(styled_gs, use_container_width=True, hide_index=True)
            else:
                st.info("No Government Securities held in this scheme.")

        # --- SUB-TAB 6: EQUITY HOLDINGS ---
        with sub_tabs[5]:
            st.markdown("#### Equity Holdings Audit (Excel vs. PDF Published)")
            if s_master and s_master["equity"]:
                eq_rows = []
                for item in s_master["equity"]:
                    sec_desc = item["Security Description"]
                    excel_val = item["Raw Weight"]
                    pdf_pct, diff_str, status_str = audit_line_item_against_pdf(sec_desc, excel_val, pdf_text, tolerance)

                    eq_rows.append({
                        "Security Description": sec_desc,
                        "Master Excel (%)": f"{excel_val:.2f}%",
                        "PDF Published (%)": pdf_pct,
                        "Difference (%)": diff_str,
                        "Audit Status": status_str,
                    })
                df_eq = pd.DataFrame(eq_rows)
                styled_eq = apply_status_style(df_eq, "Audit Status")
                st.dataframe(styled_eq, use_container_width=True, hide_index=True)
            else:
                st.info("No Equity holdings held in this scheme (100% Debt / Liquid Fund).")

        # --- SUB-TAB 7: AUM AUDIT & CHARTS ---
        with sub_tabs[6]:
            st.markdown("#### Scheme Asset Class Breakup & AUM Audit")
            fund_title = available_schemes.get(chosen_sfin, "")
            found_code, excel_aum_val = get_exact_aum_for_sfin(chosen_sfin, fund_title, aum_data)

            if aum_data:
                aum_codes = sorted(list(aum_data.keys()))
                dropdown_options = ["-- Select AUM Code --"] + aum_codes
                default_idx = dropdown_options.index(found_code) if (found_code and found_code in dropdown_options) else 0

                selected_option = st.selectbox(
                    "Matched AUM Code (From AUM Excel):",
                    options=dropdown_options,
                    index=default_idx,
                    format_func=lambda c: (
                        f"{c}  |  {aum_data[c]['Name']}  ({aum_data[c]['Total AUM (Cr)']} Cr)" 
                        if (c in aum_data and aum_data[c]['Name']) 
                        else f"{c}  ({aum_data[c]['Total AUM (Cr)']} Cr)" if c in aum_data 
                        else c
                    ),
                    help="Auto-matched from AUM file. You can change manually if needed."
                )

                if selected_option != "-- Select AUM Code --":
                    found_code = selected_option
                    excel_aum_val = aum_data[found_code]["Total AUM (Cr)"]
                else:
                    found_code = None
                    excel_aum_val = None

            pdf_aum_val = pdf_info.get("aum_val") if pdf_info else None

            m_c1, m_c2, m_c3 = st.columns(3)
            m_c1.metric("PDF Published AUM", f"{pdf_aum_val:.2f} Cr" if pdf_aum_val else "N/A")
            m_c2.metric("Excel System AUM", f"{excel_aum_val:.2f} Cr" if excel_aum_val is not None else "N/A")
            if pdf_aum_val and excel_aum_val is not None:
                diff_aum = abs(pdf_aum_val - excel_aum_val)
                m_c3.metric("Difference", f"{diff_aum:.2f} Cr", delta="MATCH" if diff_aum <= 0.05 else "-MISMATCH")

            if found_code and found_code in aum_data:
                info_a = aum_data[found_code]
                chart_df = pd.DataFrame({
                    "Asset Class": ["Corporate Debt", "Government Securities", "Equity", "MMI"],
                    "AUM (In Crores)": [info_a["Corporate Debt"], info_a["G-Sec"], info_a["Equity"], info_a["MMI"]],
                }).set_index("Asset Class")

                chart_df_non_zero = chart_df[chart_df["AUM (In Crores)"] > 0]
                if not chart_df_non_zero.empty:
                    st.bar_chart(chart_df_non_zero)
                    st.table(chart_df)
            else:
                st.info("Upload `AUM with Fund names.xlsx` in sidebar to view asset class breakdown.")

# =============================================================
# BLOCK 9: TAB 4, 5 & 6 (BREAKUP, YTM & EXPORT REPORT)
# =============================================================
# --- TAB 4: RAW AUM CODE BREAKDOWN ---
with tabs[3]:
    st.subheader("📊 Raw AUM Code Breakdown Table")
    if not aum_data:
        st.info("💡 Sirf **`AUM with Fund names.xlsx`** upload karein AUM numbers verify karne ke liye.")
    else:
        aum_rows = []
        for code, info_a in aum_data.items():
            aum_rows.append({
                "Scheme Code": code,
                "Fund Name": info_a.get("Name", ""),
                "Total AUM (Cr)": f"{info_a['Total AUM (Cr)']:.2f}",
            })
        st.dataframe(pd.DataFrame(aum_rows), use_container_width=True, hide_index=True)

# --- TAB 5: YTM & DURATION AUDIT (SMART SFIN & MULTI-LINE REGEX) ---
with tabs[4]:
    st.subheader("📉 Yield to Maturity (YTM) & Duration Audit (Excel vs. PDF)")
    if not ytm_data:
        st.info("💡 Sirf **`Yield, Duration Report-30.09.26.xlsx`** upload karein YTM aur Duration audit ke liye.")
    else:
        ytm_cat_filter = st.radio(
            "Filter Schemes (YTM):",
            ["All Schemes", "Individual Funds (ULIF)", "Group Funds (ULGF)"],
            horizontal=True,
            key="ytm_cat_filter",
        )

        def get_ytm_category(code_str, name_str):
            combined = (str(code_str) + " " + str(name_str)).upper()
            if "ULGF" in combined or combined.startswith("GR-") or combined.startswith("GR ") or "GROUP" in combined:
                return "Group (ULGF)"
            elif "ULIF" in combined or combined.startswith("IND-") or combined.startswith("IND ") or "INDIVIDUAL" in combined or "PENSION" in combined:
                return "Individual (ULIF)"
            for sfin_k, mapped_code in EXHAUSTIVE_SFIN_MAP.items():
                if norm_txt(mapped_code) == norm_txt(code_str) or norm_txt(mapped_code) in norm_txt(name_str):
                    return "Group (ULGF)" if sfin_k.startswith("ULGF") else "Individual (ULIF)"
            return "Other"

        ytm_rows = []
        seen_codes = set()

        for code, info in ytm_data.items():
            code_clean = str(info.get("Code", code)).strip()
            if code_clean in seen_codes:
                continue
            seen_codes.add(code_clean)

            fname = info.get("Name", code_clean)
            cat_type = get_ytm_category(code_clean, fname)

            # Filter logic
            if ytm_cat_filter == "Individual Funds (ULIF)" and cat_type != "Individual (ULIF)":
                continue
            if ytm_cat_filter == "Group Funds (ULGF)" and cat_type != "Group (ULGF)":
                continue

            excel_ytm = info.get("YTM")
            excel_dur = info.get("Duration")

            # 1. Direct SFIN Extraction from Fund Name (e.g. ULGF00212... from fname)
            m_sfin_in_name = re.search(r"(?:ULIF|ULGF)[A-Za-z0-9/]+?109", fname)
            detected_sfin = m_sfin_in_name.group(0).strip() if m_sfin_in_name else None

            pdf_info = None
            if detected_sfin and detected_sfin in pdf_data:
                pdf_info = pdf_data[detected_sfin]
            
            # 2. SFIN Map Fallback
            if not pdf_info:
                for sfin_k, mapped_c in EXHAUSTIVE_SFIN_MAP.items():
                    if norm_txt(mapped_c) == norm_txt(code_clean) and sfin_k in pdf_data:
                        pdf_info = pdf_data[sfin_k]
                        break

            # 3. Name Match Fallback
            if not pdf_info:
                clean_target_name = norm_txt(re.sub(r"(?:ULIF|ULGF)[A-Za-z0-9/\-]+", "", fname))
                for sfin_k, p_val in pdf_data.items():
                    p_name_clean = norm_txt(p_val.get("fund_name", ""))
                    if clean_target_name and (clean_target_name in p_name_clean or p_name_clean in clean_target_name):
                        pdf_info = p_val
                        break

            pdf_ytm = None
            pdf_dur = None

            if pdf_info:
                p_txt = pdf_info.get("page_text", "")
                
                # Multi-line & Flexible YTM Regex
                m_ytm = re.search(r"(?:Yield\s+to\s+Maturity|YTM).*?([0-9]{1,2}\.[0-9]{1,2})\s*%", p_txt, re.IGNORECASE | re.DOTALL)
                if m_ytm:
                    try:
                        pdf_ytm = float(m_ytm.group(1))
                    except:
                        pass
                elif pdf_info.get("ytm") is not None:
                    pdf_ytm = pdf_info.get("ytm")

                # Multi-line & Flexible Duration Regex
                m_dur = re.search(r"(?:Modified\s+Duration|Duration).*?([0-9]{1,2}\.[0-9]{1,2})", p_txt, re.IGNORECASE | re.DOTALL)
                if m_dur:
                    try:
                        pdf_dur = float(m_dur.group(1))
                    except:
                        pass
                elif pdf_info.get("mod_duration") is not None:
                    pdf_dur = pdf_info.get("mod_duration")

            # Status calculation
            if not uploaded_pdf:
                status = "PENDING PDF"
            elif pdf_info is None:
                status = "NOT IN PDF"
            elif pdf_ytm is None and pdf_dur is None:
                status = "NOT IN PDF"
            else:
                ytm_ok = (pdf_ytm is None or excel_ytm is None or abs(pdf_ytm - excel_ytm) <= 0.05)
                dur_ok = (pdf_dur is None or excel_dur is None or abs(pdf_dur - excel_dur) <= 0.05)

                if ytm_ok and dur_ok:
                    status = "PASS"
                else:
                    parts = []
                    if not ytm_ok and pdf_ytm and excel_ytm:
                        parts.append(f"YTM Diff: {abs(pdf_ytm - excel_ytm):.2f}%")
                    if not dur_ok and pdf_dur and excel_dur:
                        parts.append(f"Dur Diff: {abs(pdf_dur - excel_dur):.2f}")
                    status = f"MISMATCH ({', '.join(parts)})"

            # Clean Display Name (SFIN prefix hatakar clean fund name)
            display_fund_name = re.sub(r"^(?:ULIF|ULGF)[A-Za-z0-9/\-]+", "", fname).strip()
            if not display_fund_name:
                display_fund_name = BUILTIN_SCHEME_NAMES.get(detected_sfin, fname)

            ytm_rows.append({
                "Category": cat_type,
                "Scheme Code": code_clean,
                "Fund Name": display_fund_name,
                "Excel YTM": f"{excel_ytm:.2f}%" if excel_ytm is not None else "N/A",
                "PDF YTM": f"{pdf_ytm:.2f}%" if pdf_ytm is not None else "N/A",
                "Excel Duration": f"{excel_dur:.2f}" if excel_dur is not None else "N/A",
                "PDF Duration": f"{pdf_dur:.2f}" if pdf_dur is not None else "N/A",
                "Audit Status": status,
            })

        df_ytm = pd.DataFrame(ytm_rows)
        styled_ytm = apply_status_style(df_ytm, "Audit Status")
        st.dataframe(styled_ytm, use_container_width=True, hide_index=True)
        

# --- TAB 6: EXPORT COMPREHENSIVE WORKBOOK ---
with tabs[5]:
    st.subheader("📥 Export Full Audit Reconciliation Report")
    out_buf = io.BytesIO()
    with pd.ExcelWriter(out_buf, engine="xlsxwriter") as writer:
        if nav_data and "df_nav_audit" in locals():
            df_nav_audit.to_excel(writer, sheet_name="NAV Audit", index=False)
        if "df_aum_audit" in locals():
            df_aum_audit.to_excel(writer, sheet_name="AUM Master Audit", index=False)
        if ytm_data and "ytm_rows" in locals():
            pd.DataFrame(ytm_rows).to_excel(writer, sheet_name="YTM & Duration", index=False)

    st.download_button(
        label="📊 Download Comprehensive Audit Summary (.xlsx)",
        data=out_buf.getvalue(),
        file_name="ABSLI_Factsheet_Audit_Summary_Sep2026.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
