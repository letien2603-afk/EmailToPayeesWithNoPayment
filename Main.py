import streamlit as st
import pandas as pd
import smtplib
from datetime import datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# Set up Streamlit Page Configuration
st.set_page_config(
    page_title="No SIP Payout Notification - Email Automator",
    page_icon="📧",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Helpers for parsing and formatting
def parse_value(val):
    """Safely parse percentage or currency to float."""
    if pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    val_str = str(val).strip().replace('$', '').replace(',', '')
    if val_str.endswith('%'):
        try:
            return float(val_str.rstrip('%')) / 100.0
        except ValueError:
            return 0.0
    try:
        return float(val_str)
    except ValueError:
        return 0.0

def parse_pct_value(val):
    """Safely parse percentage to decimal float (e.g. '20%' -> 0.20, 20.0 -> 0.20 if > 1.0)."""
    parsed = parse_value(val)
    return (parsed / 100.0) if parsed > 1.0 else parsed

def format_pct(val):
    """Format float as percentage (e.g. 0.45 -> '45.00%')."""
    return f"{parse_value(val) * 100:.2f}%"

def format_curr(val):
    """Format float as standard currency with 2 decimals (e.g. 15000.0 -> '15,000.00')."""
    return f"{parse_value(val):,.2f}"

def adjust_pay_date_for_currency(pay_date_val, currency_code):
    """If Currency Code is CAD, advance pay_date by 1 week (7 days)."""
    if not pay_date_val or str(currency_code).strip().upper() != 'CAD':
        return str(pay_date_val).strip() if pay_date_val else ""
    
    clean_date = str(pay_date_val).strip()
    for fmt in ('%m/%d/%Y', '%Y-%m-%d', '%m-%d-%Y', '%d/%m/%Y', '%Y/%m/%d'):
        try:
            dt = datetime.strptime(clean_date, fmt)
            return (dt + timedelta(days=7)).strftime(fmt)
        except ValueError:
            pass
    try:
        dt = pd.to_datetime(clean_date)
        return (dt + timedelta(days=7)).strftime('%m/%d/%Y' if '/' in clean_date else '%Y-%m-%d')
    except Exception:
        return clean_date

def find_column(df, aliases):
    """Find a column in a dataframe matching any of the specified aliases (case-insensitive)."""
    for alias in aliases:
        for col in df.columns:
            if str(col).strip().lower() == alias.strip().lower():
                return col
    return None

def get_dict_val(p_name, val_dict, default=0.0):
    """Safely lookup value by participant name, case-insensitively."""
    if not val_dict or pd.isna(p_name):
        return default
    p_str = str(p_name).strip()
    if p_str in val_dict:
        return val_dict[p_str]
    p_lower = p_str.lower()
    for k, v in val_dict.items():
        if str(k).strip().lower() == p_lower:
            return v
    return default

def extract_target_payees_from_excel(target_file):
    """Extract list of target payee names from an uploaded Excel file (.xlsx or .xls)."""
    try:
        xls = pd.ExcelFile(target_file)
        names = []
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name)
            df.columns = [str(c).strip() for c in df.columns]
            target_col = find_column(df, [
                "*Participant", "Participant", "Payee", "Payee Name", 
                "Employee", "Employee Name", "Name", "Participant Name", "PayeeName"
            ])
            if target_col is not None:
                sheet_names_list = df[target_col].dropna().astype(str).str.strip().tolist()
                names.extend([n for n in sheet_names_list if n and n.lower() != 'nan'])
            elif len(df.columns) >= 1:
                first_col_names = df.iloc[:, 0].dropna().astype(str).str.strip().tolist()
                names.extend([n for n in first_col_names if n and n.lower() != 'nan'])
        # Deduplicate preserving order
        seen = set()
        deduped = []
        for n in names:
            if n.lower() not in seen:
                seen.add(n.lower())
                deduped.append(n)
        return deduped
    except Exception as e:
        st.error(f"Error reading target payees Excel file: {e}")
        return []

# Dialog modal popup for missing credentials
@st.dialog("Configuration Error")
def show_credentials_error_dialog():
    st.error("Please fill in both the Sender Email and Gmail App Password in the sidebar before sending!")
    if st.button("OK", use_container_width=True):
        st.rerun()

def generate_email_html(participant, is_both_goals, m, pay_date, sip_month, has_old_role=False, old_role_sip_earned=0.0, total_ytd_earnings=0.0):
    """Generate professional Excel-style HTML table for email body matching Both Goals or Only 1 Goal layout."""
    month_names = {3: "March", 6: "June", 9: "September", 12: "December"}
    month_name = month_names.get(sip_month, "Current Month")
    
    total_earned = m.get('total_sip_earned', 0.0)
    total_paid = m.get('total_sip_paid', 0.0)
    total_ytd_sip_earned = total_earned + old_role_sip_earned
    earned_disp_str = format_curr(total_ytd_earnings if has_old_role and total_ytd_earnings > 0 else total_earned)

    intro_text = f"For your information, you will not receive any {month_name} SIP payment in the upcoming pay cycle on <strong>{pay_date}</strong> because your YTD SIP Earned is {earned_disp_str} while your YTD SIP Paid is {format_curr(total_paid)}."

    colgroup_html = """
        <colgroup>
          <col style="width: 275px;">
          <col style="width: 110px;">
        </colgroup>"""

    if is_both_goals:
        base_rows = [
            ("Goal 1 Annual SIP", format_curr(m['goal1_annual_sip']), ""),
            ("Goal 1 Annual Quota", format_curr(m['goal1_annual_quota']), ""),
            (f"Goal 1 YTD Revenue (Up to {month_name})", format_curr(m['goal1_ytd_rev']), ""),
            (f"Goal 1 FY Attainment (Up to {month_name})", format_pct(m['goal1_fy_attainment']), ""),
            ("Goal 1 SIP Earned", format_curr(m['goal1_sip_earned']), ""),
            ("Goal 2 Annual SIP", format_curr(m['goal2_annual_sip']), ""),
            ("Goal 2 Annual Quota", format_curr(m['goal2_annual_quota']), ""),
            (f"Goal 2 YTD Revenue (Up to {month_name})", format_curr(m['goal2_ytd_rev']), ""),
            (f"Goal 2 FY Attainment (Up to {month_name})", format_pct(m['goal2_fy_attainment']), ""),
            ("Goal 2 SIP Earned", format_curr(m['goal2_sip_earned']), "")
        ]
    else:
        base_rows = [
            ("Annual SIP", format_curr(m['annual_sip']), ""),
            ("Annual Quota", format_curr(m['annual_quota']), ""),
            (f"YTD Revenue (Up to {month_name})", format_curr(m['ytd_rev']), ""),
            (f"FY Attainment (Up to {month_name})", format_pct(m['fy_attainment']), "")
        ]

    def render_rows(rows):
        res = ""
        for i, (label, val_str, *css) in enumerate(rows):
            cls = f' class="{css[0]}"' if css and css[0] else ""
            w_left = ' style="width: 275px;"' if i == 0 else ""
            w_right = ' style="width: 110px;"' if i == 0 else ""
            res += f"""
        <tr{cls}>
          <td class="left-align"{w_left}>{label}</td>
          <td class="right-align"{w_right}>{val_str}</td>
        </tr>"""
        return res

    if has_old_role:
        new_role_rows = base_rows + [("Total SIP Earned", format_curr(total_earned), "")]
        tables_content_html = f"""
      <p class="bold underline" style="margin-top: 15px; margin-bottom: 6px;">Old Role</p>
      <table class="paysheet-table">
        {colgroup_html}
        <tr>
          <td class="left-align" style="width: 275px;">SIP Earned</td>
          <td class="right-align" style="width: 110px;">{format_curr(old_role_sip_earned)}</td>
        </tr>
      </table>
      
      <p class="bold underline" style="margin-top: 18px; margin-bottom: 6px;">New Role</p>
      <table class="paysheet-table">
        {colgroup_html}
        {render_rows(new_role_rows)}
      </table>

      <!-- Summary table for Total YTD SIP Earned and Total SIP Paid -->
      <table class="paysheet-table" style="margin-top: 15px; margin-bottom: 15px;">
        {colgroup_html}
        <tr class="bg-earned">
          <td class="left-align" style="width: 275px;">Total YTD SIP Earned</td>
          <td class="right-align" style="width: 110px;">{format_curr(total_ytd_sip_earned)}</td>
        </tr>
        <tr class="bg-paid">
          <td class="left-align">Total SIP Paid</td>
          <td class="right-align">{format_curr(total_paid)}</td>
        </tr>
      </table>
        """
    else:
        table_rows = base_rows + [
            ("Total SIP Earned", format_curr(total_earned), "bg-earned"),
            ("Total SIP Paid", format_curr(total_paid), "bg-paid")
        ]
        tables_content_html = f"""
      <table class="paysheet-table" style="margin-top: 15px; margin-bottom: 15px;">
        {colgroup_html}
        {render_rows(table_rows)}
      </table>
        """

    return f"""
    <html>
    <head>
    <style>
      body {{
        font-family: Calibri, Arial, sans-serif;
        font-size: 11pt;
        color: #000000;
        line-height: 1.5;
      }}
      p {{
        margin-bottom: 16px;
      }}
      table.paysheet-table {{
        border-collapse: collapse;
        font-family: Calibri, Arial, sans-serif;
        font-size: 11pt;
        color: #000000;
        background-color: #ffffff;
        width: 385px;
        table-layout: fixed;
        margin-top: 5px;
        margin-bottom: 15px;
      }}
      table.paysheet-table td {{
        border: 1px solid #d3d3d3;
        padding: 4px 8px;
        vertical-align: middle;
        overflow: hidden;
        word-wrap: break-word;
      }}
      .bold {{
        font-weight: bold;
      }}
      .underline {{
        text-decoration: underline;
      }}
      .left-align {{
        text-align: left;
      }}
      .right-align {{
        text-align: right;
      }}
      .bg-earned {{
        background-color: #d9d9d9;
        font-weight: bold;
      }}
      .bg-paid {{
        background-color: #c4bd97;
        font-weight: bold;
      }}
    </style>
    </head>
    <body>
      <p>Hi {participant},</p>
      <p>{intro_text}</p>
      <p>Please see the YTD payment breakdown below for more information:</p>
      {tables_content_html}
      <p>Regards,<br>Sales Compensation</p>
    </body>
    </html>
    """

def is_file_for_participant(filename, participant_name):
    """Check if the uploaded attachment file belongs to the participant."""
    fn = filename.lower()
    pn = participant_name.lower().strip()
    pn_variations = [pn, pn.replace(" ", "_"), pn.replace(" ", "-"), pn.replace(" ", "")]
    return any(var in fn for var in pn_variations if var)

def get_participant_attachments(participant_name, uploaded_files):
    """Retrieve all uploaded files that match the participant's name."""
    if not uploaded_files:
        return []
    return [f for f in uploaded_files if is_file_for_participant(f.name, participant_name)]

# App UI Design
st.title("📧 No SIP Payout Notification - Email Automator")

# Sidebar Configuration
st.sidebar.header("⚙️ Fill in required information")

sender_email = st.sidebar.text_input("Sender Gmail:", placeholder="your_email@gmail.com")
password = st.sidebar.text_input("Gmail App Password (16 characters):", type="password", placeholder="xxxx xxxx xxxx xxxx",
           help="Settings -> Accounts -> Google Account settings -> Security & Sign-in -> 2-Step Verification -> App passwords -> Add the App name, then click Create")
cc_input = st.sidebar.text_input("CC Emails (Separated by commas):", placeholder="e.g. manager@example.com, hr@example.com")

st.sidebar.markdown("---")
st.sidebar.header("✉️ Email Subject")
email_subject_template = st.sidebar.text_input(
    "Email Subject:", 
    placeholder="e.g. June No SIP Payout Notification"
)

st.sidebar.markdown("---")
st.sidebar.header("🗓️ Payroll Cycle Information")
pay_date = st.sidebar.text_input("Pay Date:", placeholder="e.g. 08/31/2026")
sip_month = st.sidebar.selectbox("SIP month to pay:", options=["", 3, 6, 9, 12])

# File uploaders on main screen
st.subheader("📁 Step 1: Upload Payfile")
uploaded_file = st.file_uploader("Choose the Payfile file (.xlsx)")

st.subheader("🎯 Step 2: Upload Target Payees List (Optional)")
uploaded_target_payees_file = st.file_uploader(
    "Only send email notification to payees in the uploaded file:",
    type=["xlsx", "xls"],
    key="target_payees_uploader"
)

st.subheader("📎 Step 3: Upload Email Attachments (Optional)")
uploaded_attachments = st.file_uploader(
    "Upload files to attach to participant emails:",
    accept_multiple_files=True
)

# Parse global CC list
cc_list = [e.strip() for e in cc_input.split(",") if e.strip()]

if uploaded_file is not None:
    st.success("File uploaded successfully! Processing data...")
    
    try:
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names
        
        # Auto-detect sheet names
        paysheet_name = None
        for name in ["Paysheet", "Pay Sheet", "PaySheet", "paysheet", "OIC Pay Sheet"]:
            if name in sheet_names:
                paysheet_name = name
                break
                
        oic_sheet_name = None
        for name in ["OIC Participant Details", "Participant Details", "OIC Participants", "OIC Participant Details "]:
            if name in sheet_names:
                oic_sheet_name = name
                break

        fx_sheet_name = None
        for name in ["FX Rate", "FX Rates", "FXRate", "Fx Rate", "fx rate", "FX_Rate", "FX rate"]:
            if name in sheet_names:
                fx_sheet_name = name
                break
                
        if paysheet_name is None:
            st.error(f"Could not find 'Paysheet' or 'Pay Sheet' sheet in the file. Available sheets: {sheet_names}")
        elif oic_sheet_name is None:
            st.error(f"Could not find 'OIC Participant Details' sheet in the file. Available sheets: {sheet_names}")
        else:
            paysheet_df = pd.read_excel(xls, sheet_name=paysheet_name)
            oic_df = pd.read_excel(xls, sheet_name=oic_sheet_name)
            
            paysheet_df.columns = [str(c).strip() for c in paysheet_df.columns]
            oic_df.columns = [str(c).strip() for c in oic_df.columns]
            
            participant_col_pay = find_column(paysheet_df, ["*Participant", "Participant"])
            participant_col_oic = find_column(oic_df, ["*Participant", "Participant"])
            email_col_oic = find_column(oic_df, ["E-Mail Address", "Email Address", "Email", "E-Mail"])
            manager_email_col_oic = find_column(oic_df, [
                "Managers Email Address", "Manager's Email Address", 
                "Manager Email Address", "Manager Email", "Managers Email"
            ])
            annual_quota_col_oic = find_column(oic_df, ["Annual Quota", "Annual Quota 1", "Goal 1 Annual Quota"])
            annual_quota2_col_oic = find_column(oic_df, ["Annual Quota2", "Annual Quota 2", "Goal 2 Annual Quota"])
            goal1_weight_col_oic = find_column(oic_df, ["Goal1 Weightage", "Goal 1 Weightage", "Goal1 Weight", "Goal 1 Weight", "Goal 1 Weight%"])
            goal2_weight_col_oic = find_column(oic_df, ["Goal2 Weightage", "Goal 2 Weightage", "Goal2 Weight", "Goal 2 Weight", "Goal 2 Weight%"])

            if not participant_col_pay or not participant_col_oic or not email_col_oic:
                st.error("Excel file structure mismatch! Please check '*Participant' and 'E-Mail Address' columns.")
            else:
                # Build OIC lookup maps
                email_map = oic_df.set_index(participant_col_oic)[email_col_oic].to_dict()
                paysheet_df['Email'] = paysheet_df[participant_col_pay].map(email_map)
                
                manager_email_map = oic_df.set_index(participant_col_oic)[manager_email_col_oic].to_dict() if manager_email_col_oic else {}
                annual_quota_map = oic_df.set_index(participant_col_oic)[annual_quota_col_oic].to_dict() if annual_quota_col_oic else {}
                annual_quota2_map = oic_df.set_index(participant_col_oic)[annual_quota2_col_oic].to_dict() if annual_quota2_col_oic else {}
                goal1_weight_map = oic_df.set_index(participant_col_oic)[goal1_weight_col_oic].to_dict() if goal1_weight_col_oic else {}
                goal2_weight_map = oic_df.set_index(participant_col_oic)[goal2_weight_col_oic].to_dict() if goal2_weight_col_oic else {}

                # Build FX conversion map
                fx_map = {"USD": 1.0}
                if fx_sheet_name is not None:
                    fx_df = pd.read_excel(xls, sheet_name=fx_sheet_name)
                    fx_df.columns = [str(c).strip() for c in fx_df.columns]
                    payment_curr_col = find_column(fx_df, ["Payment Currency", "PaymentCurrency", "Currency", "Currency Code"])
                    convert_usd_col = find_column(fx_df, ["Convert to USD", "ConvertToUSD", "Convert To USD", "USD Conversion Rate", "Conversion Rate", "FX Rate", "Rate"])
                    if payment_curr_col and convert_usd_col:
                        for _, fx_row in fx_df.iterrows():
                            p_curr = fx_row[payment_curr_col]
                            c_rate = fx_row[convert_usd_col]
                            if pd.notna(p_curr) and str(p_curr).strip():
                                key = str(p_curr).strip().upper()
                                try:
                                    fx_map[key] = float(parse_value(c_rate))
                                except Exception:
                                    fx_map[key] = 1.0

                currency_code_col_pay = find_column(paysheet_df, ["Currency Code", "Currency", "CurrencyCode", "Payment Currency"])
                def get_conversion_rate(val):
                    if pd.isna(val):
                        return 1.0
                    val_str = str(val).strip().upper()
                    if val_str == "USD" or not val_str or val_str == "NAN":
                        return 1.0
                    return fx_map.get(val_str, 1.0)

                paysheet_df['Conversion Rate'] = paysheet_df[currency_code_col_pay].apply(get_conversion_rate) if currency_code_col_pay else 1.0

                # Pay Sheet column identifiers
                salary_col = find_column(paysheet_df, ["Salary (USD)", "Salary", "Base Salary (USD)", "Base Salary"])
                target_inc_col = find_column(paysheet_df, ["Target Incentive%", "Target Incentive %", "Target Incentive", "Target Incentive Pct"])
                goal1_quota_col_pay = find_column(paysheet_df, ["Goal 1 Quota", "Goal 1 Annual Quota", "Goal1 Quota", "Goal1 Annual Quota"])
                goal2_quota_col_pay = find_column(paysheet_df, ["Goal 2 Quota", "Goal 2 Annual Quota", "Goal2 Quota", "Goal2 Annual Quota"])
                ytd_rev_col = find_column(paysheet_df, ["YTD REV", "YTD Rev", "YTD Revenue", "Goal 1 YTD REV", "Goal 1 YTD Revenue"])
                ytd_territory_rev_col = find_column(paysheet_df, ["YTD TERRITORY REV", "YTD Territory REV", "YTD Territory Rev", "YTD Territory Revenue", "Territory YTD REV"])
                total_ytd_earnings_col = find_column(paysheet_df, [
                    "Total YTD Earnings", "Total YTD Earning", "Total YTD SIP Earnings", "Total YTD SIP Earned",
                    "YTD Earnings", "YTD Earning", "YTD SIP Earnings", "YTD SIP Earned"
                ])
                
                # Find Total SIP Paid column ("YTD Paid (Month...")
                total_paid_col = None
                for col in paysheet_df.columns:
                    if "ytd paid (month" in str(col).lower():
                        total_paid_col = col
                        break
                if total_paid_col is None:
                    total_paid_col = find_column(paysheet_df, ["Total SIP Paid", "YTD SIP Paid", "YTD Paid"])

                # Filtering: Include in Payfile == 1 and Net Commissions < 0
                include_col = find_column(paysheet_df, ["Include in This Month's PayFile? (1=Yes, 0=No)", "Include in This Month's PayFile?", "Include in PayFile"])
                net_comm_col = find_column(paysheet_df, ["Net Commissions Payable", "Net Commission Payable"])
                
                if not include_col or not net_comm_col:
                    st.error("Filtering condition columns 'Include in This Month's PayFile?' or 'Net Commissions Payable' not found!")
                else:
                    filtered_df = paysheet_df[
                        (paysheet_df[include_col] == 1) & 
                        (paysheet_df[net_comm_col].apply(parse_value) < 0)
                    ]
                    
                    st.subheader("📊 Step 4: Filtering Results Check")
                    
                    # Target Payees List Filter (if uploaded)
                    if uploaded_target_payees_file is not None:
                        target_payees = extract_target_payees_from_excel(uploaded_target_payees_file)
                        if target_payees:
                            target_payees_lower = [t.lower() for t in target_payees]
                            raw_before = len(filtered_df)
                            filtered_df = filtered_df[filtered_df[participant_col_pay].astype(str).str.strip().str.lower().isin(target_payees_lower)]
                            kept_count = len(filtered_df)
                            st.success(f"🎯 **Target Payees File Applied:** Kept **{kept_count}** payee(s) from target file (excluded {raw_before - kept_count} non-target payees).")
                        else:
                            st.warning("⚠️ The uploaded target payees file contained no recognizable participant names.")

                    # Check for missing email addresses
                    missing_email_df = filtered_df[filtered_df['Email'].isna() | (filtered_df['Email'].astype(str).str.strip() == '')]
                    if len(missing_email_df) > 0:
                        missing_names = missing_email_df[participant_col_pay].astype(str).str.strip().tolist()
                        st.warning(f"⚠️ **Warning: Missing Email Address** for **{len(missing_email_df)}** participant(s): `{', '.join(missing_names)}`. They have been excluded from email sending.")
                        filtered_df = filtered_df[~(filtered_df['Email'].isna() | (filtered_df['Email'].astype(str).str.strip() == ''))]

                    total_matches = len(filtered_df)
                    st.success(f"Found **{total_matches}** active participant(s) for the email list.")

                    if total_matches > 0:
                        email_records = []
                        preview_table_data = []

                        for idx, row in filtered_df.iterrows():
                            participant = str(row[participant_col_pay]).strip()
                            recipient_email = str(row['Email']).strip()
                            conversion_rate = float(row.get('Conversion Rate', 1.0))
                            
                            salary = parse_value(row[salary_col]) if salary_col else 0.0
                            target_incentive = parse_pct_value(row[target_inc_col]) if target_inc_col else 0.0
                            total_paid = parse_value(row[total_paid_col]) if total_paid_col else 0.0
                            total_ytd_earnings = parse_value(row[total_ytd_earnings_col]) if total_ytd_earnings_col else 0.0

                            # Determine currency and payee-specific pay date (CAD = +1 week)
                            currency_code = str(row[currency_code_col_pay]).strip().upper() if currency_code_col_pay and pd.notna(row[currency_code_col_pay]) else "USD"
                            payee_pay_date = adjust_pay_date_for_currency(pay_date, currency_code)

                            # Determine whether payee has Annual Quota2 (>0 and not blank)
                            quota2_val_raw = get_dict_val(participant, annual_quota2_map, default=0.0)
                            quota2_parsed = parse_value(quota2_val_raw)
                            has_both_goals = (pd.notna(quota2_val_raw) and str(quota2_val_raw).strip() != "" and quota2_parsed > 0.0)

                            m = {}
                            if has_both_goals:
                                # Both Goals logic
                                g1_weight = parse_pct_value(get_dict_val(participant, goal1_weight_map, default=0.0))
                                g2_weight = parse_pct_value(get_dict_val(participant, goal2_weight_map, default=0.0))
                                
                                g1_annual_sip = salary * target_incentive * g1_weight * conversion_rate
                                # Goal 1 Annual Quota: taken from Goal 1 Quota in Pay Sheet
                                if goal1_quota_col_pay and pd.notna(row[goal1_quota_col_pay]):
                                    g1_annual_quota = parse_value(row[goal1_quota_col_pay])
                                else:
                                    g1_annual_quota = parse_value(get_dict_val(participant, annual_quota_map, default=0.0))
                                    
                                g1_ytd_rev = parse_value(row[ytd_rev_col]) if ytd_rev_col else 0.0
                                g1_fy_att = (g1_ytd_rev / g1_annual_quota) if g1_annual_quota > 0 else 0.0
                                g1_sip_earned = g1_fy_att * g1_annual_sip
                                
                                g2_annual_sip = salary * target_incentive * g2_weight * conversion_rate
                                # Goal 2 Annual Quota: taken from Goal 2 Quota in Pay Sheet
                                if goal2_quota_col_pay and pd.notna(row[goal2_quota_col_pay]):
                                    g2_annual_quota = parse_value(row[goal2_quota_col_pay])
                                else:
                                    g2_annual_quota = quota2_parsed
                                    
                                g2_ytd_rev = parse_value(row[ytd_territory_rev_col]) if ytd_territory_rev_col else 0.0
                                g2_fy_att = (g2_ytd_rev / g2_annual_quota) if g2_annual_quota > 0 else 0.0
                                g2_sip_earned = g2_fy_att * g2_annual_sip
                                
                                total_earned = g1_sip_earned + g2_sip_earned
                                
                                m = {
                                    'goal1_annual_sip': g1_annual_sip,
                                    'goal1_annual_quota': g1_annual_quota,
                                    'goal1_ytd_rev': g1_ytd_rev,
                                    'goal1_fy_attainment': g1_fy_att,
                                    'goal1_sip_earned': g1_sip_earned,
                                    'goal2_annual_sip': g2_annual_sip,
                                    'goal2_annual_quota': g2_annual_quota,
                                    'goal2_ytd_rev': g2_ytd_rev,
                                    'goal2_fy_attainment': g2_fy_att,
                                    'goal2_sip_earned': g2_sip_earned,
                                    'total_sip_earned': total_earned,
                                    'total_sip_paid': total_paid
                                }
                            else:
                                # Only 1 Goal logic
                                annual_sip = salary * target_incentive * conversion_rate
                                # Annual Quota: taken from Goal 1 Quota in Pay Sheet
                                if goal1_quota_col_pay and pd.notna(row[goal1_quota_col_pay]):
                                    annual_quota = parse_value(row[goal1_quota_col_pay])
                                else:
                                    annual_quota = parse_value(get_dict_val(participant, annual_quota_map, default=0.0))
                                    
                                ytd_rev = parse_value(row[ytd_rev_col]) if ytd_rev_col else 0.0
                                fy_att = (ytd_rev / annual_quota) if annual_quota > 0 else 0.0
                                total_earned = fy_att * annual_sip
                                
                                m = {
                                    'annual_sip': annual_sip,
                                    'annual_quota': annual_quota,
                                    'ytd_rev': ytd_rev,
                                    'fy_attainment': fy_att,
                                    'total_sip_earned': total_earned,
                                    'total_sip_paid': total_paid
                                }

                            # Manager Email & CC list
                            mgr_email_raw = get_dict_val(participant, manager_email_map, default="")
                            mgr_email_str = str(mgr_email_raw).strip() if pd.notna(mgr_email_raw) and str(mgr_email_raw).strip().lower() != 'nan' else ""
                            
                            payee_cc_list = []
                            if mgr_email_str:
                                for m_email in mgr_email_str.replace(';', ',').split(','):
                                    m_clean = m_email.strip()
                                    if m_clean and m_clean.lower() not in [c.lower() for c in payee_cc_list]:
                                        payee_cc_list.append(m_clean)
                            for c in cc_list:
                                if c.lower() not in [m_item.lower() for m_item in payee_cc_list]:
                                    payee_cc_list.append(c)

                            # Requirement: if Total SIP Earned in either table is less than Total YTD Earnings in Pay Sheet
                            has_old_role = (round(m['total_sip_earned'], 2) < round(total_ytd_earnings, 2)) and (total_ytd_earnings > 0.0)
                            old_role_sip_earned = max(0.0, total_ytd_earnings - m['total_sip_earned']) if has_old_role else 0.0

                            email_html = generate_email_html(
                                participant=participant,
                                is_both_goals=has_both_goals,
                                m=m,
                                pay_date=payee_pay_date,
                                sip_month=sip_month,
                                has_old_role=has_old_role,
                                old_role_sip_earned=old_role_sip_earned,
                                total_ytd_earnings=total_ytd_earnings
                            )

                            total_ytd_sip_earned_calc = (m['total_sip_earned'] + old_role_sip_earned) if has_old_role else m['total_sip_earned']

                            email_records.append({
                                'participant': participant,
                                'email': recipient_email,
                                'manager_email': mgr_email_str,
                                'cc_list': payee_cc_list,
                                'pay_date': payee_pay_date,
                                'currency': currency_code,
                                'total_sip_earned': m['total_sip_earned'],
                                'total_sip_paid': m['total_sip_paid'],
                                'total_ytd_earnings': total_ytd_earnings,
                                'has_old_role': has_old_role,
                                'old_role_sip_earned': old_role_sip_earned,
                                'total_ytd_sip_earned': total_ytd_sip_earned_calc,
                                'is_both_goals': has_both_goals,
                                'html': email_html
                            })

                            table_type_label = ("Both Goals" if has_both_goals else "Only 1 Goal") + (" (+ Old Role)" if has_old_role else "")
                            preview_table_data.append({
                                'Participant': participant,
                                'Email': recipient_email,
                                'Manager Email': mgr_email_str if mgr_email_str else "-",
                                'Currency': currency_code,
                                'Pay Date': payee_pay_date,
                                'Table Type': table_type_label,
                                'Conversion Rate': conversion_rate,
                                'Total YTD SIP Earned': format_curr(total_ytd_sip_earned_calc),
                                'Total SIP Earned (New Role)': format_curr(m['total_sip_earned']),
                                'Old Role SIP Earned': format_curr(old_role_sip_earned) if has_old_role else "-",
                                'Total SIP Paid': format_curr(m['total_sip_paid'])
                            })

                        # Show summary preview dataframe
                        st.dataframe(pd.DataFrame(preview_table_data), use_container_width=True)

                        # Actions
                        st.subheader("🛠 Step 5: Preview & Send Emails")
                        col_btn1, col_btn2 = st.columns(2)
                        
                        if 'preview_clicked' not in st.session_state:
                            st.session_state.preview_clicked = False
                            
                        with col_btn1:
                            if st.button("🔍Preview Emails", use_container_width=True):
                                st.session_state.preview_clicked = True
                                
                        with col_btn2:
                            send_clicked = st.button("🚀Send Emails", use_container_width=True)
                            
                        # Handle Preview Trigger
                        if st.session_state.preview_clicked:
                            st.markdown("### 📝 Preview of each participant's email:")
                            st.caption("The emails below have not been sent yet. Verify table format, metrics, and matched attachments.")
                            for item in email_records:
                                part_name = item['participant']
                                part_email = item['email']
                                html_body = item['html']
                                resolved_subject = email_subject_template.replace("{participant}", part_name)
                                participant_files = get_participant_attachments(part_name, uploaded_attachments)
                                item_pay_date = item.get('pay_date', pay_date)
                                
                                with st.expander(f"👤 {part_name} ({part_email}) - [{ 'Both Goals' if item['is_both_goals'] else 'Only 1 Goal' }]"):
                                    st.markdown(f"**From:** `{sender_email if sender_email else 'your_email@gmail.com'}`")
                                    st.markdown(f"**To:** `{part_email}`")
                                    if item['cc_list']:
                                        st.markdown(f"**CC:** `{', '.join(item['cc_list'])}`")
                                    st.markdown(f"**Subject:** `{resolved_subject}`")
                                    st.markdown(f"**Pay Date:** `{item_pay_date}` (Currency: `{item.get('currency', 'USD')}`)")
                                    
                                    if participant_files:
                                        st.markdown("**📎 Matched Attachments:**")
                                        for pf in participant_files:
                                            st.markdown(f"- `{pf.name}` ({len(pf.getvalue())/1024:.1f} KB)")
                                    else:
                                        st.markdown("*No attachments for this participant.*")
                                        
                                    st.markdown("---")
                                    st.components.v1.html(html_body, height=450, scrolling=True)

                        # Handle Send Trigger
                        if send_clicked:
                            if not sender_email or not password:
                                show_credentials_error_dialog()
                            else:
                                with st.status("🚀 Connecting to server and sending emails...") as status:
                                    try:
                                        st.write("Connecting to Gmail SMTP...")
                                        server = smtplib.SMTP("smtp.gmail.com", 587)
                                        server.starttls()
                                        server.login(sender_email, password)
                                        st.write("Login successful! Starting email dispatch...")
                                        
                                        success_count = 0
                                        error_count = 0
                                        progress_bar = st.progress(0)
                                        
                                        month_names = {3: "March", 6: "June", 9: "September", 12: "December"}
                                        month_name = month_names.get(sip_month, "Current Month")
                                        
                                        for i, item in enumerate(email_records):
                                            part_name = item['participant']
                                            part_email = item['email']
                                            html_body = item['html']
                                            item_cc = item.get('cc_list', [])
                                            total_earned = item.get('total_sip_earned', 0.0)
                                            total_paid = item.get('total_sip_paid', 0.0)
                                            has_old = item.get('has_old_role', False)
                                            earned_disp = item.get('total_ytd_sip_earned', total_earned) if has_old else total_earned
                                            item_pay_date = item.get('pay_date', pay_date)
                                            
                                            resolved_subject = email_subject_template.replace("{participant}", part_name)
                                            
                                            msg = MIMEMultipart('mixed')
                                            msg['Subject'] = resolved_subject
                                            msg['From'] = sender_email
                                            msg['To'] = part_email
                                            
                                            if item_cc:
                                                msg['Cc'] = ", ".join(item_cc)
                                                to_addrs = [part_email] + [c for c in item_cc if c.lower() != part_email.lower()]
                                            else:
                                                to_addrs = [part_email]
                                                
                                            msg_alternative = MIMEMultipart('alternative')
                                            plain_text = f"Hi {part_name},\n\nFor your information, you will not receive any {month_name} SIP payment in the upcoming pay cycle on {item_pay_date} because your YTD SIP Earned is {format_curr(earned_disp)} while your YTD SIP Paid is {format_curr(total_paid)}.\n\nPlease check your email client with HTML support to view your full YTD performance and payment breakdown table."
                                            msg_alternative.attach(MIMEText(plain_text, 'plain'))
                                            msg_alternative.attach(MIMEText(html_body, 'html'))
                                            msg.attach(msg_alternative)
                                            
                                            # Match and add file attachments
                                            participant_files = get_participant_attachments(part_name, uploaded_attachments)
                                            for pf in participant_files:
                                                part = MIMEBase('application', 'octet-stream')
                                                part.set_payload(pf.getvalue())
                                                encoders.encode_base64(part)
                                                part.add_header('Content-Disposition', f'attachment; filename="{pf.name}"')
                                                msg.attach(part)
                                            
                                            try:
                                                server.sendmail(sender_email, to_addrs, msg.as_string())
                                                attachment_info = f" with {len(participant_files)} attachment(s)" if participant_files else ""
                                                st.write(f"✅ Successfully sent to: **{part_name}** ({part_email}){attachment_info}")
                                                success_count += 1
                                            except Exception as send_err:
                                                st.write(f"❌ Error sending to **{part_name}** ({part_email}): {send_err}")
                                                error_count += 1
                                                
                                            progress_bar.progress((i + 1) / len(email_records))
                                            
                                        server.quit()
                                        status.update(label="Email dispatch process completed!", state="complete", expanded=False)
                                        st.balloons()
                                        st.success(f"🎉 Successfully sent **{success_count}** emails. Failed: **{error_count}** emails.")
                                        
                                    except Exception as smtp_err:
                                        st.error(f"❌ Gmail SMTP Connection Error: {smtp_err}")
                                        status.update(label="An error occurred during connection!", state="error", expanded=False)
                    else:
                        st.warning("⚠️ No participants match the current filtering criteria.")
                        
    except Exception as e:
        st.error(f"An error occurred while reading the Excel file: {e}")
