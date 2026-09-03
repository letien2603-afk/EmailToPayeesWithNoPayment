import streamlit as st
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
import sys

# Set up Streamlit Page Configuration
st.set_page_config(
    page_title="Millie Agro Payroll Summary & Email Automator",
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

def format_pct(val):
    """Format float as percentage (e.g. 0.45 -> '45.00%')."""
    return f"{parse_value(val) * 100:.2f}%"

def format_curr(val):
    """Format float as standard currency (e.g. 15000.0 -> '15,000.00')."""
    return f"{parse_value(val):,.2f}"

def find_column(df, aliases):
    """Find a column in a dataframe matching any of the specified aliases (case-insensitive)."""
    for alias in aliases:
        for col in df.columns:
            if str(col).strip().lower() == alias.strip().lower():
                return col
    return None

def get_row_val(row, df, aliases, default=0.0):
    """Get the row's column value based on aliases."""
    col = find_column(df, aliases)
    if col is not None:
        return row[col]
    return default

def check_column_exists(df, aliases):
    """Check if any of the aliases exist in the dataframe columns."""
    return find_column(df, aliases) is not None

# Dialog modal popup for missing credentials (to pop out on screen instead of showing at the bottom)
@st.dialog("Configuration Error")
def show_credentials_error_dialog():
    st.error("Please fill in both the Sender Email and Gmail App Password in the sidebar before sending!")
    if st.button("OK", use_container_width=True):
        st.rerun()

def generate_email_html(participant, goal1_val, goal1_earning, goal2_val, goal2_earning, has_goal2, ytd_earnings, months_data, ytd_sum, adjustments_data, pay_date, sip_month):
    """Generate professional Excel-style HTML table for email body with updated Table2 and Table3 logic."""
    # Determine Goal names and rows based on has_goal2 (Territory Attainment > 0%)
    if has_goal2:
        goal1_label = "Goal 1"
        goal2_row_html = f"""
        <tr>
          <td class="left-align">Goal 2</td>
          <td class="center-align">{format_pct(goal2_val)}</td>
          <td class="right-align">{format_curr(goal2_earning)}</td>
        </tr>
        """
    else:
        goal1_label = "Goal"
        goal2_row_html = ""

    # Dynamic intro message based on sip_month (Requirements 1, 2, 3, 4)
    month_names = {
        3: "March",
        6: "June",
        9: "September",
        12: "December"
    }
    month_name = month_names.get(sip_month, "Current Month")
    intro_text = f"For your information, you will not receive any SIP payment in the upcoming pay cycle on <strong>{pay_date}</strong> because you have a negative true-up in {month_name} payout."

    # Header & Intro text
    html = f"""
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
        min-width: 360px;
        margin-top: 15px;
        margin-bottom: 15px;
      }}
      table.paysheet-table td {{
        border: 1px solid #d3d3d3;
        padding: 4px 8px;
        vertical-align: middle;
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
      .center-align {{
        text-align: center;
      }}
      .right-align {{
        text-align: right;
      }}
    </style>
    </head>
    <body>
      <p>Hi {participant},</p>
      <p>{intro_text}</p>
      <p>Please see the YTD payment breakdown below for more information:</p>
      
      <table class="paysheet-table">
        <!-- Participant Row -->
        <tr style="border: none;">
          <td colspan="3" class="bold" style="border: none; font-size: 12pt; padding: 6px 8px; text-align: left;">{participant}</td>
        </tr>
        
        <!-- Column Headers (as requested in Table2 structure) -->
        <tr>
          <td style="border: none;"></td>
          <td class="bold center-align">FY Attainment</td>
          <td class="bold right-align">YTD SIP Earned</td>
        </tr>
        
        <!-- Goal 1 / Goal Row -->
        <tr>
          <td class="left-align">{goal1_label}</td>
          <td class="center-align">{format_pct(goal1_val)}</td>
          <td class="right-align">{format_curr(goal1_earning)}</td>
        </tr>
        
        {goal2_row_html}
        
        <!-- Total YTD SIP Earned row -->
        <tr>
          <td class="left-align bold underline">YTD SIP Earned</td>
          <td></td>
          <td class="right-align bold underline">{format_curr(ytd_earnings)}</td>
        </tr>
        
        <!-- Blank separator row -->
        <tr style="border: none; height: 15px;">
          <td colspan="3" style="border: none; height: 15px;"></td>
        </tr>
        
        <!-- Submitted to Payroll Header -->
        <tr style="border: none;">
          <td colspan="3" class="bold" style="border: none; text-align: left; padding: 4px 8px;">Submitted to Payroll</td>
        </tr>
    """
    
    # Monthly Rows (aligned to the 3-column format)
    for month_lbl, val in months_data:
        html += f"""
        <tr>
          <td colspan="2" class="left-align">{month_lbl}</td>
          <td class="right-align">{format_curr(val)}</td>
        </tr>
        """
        
    # YTD Sum Row (YTD SIP Paid aligned under the earnings column)
    html += f"""
        <tr>
          <td colspan="2" class="left-align bold underline">YTD SIP Paid</td>
          <td class="right-align bold underline">{format_curr(ytd_sum)}</td>
        </tr>
    """

    # 1. Negative Adjustment Rows (Requirement 1)
    # Only added if adjustments_data contains elements (which only happens if values are < 0)
    if adjustments_data:
        for adj_title, adj_val in adjustments_data:
            html += f"""
            <tr>
              <td colspan="2" class="left-align">{adj_title}</td>
              <td class="right-align">{format_curr(adj_val)}</td>
            </tr>
            """

        # 2. Current Month Payout Row calculation (Requirement 2)
        # Value = (value of row above "YTD SIP Paid" i.e. last month in months_data) + (sum of all adjustments)
        last_month_val = months_data[-1][1] if months_data else 0.0
        adjustments_sum = sum(val for lbl, val in adjustments_data)
        current_month_payout = last_month_val + adjustments_sum
        
        # Payout row title corresponds to the payroll month (March/June/September/December Payout)
        payout_row_title = f"{month_name} Payout"

        html += f"""
            <tr>
              <td colspan="2" class="left-align bold">{payout_row_title}</td>
              <td class="right-align bold">{format_curr(current_month_payout)}</td>
            </tr>
        """

    html += f"""
      </table>
      
      <p>Regards</p>
    </body>
    </html>
    """
    return html

def is_file_for_participant(filename, participant_name):
    """Check if the uploaded attachment file belongs to the participant (Requirement 5)."""
    fn = filename.lower()
    pn = participant_name.lower().strip()
    
    # Generate variations (e.g. spaces replaced by underscores, dashes, or removed)
    pn_variations = [
        pn,
        pn.replace(" ", "_"),
        pn.replace(" ", "-"),
        pn.replace(" ", "")
    ]
    return any(var in fn for var in pn_variations if var)

def get_participant_attachments(participant_name, uploaded_files):
    """Retrieve all uploaded files that match the participant's name."""
    matched_files = []
    if not uploaded_files:
        return matched_files
    for f in uploaded_files:
        if is_file_for_participant(f.name, participant_name):
            matched_files.append(f)
    return matched_files

# App UI Design
st.title("📧 Millie Agro SIP Payroll Email Automator")
st.markdown("Automated system to match participant lists, apply filtering conditions, attach matching files, and send/preview personalized SIP payroll email reports.")

# Sidebar Configuration
st.sidebar.header("⚙️ System Configuration")

sender_email = st.sidebar.text_input("Sender Gmail:", placeholder="your_email@gmail.com")
password = st.sidebar.text_input("Gmail App Password (16 characters):", type="password", placeholder="xxxx xxxx xxxx xxxx")
cc_input = st.sidebar.text_input("CC Emails (Separated by commas):", placeholder="manager@example.com, hr@example.com")

st.sidebar.markdown("---")
st.sidebar.header("✉️ Email Subject Configuration")
email_subject_template = st.sidebar.text_input(
    "Email Subject:", 
    placeholder="June No SIP Payout Notification"
)

st.sidebar.markdown("---")
st.sidebar.header("🚫 Participant Exclusion")
exclude_input = st.sidebar.text_input(
    "Exclude Participants (Comma-separated names):",
    placeholder="e.g. John Doe, Jane Smith"
)

st.sidebar.markdown("---")
st.sidebar.header("🗓️ Payroll Cycle Information")
pay_date = st.sidebar.text_input("4. Pay Date:", value="2026-09-15")
sip_month = st.sidebar.selectbox("5. SIP month to pay:", options=[3, 6, 9, 12], index=2)

# File uploader on main screen
st.subheader("📁 Step 1: Upload Payroll File (payfile_Dummy)")
uploaded_file = st.file_uploader("Choose payfile_Dummy Excel file (.xlsx)", type=["xlsx"])

st.subheader("📎 Step 1b: Upload Email Attachments (Optional)")
uploaded_attachments = st.file_uploader(
    "Upload files to attach to participant emails (Requirement 5):",
    accept_multiple_files=True,
    help="If you upload files, the app will match each file to a participant if their name is found in the filename (e.g. 'Christa Vish.xlsx' matches 'Christa Vish')."
)

# Parse CC list
cc_list = [e.strip() for e in cc_input.split(",") if e.strip()]

# Parse exclusion list
exclude_list = [name.strip().lower() for name in exclude_input.split(",") if name.strip()]

if uploaded_file is not None:
    st.success("File uploaded successfully! Processing data...")
    
    try:
        # Load excel file
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names
        
        # Auto-detect Paysheet sheet name
        paysheet_name = None
        for name in ["Paysheet", "Pay Sheet", "PaySheet", "paysheet", "OIC Pay Sheet"]:
            if name in sheet_names:
                paysheet_name = name
                break
                
        # Auto-detect OIC Participant Details sheet name
        oic_sheet_name = None
        for name in ["OIC Participant Details", "Participant Details", "OIC Participants", "OIC Participant Details "]:
            if name in sheet_names:
                oic_sheet_name = name
                break
                
        if paysheet_name is None:
            st.error(f"Could not find 'Paysheet' or 'Pay Sheet' sheet in the file. Available sheets: {sheet_names}")
        elif oic_sheet_name is None:
            st.error(f"Could not find 'OIC Participant Details' sheet in the file. Available sheets: {sheet_names}")
        else:
            # Load Dataframes
            paysheet_df = pd.read_excel(xls, sheet_name=paysheet_name)
            oic_df = pd.read_excel(xls, sheet_name=oic_sheet_name)
            
            # Clean column names
            paysheet_df.columns = [str(c).strip() for c in paysheet_df.columns]
            oic_df.columns = [str(c).strip() for c in oic_df.columns]
            
            # Identify columns
            participant_col_pay = find_column(paysheet_df, ["*Participant", "Participant"])
            participant_col_oic = find_column(oic_df, ["*Participant", "Participant"])
            email_col_oic = find_column(oic_df, ["E-Mail Address", "Email Address", "Email", "E-Mail"])
            
            if not participant_col_pay or not participant_col_oic or not email_col_oic:
                st.error("Excel file structure mismatch! Please check '*Participant' and 'E-Mail Address' columns.")
            else:
                # Match email addresses (Step 3)
                email_map = oic_df.set_index(participant_col_oic)[email_col_oic].to_dict()
                paysheet_df['Email'] = paysheet_df[participant_col_pay].map(email_map)
                
                # Filtering logic (Step 4 & 5)
                include_col = find_column(paysheet_df, ["Include in This Month's PayFile? (1=Yes, 0=No)", "Include in This Month's PayFile?", "Include in PayFile"])
                net_comm_col = find_column(paysheet_df, ["Net Commissions Payable", "Net Commission Payable"])
                
                if not include_col or not net_comm_col:
                    st.error("Filtering condition columns 'Include...' or 'Net Commissions Payable' not found!")
                else:
                    # Filter: Include in Payfile == 1 and Net Commissions < 0
                    filtered_df = paysheet_df[
                        (paysheet_df[include_col] == 1) & 
                        (paysheet_df[net_comm_col].apply(parse_value) < 0)
                    ]
                    
                    st.subheader("📊 Step 2: Statistics & Filtering Results Check")
                    
                    # 1. Process Manual Exclusions
                    if exclude_list:
                        raw_filtered_names = filtered_df[participant_col_pay].astype(str).str.strip().tolist()
                        to_exclude = [name for name in raw_filtered_names if name.lower() in exclude_list]
                        if to_exclude:
                            filtered_df = filtered_df[~filtered_df[participant_col_pay].astype(str).str.strip().str.lower().isin(exclude_list)]
                            st.info(f"🚫 **Excluded {len(to_exclude)} participant(s)** as requested: `{', '.join(to_exclude)}`")
                    
                    # 2. Check for Missing Email Addresses and trigger warning
                    missing_email_df = filtered_df[filtered_df['Email'].isna() | (filtered_df['Email'].astype(str).str.strip() == '')]
                    if len(missing_email_df) > 0:
                        missing_names = missing_email_df[participant_col_pay].astype(str).str.strip().tolist()
                        st.warning(f"⚠️ **Warning: Missing Email Address** for **{len(missing_email_df)}** participant(s): `{', '.join(missing_names)}`. They have been automatically removed from the active email and preview list.")
                        # Remove participants with missing emails from sending list
                        filtered_df = filtered_df[~(filtered_df['Email'].isna() | (filtered_df['Email'].astype(str).str.strip() == ''))]
                    
                    total_matches = len(filtered_df)
                    st.success(f"Found **{total_matches}** active participant(s) for the email list (Include in PayFile = 1, Net Commissions < 0, not manually excluded, and has valid email).")
                    
                    if total_matches > 0:
                        # Display preview dataframe
                        display_cols = [participant_col_pay, 'Email', include_col, net_comm_col]
                        # Show some columns if exists
                        for add_c in ["YTD ATTAINMENT", "YTD Earning", "TERRITORY YTD ATTAINMENT", "TERRITORY YTD EARNING"]:
                            found_col = find_column(paysheet_df, [add_c])
                            if found_col:
                                display_cols.append(found_col)
                        
                        st.dataframe(filtered_df[display_cols].astype(str), use_container_width=True)
                        
                        # Process each filtered participant's data to generate email contents
                        email_records = []
                        for idx, row in filtered_df.iterrows():
                            participant = row[participant_col_pay]
                            recipient_email = row['Email']
                            
                            # Retrieve Goal 1 metrics
                            goal1_val = get_row_val(row, paysheet_df, ["YTD ATTAINMENT", "YTD Attainment", "Goal 1 FY Attainment"])
                            goal1_earning = get_row_val(row, paysheet_df, ["YTD Earning", "YTD Earnings", "Total YTD Earnings", "YTD SIP Earnings", "YTD SIP Earned"])
                            goal1_earning_parsed = parse_value(goal1_earning)

                            # Retrieve Goal 2 metrics
                            goal2_val = get_row_val(row, paysheet_df, ["TERRITORY YTD ATTAINMENT", "Territory YTD Attainment", "Goal 2 FY Attainment"])
                            goal2_earning = get_row_val(row, paysheet_df, ["TERRITORY YTD EARNING", "Territory YTD Earning", "Territory YTD Earnings"])
                            
                            goal2_parsed = parse_value(goal2_val)
                            has_goal2 = goal2_parsed > 0.0  # Show Goal 2 only if Territory YTD Attainment > 0%
                            goal2_earning_parsed = parse_value(goal2_earning) if has_goal2 else 0.0
                            
                            # Dynamic YTD SIP Earned calculation based on has_goal2
                            ytd_earnings_parsed = goal1_earning_parsed + goal2_earning_parsed
                            
                            months_data = []
                            monthly_sum = 0.0
                            
                            # Jan & Feb (always included)
                            jan_val = parse_value(get_row_val(row, paysheet_df, ["January Paid (Q1 Month 1 Draw)", "Jan Paid (Q1 Month 1 Draw)"]))
                            feb_val = parse_value(get_row_val(row, paysheet_df, ["February Paid (Q1 Month 2 Draw)", "Feb Paid (Q1 Month 2 Draw)"]))
                            months_data.append(("Jan", jan_val))
                            months_data.append(("Feb", feb_val))
                            monthly_sum += jan_val + feb_val
                            
                            # Mar
                            if sip_month == 3:
                                mar_val = parse_value(get_row_val(row, paysheet_df, ["March True-up", "Mar True-up"]))
                                months_data.append(("Mar", mar_val))
                                monthly_sum += mar_val
                            else:
                                mar_val = parse_value(get_row_val(row, paysheet_df, ["March Paid (Q1 Month 3 True-up)", "Mar Paid (Q1 Month 3 True-up)", "March True-up"]))
                                months_data.append(("Mar", mar_val))
                                monthly_sum += mar_val
                                
                                # Apr (Optional)
                                if check_column_exists(paysheet_df, ["April Paid (Q2 Month 1 Draw)", "Apr Paid (Q2 Month 1 Draw)"]):
                                    apr_val = parse_value(get_row_val(row, paysheet_df, ["April Paid (Q2 Month 1 Draw)", "Apr Paid (Q2 Month 1 Draw)"]))
                                    months_data.append(("Apr", apr_val))
                                    monthly_sum += apr_val
                                    
                                # May (Optional)
                                if check_column_exists(paysheet_df, ["May Paid (Q2 Month 2 Draw)", "May Paid"]):
                                    may_val = parse_value(get_row_val(row, paysheet_df, ["May Paid (Q2 Month 2 Draw)", "May Paid"]))
                                    months_data.append(("May", may_val))
                                    monthly_sum += may_val
                                    
                                # Jun
                                if sip_month == 6:
                                    jun_val = parse_value(get_row_val(row, paysheet_df, ["Jun True-up", "June True-up", "Jun True Up"]))
                                    months_data.append(("Jun", jun_val))
                                    monthly_sum += jun_val
                                else:
                                    jun_val = parse_value(get_row_val(row, paysheet_df, ["June Paid (Q2 Month 3 True-up)", "Jun Paid (Q2 Month 3 True-up)"]))
                                    months_data.append(("Jun", jun_val))
                                    monthly_sum += jun_val
                                    
                                    # Jul (Optional)
                                    if check_column_exists(paysheet_df, ["Jul Paid (Q3 Month 1 Draw)", "July Paid (Q3 Month 1 Draw)"]):
                                        jul_val = parse_value(get_row_val(row, paysheet_df, ["Jul Paid (Q3 Month 1 Draw)", "July Paid (Q3 Month 1 Draw)"]))
                                        months_data.append(("Jul", jul_val))
                                        monthly_sum += jul_val
                                        
                                    # Aug (Optional)
                                    if check_column_exists(paysheet_df, ["Aug Paid (Q3 Month 2 Draw)", "August Paid (Q3 Month 2 Draw)"]):
                                        aug_val = parse_value(get_row_val(row, paysheet_df, ["Aug Paid (Q3 Month 2 Draw)", "August Paid (Q3 Month 2 Draw)"]))
                                        months_data.append(("Aug", aug_val))
                                        monthly_sum += aug_val
                                        
                                    # Sep
                                    if sip_month == 9:
                                        sep_val = parse_value(get_row_val(row, paysheet_df, ["Sep True-up", "September True-up", "Sep True Up"]))
                                        months_data.append(("Sep", sep_val))
                                        monthly_sum += sep_val
                                    else:
                                        sep_val = parse_value(get_row_val(row, paysheet_df, ["Sep Paid (Q3 Month 3 True-up)", "September Paid (Q3 Month 3 True-up)"]))
                                        months_data.append(("Sep", sep_val))
                                        monthly_sum += sep_val
                                        
                                        # Oct (Optional)
                                        if check_column_exists(paysheet_df, ["Oct Paid (Q4 Month 1 Draw)", "October Paid (Q4 Month 1 Draw)"]):
                                            oct_val = parse_value(get_row_val(row, paysheet_df, ["Oct Paid (Q4 Month 1 Draw)", "October Paid (Q4 Month 1 Draw)"]))
                                            months_data.append(("Oct", oct_val))
                                            monthly_sum += oct_val
                                            
                                        # Nov (Optional)
                                        if check_column_exists(paysheet_df, ["Nov Paid (Q4 Month 2 Draw)", "November Paid (Q4 Month 2 Draw)", "Nov Paid"]):
                                            nov_val = parse_value(get_row_val(row, paysheet_df, ["Nov Paid (Q4 Month 2 Draw)", "November Paid (Q4 Month 2 Draw)", "Nov Paid"]))
                                            months_data.append(("Nov", nov_val))
                                            monthly_sum += nov_val
                                            
                                        # Dec
                                        dec_val = parse_value(get_row_val(row, paysheet_df, ["Dec True-up", "December True-up", "Dec True Up", "December Paid (Q4 Month 3 True-up)"]))
                                        months_data.append(("Dec", dec_val))
                                        monthly_sum += dec_val

                            # Extract negative adjustments (SDR/ISR MBO, Prior Period/2025 True-up, One Time True-up, Leadership Payout)
                            adjustments_data = []
                            for adj_col_name in ["SDR/ISR MBO", "Prior Period/2025 True-up", "One Time True-up", "Leadership Payout"]:
                                actual_col = find_column(paysheet_df, [adj_col_name])
                                if actual_col:
                                    val = parse_value(row[actual_col])
                                    if val < 0:
                                        adjustments_data.append((adj_col_name, val))
                                        
                            email_html = generate_email_html(
                                participant=participant,
                                goal1_val=goal1_val,
                                goal1_earning=goal1_earning_parsed,
                                goal2_val=goal2_val,
                                goal2_earning=goal2_earning_parsed,
                                has_goal2=has_goal2,
                                ytd_earnings=ytd_earnings_parsed,
                                months_data=months_data,
                                ytd_sum=monthly_sum,
                                adjustments_data=adjustments_data,
                                pay_date=pay_date,
                                sip_month=sip_month
                            )
                            
                            email_records.append({
                                'participant': participant,
                                'email': recipient_email,
                                'html': email_html
                            })
                        
                        # Actions
                        st.subheader("🛠 Step 3: Preview & Send Emails")
                        col_btn1, col_btn2 = st.columns(2)
                        
                        # Store in session state to handle action states cleanly
                        if 'preview_clicked' not in st.session_state:
                            st.session_state.preview_clicked = False
                            
                        with col_btn1:
                            if st.button("🔍 (1) Preview Emails", use_container_width=True):
                                st.session_state.preview_clicked = True
                                
                        with col_btn2:
                            send_clicked = st.button("🚀 (2) Send Emails", use_container_width=True)
                            
                        # Handle Preview Trigger
                        if st.session_state.preview_clicked:
                            st.markdown("### 📝 Preview of each participant's email:")
                            st.caption("The emails below have not been sent yet. You can verify the table format, content, and matched attachments.")
                            for item in email_records:
                                part_name = item['participant']
                                part_email = item['email']
                                html_body = item['html']
                                
                                # Process the subject dynamically
                                resolved_subject = email_subject_template.replace("{participant}", part_name)
                                
                                # Match attachments for display in preview
                                participant_files = get_participant_attachments(part_name, uploaded_attachments)
                                
                                with st.expander(f"👤 {part_name} ({part_email})"):
                                    st.markdown(f"**From:** `{sender_email if sender_email else 'your_email@gmail.com'}`")
                                    st.markdown(f"**To:** `{part_email}`")
                                    if cc_list:
                                        st.markdown(f"**CC:** `{', '.join(cc_list)}`")
                                    st.markdown(f"**Subject:** `{resolved_subject}`")
                                    
                                    # List matched attachments in preview
                                    if participant_files:
                                        st.markdown("**📎 Matched Attachments:**")
                                        for pf in participant_files:
                                            st.markdown(f"- `{pf.name}` ({len(pf.getvalue())/1024:.1f} KB)")
                                    else:
                                        st.markdown("*No attachments matched for this participant.*")
                                        
                                    st.markdown("---")
                                    # Render raw HTML in streamlit safely
                                    st.components.v1.html(html_body, height=520, scrolling=True)
                                    
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
                                        
                                        # Dynamic month name for dynamic plain text
                                        month_names = {3: "March", 6: "June", 9: "September", 12: "December"}
                                        month_name = month_names.get(sip_month, "Current Month")
                                        
                                        for i, item in enumerate(email_records):
                                            part_name = item['participant']
                                            part_email = item['email']
                                            html_body = item['html']
                                            
                                            # Process the subject dynamically
                                            resolved_subject = email_subject_template.replace("{participant}", part_name)
                                            
                                            # Create Multipart Email
                                            msg = MIMEMultipart('mixed')  # use 'mixed' to support text alternative and file attachments
                                            
                                            msg['Subject'] = resolved_subject
                                            msg['From'] = sender_email
                                            msg['To'] = part_email
                                            
                                            # CC Addresses
                                            if cc_list:
                                                msg['Cc'] = ", ".join(cc_list)
                                                to_addrs = [part_email] + cc_list
                                            else:
                                                to_addrs = [part_email]
                                                
                                            # Add message bodies as alternative parts
                                            msg_alternative = MIMEMultipart('alternative')
                                            
                                            plain_text = f"Hi {part_name},\n\nFor your information, you will not receive any SIP payment in the upcoming pay cycle on {pay_date} because you have a negative true-up in {month_name} payout.\n\nPlease check your email client with HTML support to view your full YTD performance and payment breakdown table."
                                            msg_alternative.attach(MIMEText(plain_text, 'plain'))
                                            msg_alternative.attach(MIMEText(html_body, 'html'))
                                            
                                            msg.attach(msg_alternative)
                                            
                                            # Match and add file attachments (Requirement 5)
                                            participant_files = get_participant_attachments(part_name, uploaded_attachments)
                                            for pf in participant_files:
                                                part = MIMEBase('application', 'octet-stream')
                                                part.set_payload(pf.getvalue())
                                                encoders.encode_base64(part)
                                                part.add_header(
                                                    'Content-Disposition',
                                                    f'attachment; filename="{pf.name}"'
                                                )
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
else:
    st.info("💡 Please drag and drop or select the payfile_Dummy (.xlsx) Excel file to start.")
