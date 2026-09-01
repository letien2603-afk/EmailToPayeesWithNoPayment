import streamlit as st
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
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

def generate_email_html(participant, goal1_val, goal2_val, has_goal2, ytd_earnings, months_data, ytd_sum, pay_date):
    """Generate professional Excel-style HTML table for email body."""
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
        min-width: 320px;
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
      .right-align {{
        text-align: right;
      }}
    </style>
    </head>
    <body>
      <p>Hi {participant},</p>
      <p>For your information, you will not receive any SIP payment in the upcoming pay cycle on <strong>{pay_date}</strong>.</p>
      <p>Please see the YTD payment breakdown below for more information:</p>
      
      <table class="paysheet-table">
        <!-- Participant Row -->
        <tr style="border: none;">
          <td colspan="2" class="bold" style="border: none; font-size: 12pt; padding: 6px 8px; text-align: left;">{participant}</td>
        </tr>
        
        <!-- Goal 1 -->
        <tr>
          <td class="left-align">Goal 1 FY Attainment</td>
          <td class="right-align">{format_pct(goal1_val)}</td>
        </tr>
    """
    
    # Goal 2 (if applicable)
    if has_goal2:
        html += f"""
        <tr>
          <td class="left-align">Goal 2 FY Attainment</td>
          <td class="right-align">{format_pct(goal2_val)}</td>
        </tr>
        """
        
    # YTD SIP Earnings
    html += f"""
        <tr>
          <td class="left-align bold underline">YTD SIP Earned</td>
          <td class="right-align bold underline">{format_curr(ytd_earnings)}</td>
        </tr>
        
        <!-- Blank separator row -->
        <tr style="border: none; height: 15px;">
          <td colspan="2" style="border: none; height: 15px;"></td>
        </tr>
        
        <!-- Submitted to Payroll Header -->
        <tr style="border: none;">
          <td colspan="2" class="bold" style="border: none; text-align: left; padding: 4px 8px;">Submitted to Payroll</td>
        </tr>
    """
    
    # Monthly Rows
    for month_lbl, val in months_data:
        html += f"""
        <tr>
          <td class="left-align">{month_lbl}</td>
          <td class="right-align">{format_curr(val)}</td>
        </tr>
        """
        
    # YTD Sum Row (YTD SIP Paid)
    html += f"""
        <tr>
          <td class="left-align bold underline">YTD SIP Paid</td>
          <td class="right-align bold underline">{format_curr(ytd_sum)}</td>
        </tr>
      </table>
      
      <p>Regards</p>
    </body>
    </html>
    """
    return html

# App UI Design
st.title("📧 No SIP Payments Email Automator")

# Sidebar Configuration
st.sidebar.header("⚙️ Fill in the required information")

sender_email = st.sidebar.text_input("1. Sender Email (Gmail):", placeholder="your_email@gmail.com")
password = st.sidebar.text_input("2. Gmail App Password (16 characters):", type="password", placeholder="xxxx xxxx xxxx xxxx")
cc_input = st.sidebar.text_input("3. CC Emails (Separated by commas):", placeholder="manager@example.com, hr@example.com")

st.sidebar.markdown("---")
st.sidebar.header("✉️ Email Subject")
email_subject_template = st.sidebar.text_input(
    "Email Subject:", 
    value="June SIP Payout Notification",
    help="Use the placeholder `{participant}` to dynamically replace with each recipient's name."
)

st.sidebar.markdown("---")
st.sidebar.header("🗓️ Payroll Cycle Information")
pay_date = st.sidebar.text_input("4. Pay Date:", value="2026-09-15")
sip_month = st.sidebar.selectbox("5. SIP month to pay:", options=[3, 6, 9, 12], index=2)

# File uploader on main screen
st.subheader("📁 Step 1:Upload Payroll File")
uploaded_file = st.file_uploader("Choose payfile file (.xlsx)", type=["xlsx"])

# Parse CC list
cc_list = [e.strip() for e in cc_input.split(",") if e.strip()]

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
                    
                    total_matches = len(filtered_df)
                    st.subheader("📊 Step 2: Statistics & Filtering Results Check")
                    st.info(f"Found **{total_matches}** participants matching the filtering criteria (Include in PayFile = 1 and Net Commissions < 0).")
                    
                    if total_matches > 0:
                        # Display preview dataframe
                        display_cols = [participant_col_pay, 'Email', include_col, net_comm_col]
                        # Show some columns if exists
                        for add_c in ["YTD ATTAINMENT", "Total YTD Earnings"]:
                            found_col = find_column(paysheet_df, [add_c])
                            if found_col:
                                display_cols.append(found_col)
                        
                        st.dataframe(filtered_df[display_cols].astype(str), use_container_width=True)
                        
                        # Process each filtered participant's data to generate email contents
                        email_records = []
                        for idx, row in filtered_df.iterrows():
                            participant = row[participant_col_pay]
                            recipient_email = row['Email']
                            
                            if pd.isna(recipient_email) or not str(recipient_email).strip():
                                continue
                                
                            goal1_val = get_row_val(row, paysheet_df, ["YTD ATTAINMENT", "YTD Attainment", "Goal 1 FY Attainment"])
                            goal2_val = get_row_val(row, paysheet_df, ["TERRITORY YTD ATTAINMENT", "Territory YTD Attainment", "Goal 2 FY Attainment"])
                            goal2_parsed = parse_value(goal2_val)
                            has_goal2 = goal2_parsed != 0.0
                            
                            ytd_earnings = get_row_val(row, paysheet_df, ["Total YTD Earnings", "YTD SIP Earnings", "YTD SIP Earned"])
                            
                            months_data = []
                            monthly_sum = 0.0
                            
                            # Jan & Feb
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
                                        
                            email_html = generate_email_html(
                                participant=participant,
                                goal1_val=goal1_val,
                                goal2_val=goal2_val,
                                has_goal2=has_goal2,
                                ytd_earnings=ytd_earnings,
                                months_data=months_data,
                                ytd_sum=monthly_sum,
                                pay_date=pay_date
                            )
                            
                            email_records.append({
                                'participant': participant,
                                'email': recipient_email,
                                'html': email_html
                            })
                        
                        # Actions
                        st.subheader("🛠️ Step 3: Preview & Send Emails")
                        col_btn1, col_btn2 = st.columns(2)
                        
                        # Store in session state to handle action states cleanly
                        if 'preview_clicked' not in st.session_state:
                            st.session_state.preview_clicked = False
                            
                        with col_btn1:
                            if st.button("🔍 Preview Emails", use_container_width=True):
                                st.session_state.preview_clicked = True
                                
                        with col_btn2:
                            send_clicked = st.button("🚀 Send Emails", use_container_width=True)
                            
                        # Handle Preview Trigger
                        if st.session_state.preview_clicked:
                            st.markdown("### 📝 Preview of each participant's email:")
                            st.caption("The emails below have not been sent yet. You can verify the table format and content.")
                            for item in email_records:
                                part_name = item['participant']
                                part_email = item['email']
                                html_body = item['html']
                                
                                # Process the subject dynamically
                                resolved_subject = email_subject_template.replace("{participant}", part_name)
                                
                                with st.expander(f"👤 {part_name} ({part_email})"):
                                    st.markdown(f"**From:** `{sender_email if sender_email else 'your_email@gmail.com'}`")
                                    st.markdown(f"**To:** `{part_email}`")
                                    if cc_list:
                                        st.markdown(f"**CC:** `{', '.join(cc_list)}`")
                                    st.markdown(f"**Subject:** `{resolved_subject}`")
                                    st.markdown("---")
                                    # Render raw HTML in streamlit safely
                                    st.components.v1.html(html_body, height=450, scrolling=True)
                                    
                        # Handle Send Trigger
                        if send_clicked:
                            if not sender_email or not password:
                                st.error("⚠️ Please fill in both the Sender Email and Gmail App Password in the sidebar before sending!")
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
                                        
                                        for i, item in enumerate(email_records):
                                            part_name = item['participant']
                                            part_email = item['email']
                                            html_body = item['html']
                                            
                                            # Process the subject dynamically
                                            resolved_subject = email_subject_template.replace("{participant}", part_name)
                                            
                                            # Create Multipart Email
                                            msg = MIMEMultipart('alternative')
                                            msg['Subject'] = resolved_subject
                                            msg['From'] = sender_email
                                            msg['To'] = part_email
                                            
                                            # CC Addresses
                                            if cc_list:
                                                msg['Cc'] = ", ".join(cc_list)
                                                to_addrs = [part_email] + cc_list
                                            else:
                                                to_addrs = [part_email]
                                                
                                            plain_text = f"Hi {part_name},\n\nFor your information, you will not receive any SIP payment in the upcoming pay cycle on {pay_date}.\n\nPlease check your email client with HTML support to view your full YTD performance and payment breakdown table."
                                            msg.attach(MIMEText(plain_text, 'plain'))
                                            msg.attach(MIMEText(html_body, 'html'))
                                            
                                            try:
                                                server.sendmail(sender_email, to_addrs, msg.as_string())
                                                st.write(f"✅ Successfully sent to: **{part_name}** ({part_email})")
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
    st.info("💡 Please drag and drop or select the payfile (.xlsx) Excel file to start.")
