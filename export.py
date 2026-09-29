"""
export.py
---------
Pulls fully-calculated data from the Pacific Debt Dashboard API.
Upserts into Google Sheets or emails CSVs directly to partners.
"""

import os
import sys
import time
import json
import logging
import requests
import smtplib
import pandas as pd
import gspread
from email.message import EmailMessage
from datetime import datetime
from google.oauth2 import service_account

# ═══════════════════════════════════════════════════════════════════════
# CONFIG (Populated securely via GitHub Secrets)
# ═══════════════════════════════════════════════════════════════════════

API_URL = os.environ.get("API_URL", "").strip("[]\"' \t\n\r").rstrip("/")
if API_URL and not API_URL.startswith(("http://", "https://")):
    API_URL = f"https://{API_URL}"

API_USERNAME  = os.environ.get("API_USERNAME")
API_PASSWORD  = os.environ.get("API_PASSWORD")
GOOGLE_CREDS  = os.environ.get("GOOGLE_CREDS_JSON")

SMTP_SERVER   = os.environ.get("SMTP_SERVER", "smtp.gmail.com")
SMTP_PORT     = int(os.environ.get("SMTP_PORT", 465))
SMTP_USER     = os.environ.get("SMTP_USER")
SMTP_PASSWORD = os.environ.get("SMTP_PASSWORD")

SPREADSHEET_ID = "1wWmKrRkGS9fTSYYBpZjODeVo4TBh-Gfrhf2p2Zg9QPU"
SHEET_TAB_NAME = "Export"

# ═══════════════════════════════════════════════════════════════════════
# VENDOR DESTINATIONS
# ═══════════════════════════════════════════════════════════════════════

VENDOR_CONFIGS = [
    {
        "spreadsheet_id": "19bXpsndQ4SABd1OZpO5lRnSYO3fIcAcSZNx1IGDYusQ",
        "tab_name":       "Export",
        "campaigns":      ["SMI Media", "SMI Media 2"]
    },
    {
        "spreadsheet_id": "1RoE78ouXLmNUmfoKEhGafrYwivv53qZQCXOA1sBKAaU",
        "tab_name":       "Export",
        "campaigns":      ["Scale Up Media", "Scale Up Media 2", "Scale Up Media RS"]
    },
    {
        "spreadsheet_id": "1KocCHX1RCdG6qI_ePKIuzmSRjEy2juWpCsla5rVA3wI",
        "tab_name":       "Export",
        "campaigns":      ["Bearing Fruit", "Bearing Fruit 2", "Bearing Fruit 4", "Bearing Fruit Aged", "Bearing Fruit SpinWheel", "Bearing Fruit Taboola", "BF3", "BF5", "BF6"]
    },
    {
        # EMAIL DELIVERY CONFIG
        "emails":          "data@ingest.borrowbetter.com",
        "cc_emails":       "mikenittoli@pacificdebt.com, liezlalmin@pacificdebt.com",
        "campaigns":       ["BorrowBetter"],
        "email_subject":   "BorrowBetter Daily Report",
        "email_filename":  "BorrowBetter Summary Report - {date}.csv"
    }
]

MASTER_COLUMNS = [
    "Created Date", "Campaign", "Contact ID", "Stated Debt", "Credit Pull Debt", 
    "EnrolledDebt", "City", "State", "Zip", "Status", "Contacted", "Signature Date", 
    "Enrolled Date", "FPCD", "Dropped Date", "ClickId", "EFID", "LeadSourceId", 
    "UTM_AdGroup", "UTM_Campaign", "UTM_Content", "UTM_Medium", "UTM_Network", 
    "UTM_Source", "Enrolled By", "Payout", "Servicing Company"
]

logging.basicConfig(level=logging.INFO, format="%(asctime)s  %(levelname)s  %(message)s")
log = logging.getLogger(__name__)

# ═══════════════════════════════════════════════════════════════════════
# API FETCH LOGIC
# ═══════════════════════════════════════════════════════════════════════

def get_api_data():
    if not API_URL or not API_USERNAME or not API_PASSWORD:
        log.error("API_URL, API_USERNAME, or API_PASSWORD missing from environment.")
        sys.exit(1)

    log.info(f"Logging into API at {API_URL}...")
    login_resp = requests.post(f"{API_URL}/api/auth/login", json={
        "username": API_USERNAME,
        "password": API_PASSWORD
    })
    
    if not login_resp.ok:
        log.error(f"Login failed: {login_resp.text}")
        sys.exit(1)
        
    token = login_resp.json().get("token")
    headers = {"Authorization": f"Bearer {token}"}
    
    log.info("Fetching leads from API...")
    all_rows = []
    page = 1
    total_pages = 1
    
    while page <= total_pages:
        resp = requests.get(f"{API_URL}/api/leads/search?pageSize=5000&page={page}", headers=headers)
        if not resp.ok:
            log.error(f"Failed to fetch page {page}: {resp.text}")
            sys.exit(1)
            
        data = resp.json()
        all_rows.extend(data.get("rows", []))
        total_pages = data.get("pages", 1)
        
        log.info(f"  -> Fetched page {page}/{total_pages} ({len(all_rows)} rows total)")
        page += 1

    df = pd.DataFrame(all_rows)
    return df

# ═══════════════════════════════════════════════════════════════════════
# FORMATTING
# ═══════════════════════════════════════════════════════════════════════

def fmt_currency(val):
    try:
        if val is None or pd.isna(val) or str(val).strip() == "": return ""
        return f"${float(val):,.2f}"
    except: return ""

def fmt_date(val):
    try:
        if val is None or pd.isna(val) or str(val).strip() == "": return ""
        return pd.to_datetime(val).strftime("%Y-%m-%d")
    except: return ""

def fmt_datetime(val):
    try:
        if val is None or pd.isna(val) or str(val).strip() == "": return ""
        # The API returns 'YYYY-MM-DDTHH:MM:SS'. Reformatting to standard space-separated datetime.
        return str(val).replace("T", " ")
    except: return ""

def fmt_contacted(val):
    try:
        if pd.isna(val) or str(val).strip() == "": return ""
        return "Yes" if int(float(val)) == 1 else "No"
    except: return ""

def build_output(df):
    log.info("Formatting output DataFrame...")
    out = pd.DataFrame()
    
    # Try to use the detailed DateTime, fallback to standard CreatedDate if missing
    if "CreatedDateTime" in df.columns:
        out["Created Date"]  = df["CreatedDateTime"].apply(fmt_datetime)
        # Fill any blanks that didn't have a specific time with just the date
        out["Created Date"]  = out["Created Date"].replace("", pd.NA).fillna(df.get("CreatedDate", "").apply(fmt_date))
    else:
        out["Created Date"]  = df.get("CreatedDate", "").apply(fmt_date)
        
    out["Campaign"]          = df["Campaign"].fillna("")
    out["Contact ID"]        = df["ContactId"].astype(str)
    out["Stated Debt"]       = df["StatedDebt"].apply(fmt_currency)
    out["Credit Pull Debt"]  = df["CreditPullDebt"].apply(fmt_currency)
    out["EnrolledDebt"]      = df["EnrolledDebt"].apply(fmt_currency)
    out["City"]              = df["City"].fillna("")
    out["State"]             = df["State"].fillna("")
    out["Zip"]               = df["Zip"].fillna("").astype(str)
    out["Status"]            = df["Status"].fillna("")
    out["Contacted"]         = df["Contacted"].apply(fmt_contacted)
    out["Signature Date"]    = df["SignatureDate"].apply(fmt_date)
    out["Enrolled Date"]     = df["EnrolledDate"].apply(fmt_date)
    out["FPCD"]              = df["FPCD"].apply(fmt_date)
    out["Dropped Date"]      = df["DroppedDate"].apply(fmt_date)
    out["ClickId"]           = df.get("ClickId", pd.Series("", index=df.index)).fillna("")
    out["EFID"]              = df.get("EfId", pd.Series("", index=df.index)).fillna("")
    out["LeadSourceId"]      = df.get("LeadSourceId", pd.Series("", index=df.index)).fillna("")
    out["UTM_AdGroup"]       = df.get("UtmAdGroup", pd.Series("", index=df.index)).fillna("")
    out["UTM_Campaign"]      = df.get("UtmCampaign", pd.Series("", index=df.index)).fillna("")
    out["UTM_Content"]       = df.get("UtmContent", pd.Series("", index=df.index)).fillna("")
    out["UTM_Medium"]        = df.get("UtmMedium", pd.Series("", index=df.index)).fillna("")
    out["UTM_Network"]       = df.get("UtmNetwork", pd.Series("", index=df.index)).fillna("")
    out["UTM_Source"]        = df.get("UtmSource", pd.Series("", index=df.index)).fillna("")
    out["Enrolled By"]       = df["EnrolledBy"].fillna("")
    
    out["Payout"]            = df["payout"].apply(fmt_currency) 
    out["Servicing Company"] = df["ServicingCompany"].fillna("PDR")

    out.sort_values(by="Created Date", ascending=False, inplace=True)
    out.fillna("", inplace=True)
    return out

# ═══════════════════════════════════════════════════════════════════════
# EXPORT DESTINATIONS (GSHEETS OR EMAIL)
# ═══════════════════════════════════════════════════════════════════════

def get_gspread_client():
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    if not GOOGLE_CREDS:
        raise ValueError("GOOGLE_CREDS_JSON environment variable is missing!")
    creds_info = json.loads(GOOGLE_CREDS)
    creds = service_account.Credentials.from_service_account_info(creds_info, scopes=scopes)
    return gspread.authorize(creds)

def upsert_to_sheet(client, df, spreadsheet_id, tab_name, columns, campaign_filter=None):
    if campaign_filter:
        df = df[df["Campaign"].isin(campaign_filter)].copy()
        if len(df) == 0:
            log.info(f"  No rows match campaigns {campaign_filter} — skipping.")
            return

    spreadsheet = client.open_by_key(spreadsheet_id)
    sheet       = spreadsheet.worksheet(tab_name)

    all_rows = df[columns].astype(str).values.tolist()
    data_to_upload = [columns] + all_rows

    log.info(f"  Clearing and resizing '{tab_name}' to prevent limits...")
    sheet.clear()
    try:
        sheet.resize(rows=len(data_to_upload), cols=len(columns))
    except Exception:
        pass 

    log.info(f"  Uploading {len(data_to_upload)} rows to '{tab_name}' in chunks...")
    CHUNK_SIZE = 25000
    for i in range(0, len(data_to_upload), CHUNK_SIZE):
        chunk = data_to_upload[i:i + CHUNK_SIZE]
        start_row = i + 1
        
        for attempt in range(3):
            try:
                sheet.update(range_name=f"A{start_row}", values=chunk, value_input_option="USER_ENTERED")
                time.sleep(2)
                break
            except Exception as e:
                if attempt == 2: raise e
                time.sleep(5)
    log.info(f"  Done writing to '{tab_name}'.")

def email_export(df, config, columns):
    if not SMTP_USER or not SMTP_PASSWORD:
        log.error("SMTP_USER or SMTP_PASSWORD not set. Cannot send email.")
        return
    
    campaign_filter = config.get("campaigns")
    if campaign_filter:
        df = df[df["Campaign"].isin(campaign_filter)].copy()
        if len(df) == 0:
            log.info(f"  No rows match campaigns {campaign_filter} — skipping email.")
            return

    recipients = config["emails"]
    log.info(f"  Emailing CSV to {recipients}...")
    df = df[columns]

    # Generate dynamic month and day (e.g., 'September 28')
    date_str = f"{datetime.now().strftime('%B')} {datetime.now().day}"
    
    subject_line = config.get("email_subject", f"Pacific Debt - Daily Data Feed ({date_str})")
    filename_template = config.get("email_filename", "partner_data_{date}.csv")
    csv_filename = filename_template.format(date=date_str)

    msg = EmailMessage()
    msg['Subject'] = subject_line
    msg['From'] = SMTP_USER
    msg['To'] = recipients
    
    # NEW: Apply CC addresses if provided in the config
    cc_recipients = config.get("cc_emails")
    if cc_recipients:
        msg['Cc'] = cc_recipients
        log.info(f"  CC'ing: {cc_recipients}")

    msg.set_content("Hello,\n\nAttached is your automated daily CSV data feed.\n\nBest regards,\nPacific Debt")

    csv_data = df.to_csv(index=False)
    msg.add_attachment(csv_data.encode('utf-8'), maintype='text', subtype='csv', filename=csv_filename)

    try:
        if SMTP_PORT == 465:
            with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
                server.login(SMTP_USER, SMTP_PASSWORD)
                server.send_message(msg)
        else:
            with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
                server.starttls()
                server.login(SMTP_USER, SMTP_PASSWORD)
                server.send_message(msg)
        log.info(f"  Email sent successfully with attachment: {csv_filename}")
    except Exception as e:
        log.error(f"  Failed to send email: {e}")

# ═══════════════════════════════════════════════════════════════════════
# MAIN RUNNER
# ═══════════════════════════════════════════════════════════════════════

def main():
    log.info("=" * 60)
    log.info(f"Export started: {datetime.now()}")
    try:
        raw_df = get_api_data()
        if len(raw_df) == 0:
            log.info("No rows returned from API — nothing to update.")
            return

        output_df = build_output(raw_df)
        client = get_gspread_client()

        log.info("Writing to MASTER sheet...")
        upsert_to_sheet(client, output_df, SPREADSHEET_ID, SHEET_TAB_NAME, MASTER_COLUMNS)

        for config in VENDOR_CONFIGS:
            if "spreadsheet_id" in config:
                log.info(f"Writing to vendor sheet for {config['campaigns']}...")
                upsert_to_sheet(client, output_df, config["spreadsheet_id"], config["tab_name"], MASTER_COLUMNS, config["campaigns"])
            
            if "emails" in config:
                log.info(f"Sending email export for {config['campaigns']}...")
                email_export(output_df, config, MASTER_COLUMNS)

        log.info(f"Export completed: {datetime.now()}")

    except Exception as e:
        log.error(f"Export FAILED: {e}", exc_info=True)
        sys.exit(1)

if __name__ == "__main__":
    main()
