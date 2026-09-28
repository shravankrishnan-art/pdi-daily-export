import os
import sys
import time
import json
import logging
import requests
import pandas as pd
import gspread
from datetime import datetime
from google.oauth2 import service_account

# ═══════════════════════════════════════════════════════════════════════
# CONFIG (Populated securely via GitHub Secrets)
# ═══════════════════════════════════════════════════════════════════════

# Aggressively strip any accidental brackets, quotes, or spaces from the secret
API_URL = os.environ.get("API_URL", "").strip("[]\"' \t\n\r").rstrip("/")

if API_URL and not API_URL.startswith(("http://", "https://")):
    API_URL = f"https://{API_URL}"

API_USERNAME  = os.environ.get("API_USERNAME")
API_PASSWORD  = os.environ.get("API_PASSWORD")
GOOGLE_CREDS  = os.environ.get("GOOGLE_CREDS_JSON")

SPREADSHEET_ID = "1wWmKrRkGS9fTSYYBpZjODeVo4TBh-Gfrhf2p2Zg9QPU"
SHEET_TAB_NAME = "Export"

# ═══════════════════════════════════════════════════════════════════════
# VENDOR SHEETS
# ═══════════════════════════════════════════════════════════════════════

VENDOR_SHEETS = [
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
        "spreadsheet_id": "16sleNpG_CCVlwSPianDtl83IJTD2d6xQYXRX010MDgw",
        "tab_name":       "Export",
        "campaigns":      ["BorrowBetter"]
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

def fmt_contacted(val):
    try:
        if pd.isna(val) or str(val).strip() == "": return ""
        return "Yes" if int(float(val)) == 1 else "No"
    except: return ""

def build_output(df):
    log.info("Formatting output DataFrame...")
    out = pd.DataFrame()
    
    out["Created Date"]      = df["CreatedDate"].apply(fmt_date)
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
# GOOGLE SHEETS
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

    log.info(f"  Clearing and resizing '{tab_name}' to prevent 10M cell limit...")
    sheet.clear()
    try:
        sheet.resize(rows=len(data_to_upload), cols=len(columns))
    except Exception:
        pass 

    log.info(f"  Uploading {len(data_to_upload)} rows to '{tab_name}'...")
    sheet.update(range_name="A1", values=data_to_upload, value_input_option="USER_ENTERED")
    log.info(f"  Done writing to '{tab_name}'.")

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

        for vendor in VENDOR_SHEETS:
            log.info(f"Writing to vendor sheet: {vendor['tab_name']} ({vendor['campaigns']})...")
            upsert_to_sheet(client, output_df, vendor["spreadsheet_id"], vendor["tab_name"], MASTER_COLUMNS, vendor["campaigns"])

        log.info(f"Export completed: {datetime.now()}")

    except Exception as e:
        log.error(f"Export FAILED: {e}", exc_info=True)
        sys.exit(1)

if __name__ == "__main__":
    main()
