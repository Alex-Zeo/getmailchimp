import requests, sys, os, time
import pandas as pd
from datetime import datetime
import configparser

# Initialize the parser and read the config file
config = configparser.ConfigParser()
config.read(r'C:\Users\research\Documents\MarketingMetrics\mailchimpconfig.ini')

# Retrieve configurations
# Check if config file exists
if not os.path.exists(r'C:\Users\research\Documents\MarketingMetrics\mailchimpconfig.ini'):
    raise FileNotFoundError("The configuration file 'config.ini' was not found.")
# Try to read configurations
try:
    datacenter = config.get('mailchimp', 'datacenter')
    api_key = config.get('mailchimp', 'api_key')
    folder_path = config.get('paths', 'folder_path')
except configparser.NoOptionError as e:
    raise Exception(f"Missing option in config file: {e}")


# Other variables
today = datetime.now().strftime("%Y%m%d")
excel_file_name = f'{folder_path}\\Mailchimp_data_{today}.xlsx'


# Set up logging to file
def log_print(*args, **kwargs):
    log_dir = os.path.join(folder_path, 'log')
    log_file_path = os.path.join(log_dir, f'maillog{today}.txt')

    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
        print(f"Created log directory at {log_dir}")

    original_stdout = sys.stdout
    with open(log_file_path, 'a') as f:
        sys.stdout = f
        print(*args, **kwargs)
    sys.stdout = original_stdout
    print(*args, **kwargs)

def get_campaigns(api_key, datacenter):
    log_print("Fetching all campaigns from Mailchimp...")
    campaigns = []
    count = 1000
    offset = 0
    while True:
        url = f"https://{datacenter}.api.mailchimp.com/3.0/campaigns?count={count}&offset={offset}"
        headers = {"Authorization": f"Bearer {api_key}"}
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            data = response.json()
            campaigns.extend(data['campaigns'])
            if len(data['campaigns']) < count:
                break
            offset += count
        else:
            log_print(f"Failed to fetch campaigns: {response.status_code}, {response.text}")
            break
    return {"campaigns": campaigns}

def get_campaign_details(api_key, datacenter, campaign_id):
    log_print(f"Fetching detailed data for campaign ID {campaign_id}")
    url = f"https://{datacenter}.api.mailchimp.com/3.0/reports/{campaign_id}"
    headers = {"Authorization": f"Bearer {api_key}"}
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        data = response.json()
        click_details = data.get('clicks', {})
        return {
            "campaign_title": data.get("campaign_title", "Not available"),
            "type": data.get("type", "Not available"),
            "list_id": data.get("list_id", "Not available"),
            "list_is_active": data.get("list_is_active", False),
            "list_name": data.get("list_name", "Not available"),
            "audience_name": data.get("list_name", "Not available"),  # Assuming 'list_name' is the 'audience_name'
            "subject_line": data.get("subject_line", "Not available"),
            "preview_text": data.get("preview_text", "Not available"),
            "emails_sent": data.get("emails_sent", 0),
            "abuse_reports": data.get("abuse_reports", 0),
            "unsubscribed": data.get("unsubscribed", 0),
            "send_time": data.get("send_time", "Not available"),
            "opens_total": data.get("opens", {}).get("opens_total", 0),
            "unique_opens": data.get("opens", {}).get("unique_opens", 0),
            "open_rate": data.get("opens", {}).get("open_rate", 0),
            "last_open": data.get("opens", {}).get("last_open", "Not available"),
            "total_clicks": click_details.get("clicks_total", 0),
            "unique_clicks": click_details.get("unique_clicks", 0),
            "click_rate": click_details.get("click_rate", 0),
            "last_click": click_details.get("last_click", "Not available")
        }
    else:
        log_print(f"Failed to fetch detailed data for campaign {campaign_id}: {response.status_code}, {response.text}")
        return {}

def get_list_activity(api_key, datacenter, list_id):
    log_print(f"Fetching list activity for list ID {list_id}")
    url = f"https://{datacenter}.api.mailchimp.com/3.0/lists/{list_id}/activity"
    headers = {"Authorization": f"Bearer {api_key}"}
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        activities = response.json()['activity']
        for activity in activities:
            if 'subs' in activity:
                return activity['subs']
        return 0
    else:
        log_print(f"Failed to fetch list activity for list {list_id}: {response.status_code}, {response.text}")
        return 0

def process_campaign_data(api_key, datacenter):
    log_print("Processing campaign data...")
    campaigns = get_campaigns(api_key, datacenter)
    data = []
    for campaign in campaigns.get('campaigns', []):
        campaign_id = campaign.get('id', 'Unknown')
        detailed_data = get_campaign_details(api_key, datacenter, campaign_id)

        campaign_data = {
            "Campaign ID": campaign_id,
            "Campaign Title": detailed_data.get("campaign_title", "Not available"),
            "Type": detailed_data.get("type", "Not available"),
            "List ID": detailed_data.get("list_id", "Not available"),
            "List is Active": detailed_data.get("list_is_active", False),
            "List Name": detailed_data.get("list_name", "Not available"),
            "Audience Name": detailed_data.get("audience_name", "Not available"),
            "Subject Line": detailed_data.get("subject_line", "Not available"),
            "Preview Text": detailed_data.get("preview_text", "Not available"),
            "Emails Sent": detailed_data.get("emails_sent", 0),
            "Abuse Reports": detailed_data.get("abuse_reports", 0),
            "Send Time": detailed_data.get("send_time", "Not available"),
            "Opens Total": detailed_data.get("opens_total", 0),
            "Unique Opens": detailed_data.get("unique_opens", 0),
            "Open Rate": detailed_data.get("open_rate", 0),
            "Last Open": detailed_data.get("last_open", "Not available"),
            "Total Clicks": detailed_data.get("total_clicks", 0),
            "Unique Clicks": detailed_data.get("unique_clicks", 0),
            "Click Rate": detailed_data.get("click_rate", 0),
            "Last Click": detailed_data.get("last_click", "Not available"),
        }
        data.append(campaign_data)

    df = pd.DataFrame(data)
    log_print("Data processing complete.")
    try:
        df.to_excel(excel_file_name, index=False)
        log_print(f"DataFrame saved successfully to {excel_file_name}.")
    except Exception as e:
        log_print(f"Failed to save DataFrame: {e}")
    return df

# Usage
processed_data = process_campaign_data(api_key, datacenter)
log_print(processed_data.head())
log_print(f"Data exported to Excel successfully as {excel_file_name}.")
