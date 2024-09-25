import mailchimp_marketing as MailchimpMarketing
from mailchimp_marketing.api_client import ApiClientError
import pandas as pd
from datetime import datetime
import os, configparser

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

# Setup basic variables
today = datetime.now().strftime("%Y%m%d")
excel_file_name = f'{folder_path}\\Mailchimp_subs_stats.xlsx'
list_id = '039e1217a3'  # Replace with your actual list ID

# Setup Mailchimp client
print("Initializing Mailchimp client...")
client = MailchimpMarketing.Client()
client.set_config({
  "api_key": api_key,
  "server": datacenter
})
print("Mailchimp client initialized successfully.")

def get_all_contacts(list_id):
    print("Starting to fetch contacts all contacts in eNews...")
    all_contacts = []
    try:
        # Define date range for 2024
        #start_date = datetime(2024, 1, 1).isoformat()
        #end_date = datetime(2024, 12, 31).isoformat()
        #print(f"Date range set from {start_date} to {end_date}")

        offset = 0
        count = 100  # Fetch 100 contacts per call
        total_contacts_retrieved = 0
        
        while True:
            print(f"Fetching contacts with offset {offset} and count {count}...")
            contacts = client.lists.get_list_members_info(list_id=list_id, count=count, offset=offset)
            
            num_fetched = len(contacts['members'])
            total_contacts_retrieved += num_fetched
            print(f"Retrieved {num_fetched} contacts. Total contacts retrieved so far: {total_contacts_retrieved}")
            
            for contact in contacts['members']:
                signup_timestamp = contact.get('timestamp_signup', 'N/A')
                optin_timestamp = contact.get('timestamp_opt', 'N/A')
                status = contact.get('status', 'N/A')
                last_changed = contact.get('last_changed', 'N/A')

                all_contacts.append({
                    'email': contact['email_address'],
                    'timestamp_signup': signup_timestamp,
                    'timestamp_opt': optin_timestamp,
                    'status': status,  # Include the status field
                    'last_changed': last_changed
                })
                print(f"Contact added - Email: {contact['email_address']}, Signup: {signup_timestamp}, Opt-in: {optin_timestamp}, Status: {status}")
            
            if len(contacts['members']) < count:
                print("No more contacts to fetch. Exiting the loop.")
                break  # No more pages to fetch
            
            offset += count
            print(f"Moving to next batch with offset {offset}.")

    except ApiClientError as error:
        print(f"Error occurred while fetching contacts: {error.text}")
    
    print(f"Finished fetching contacts. Total contacts retrieved: {len(all_contacts)}")
    return all_contacts

def generate_excel_file(contacts, file_path):
    print("Starting to generate Excel file...")
    if contacts:
        # Convert contacts to DataFrame
        df = pd.DataFrame(contacts)
        print(f"Dataframe created with {len(df)} rows.")
        
        # Saving DataFrame to Excel
        df.to_excel(file_path, index=False)
        print(f"Excel file generated successfully at {file_path}")
    else:
        print("No contacts data to write to Excel.")

# Example usage
print("Fetching contacts for list ID:", list_id)
contacts = get_all_contacts(list_id)

if contacts:
    print("Generating Excel report...")
    generate_excel_file(contacts, excel_file_name)
else:
    print("No contacts found. Skipping Excel file generation.")