**# getmailchimp**
This Python script is designed to automate the extraction and processing of marketing campaign data from Mailchimp. By leveraging the Mailchimp API, the script retrieves comprehensive information about all campaigns, including details such as campaign titles, types, audience information, performance metrics like open rates and click rates, and more. The requests library facilitates seamless communication with the Mailchimp API, while pandas organizes the fetched data into a structured Excel file for easy analysis and reporting.

![image](https://github.com/user-attachments/assets/2d8ba794-1d97-4f75-837c-46ae07913a84)

Configuration management is handled through a dedicated config.ini file, which stores essential parameters such as the Mailchimp datacenter, API key, and the folder path where data and logs will be stored. The script begins by verifying the existence of this configuration file and then securely reads the necessary configurations. Robust error handling ensures that any missing configuration options or issues with file access are promptly reported, preventing the script from running with incomplete or incorrect settings. Additionally, a custom logging function log_print is implemented to record the script’s operations and any encountered errors into a log file, providing transparency and facilitating troubleshooting.

![image](https://github.com/user-attachments/assets/1249cef8-c2ed-44ee-b286-f7b4eece842d)


The core functionality of the script revolves around fetching campaign data, processing it, and exporting the results to an Excel file named with the current date for easy reference. The process_campaign_data function orchestrates the entire workflow by first retrieving all campaigns, then obtaining detailed information for each campaign, and finally compiling this data into a pandas DataFrame. Once the data is processed, it is saved to an Excel file in a sharepoint folder. This sharepoint folder enables data refreshes of our Tableau workbook. Throughout this process, the script logs each significant step, including the creation of directories, data fetching statuses, and the successful saving of the Excel file. This automation not only streamlines the data collection process but also ensures that marketing teams have timely and accurate insights into their campaign performance.

**#getmailchimpsubs**
