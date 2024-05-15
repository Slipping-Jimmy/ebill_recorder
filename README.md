# E-bill Recorder for Google Sheets

## Overview
This Python script automatically fetches electronic billing information from Gmail for orders made via Uber Eats and Foodpanda, and records this data into a specified Google Sheet. It is designed to help users automate the tracking of their food delivery expenses over time.

## Features
- Fetches email data using the Gmail API.
- Filters emails based on specific subjects related to Uber Eats and Foodpanda order confirmations.
- Extracts relevant billing information such as restaurant name, order date, and total cost.
- Records the extracted information into designated columns in a Google Sheet.

## Prerequisites
- Python 3.6 or higher
- Google account with access to Gmail and Google Sheets
- `google-auth`, `google-auth-oauthlib`, `google-auth-httplib2`, `google-api-python-client`, and `gspread` Python libraries.

## Setup Instructions

1. **Google API Console Setup:**
   - Create a new project in the Google Developers Console.
   - Enable the Gmail API and Google Sheets API.
   - Create credentials for an OAuth 2.0 client ID. Download the JSON file and rename it to `credentials.json`.
   - Save `credentials.json` in your project directory.

2. **Google Sheet Setup:**
   - Create a new Google Sheet and name it appropriately.
   - Share the sheet with the email address provided in your `credentials.json` to allow the script to access it.

3. **Python Environment:**
   - Install necessary Python libraries:
     ```
     pip install google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client gspread
     ```

4. **Token Generation:**
   - Run the script once to generate an authorization token:
     ```
     python script_name.py
     ```
   - Follow the on-screen instructions to log in with your Google account. This will create a `token.json` file in your project directory, which stores your access and refresh tokens.

## Usage
Run the script to fetch new e-bill records and update the Google Sheet:
```
python ebill_recorder_local.py
```

Adjust the `SHEET_NAME` and `WORKSHEET_NAME` variables in the script to match your Google Sheet configuration.

## Customization
You can customize the script by modifying the regular expressions used to extract data from emails or by changing the Google Sheet structure in the script settings.

## Important Notes
- Ensure you comply with all applicable laws and Google's terms of service when using this script.
- Handle sensitive data with care, especially when authorizing third-party access to your Gmail and Google Sheets.

---
