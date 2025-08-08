# Budget Tracker

A Google Sheets-based personal finance tool that automatically retrieves and categorizes your transactions using Plaid API integration and Gemini AI. Built on top of the free Google Sheets budget template from https://aspirebudget.com/.

## Features

- **Automatic Transaction Syncing**: Connect to your bank accounts and credit cards via Plaid to automatically import transactions
- **AI-Powered Categorization**: Uses Google's Gemini AI to intelligently categorize your transactions
- **Custom Categories**: Easily customize transaction categories to match your budgeting needs
- **Bank Account Management**: Support for multiple bank accounts and credit cards
- **Transaction Status Tracking**: Easily see which transactions are pending or cleared
- **Google Sheets Integration**: Built entirely in Google Sheets for easy access and customization

## Setup Instructions

### 1. Download the Template

Make a copy of my modified version of the Aspire Budgeting template from [here](https://docs.google.com/spreadsheets/d/1sTB5Cglan__ZVtzeAY1aHqnpr6ka3432XJYZ3n2FQOs/edit?usp=sharing).

### 2. API Keys Setup

Navigate to Extensions > Apps Script and open `code.gs`. Fill in the following constants at the top of the file:

```javascript
// Plaid Constants
const CLIENT_ID = ""; // Fill in Client ID from https://dashboard.plaid.com/developers/keys
const SECRET = "";    // Fill in Secret Key from https://dashboard.plaid.com/developers/keys

// Gemini API Constants
const GEMINI_API_KEY = ""; // Add your Gemini API key here
```

To obtain these keys:

- **Plaid API Keys**: Sign up at [Plaid Dashboard](https://dashboard.plaid.com/) and get your Client ID and Secret from the Developer section
- **Gemini API Key**: Get your API key from [Google AI Studio](https://makersuite.google.com/app/apikey)

### 3. Additional Configuration (Optional)

You can also customize these settings in `code.gs`:

```javascript
const TRANSACTION_DATA_START_DATE = "2025-01-01"; // Start date for transaction history
const PLAID_BASE_URL = "https://sandbox.plaid.com"; // Change to production URL when ready
```

### 4. Deploy the Script

1. Open your Google Sheet
2. Go to Extensions > Apps Script
3. Copy the contents of `code.gs` into the script editor
4. Create a new HTML file named `plaid_connect.html` and copy the contents from the project
5. Save the scripts and create a new deployment.

## Usage Instructions

### Connecting to Your Bank

1. In your Google Sheet, click on the newly created menu item **Budget Tracker**
2. Select **Connect to Bank** to begin the connection process
3. A dialog will appear with the Plaid interface
4. Follow the prompts to securely connect your bank account(s)
5. Once connected, your accounts will appear in the Configuration tab, and transactions will automatically start syncing to the Transactions tab

Alternatively, you can click the **Add Account** button in the Configuration tab.

### Managing Transactions

- **Sync Transactions**: Click on **Budget Tracker > Sync Transactions** to update your transaction data
- **Recategorize Transactions**: Click on **Budget Tracker > Recategorize Transactions** to apply AI categorization to uncategorized transactions
- **Manual Categorization**: You can manually change categories in the Transactions sheet

### Resetting the Spreadsheet

If you need to start over:

1. Click on **Budget Tracker > Reset Spreadsheet**
2. This will clear all transactions and remove all connected banks

## Security Note

This application uses Plaid's secure API to connect to your financial institutions. Your banking credentials are never stored in the Google Sheet or in the script properties. The application only stores access tokens provided by Plaid. However, Google may store the data sent to the Gemini API, so swap out for a different API with no data storage if you are concerned about privacy. The app only uses the memo of transactions for AI categorization (no amounts or account info) you can edit the prompt in code.gs to send different information.