// Plaid Constants
const CLIENT_ID = ""; // Fill in Client ID and Secret Key from https://dashboard.plaid.com/developers/keys
const SECRET = "";
const TRANSACTION_DATA_START_DATE = "2025-01-01"; // Start date to pull transactions from when connecting to a new bank
const CLIENT_NAME = "Budget Tracker";
const CLIENT_USER_ID = "1";
const PLAID_BASE_URL = "https://sandbox.plaid.com"; // Change to 'production' as needed
const PLAID_PRODUCTS = ["transactions"];
const PLAID_COUNTRY_CODES = "US";


// Spreadsheet Constants
const CONFIGURATION_SHEET_ID = "1381484621";
const ACCOUNTS_RANGE = "H6:I20"; // Has 2 columns, first one is for Bank Accounts/Cash, second one is for Credit Cards

const TRANSACTIONS_SHEET_ID = "925237105";
const TRANSACTIONS_RANGE = "B8:K3000"; // Includes two hidden columns: Transaction ID and Categorization Source
const TRANSACTION_COLUMNS = ["Date", "Outflow", "Inflow", "Category", "Account", "Memo", "Status"];
const CATEGORIES_RANGE = "E8:E3000";
const CATEGORY_OPTIONS_RANGE = "P9:P157";
const CLEARED_STATUS = "✅";
const PENDING_STATUS = "🅿️";

// Gemini API Constants
const GEMINI_API_KEY = ""; // Add your Gemini API key here
const GEMINI_API_URL = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash-lite:generateContent"; // Using Gemini 2.5 Flash Lite for speed

// Configuration for AI categorization:

// PRIVACY WARNING: Including the transaction memos will result in more accurate categorization, but Google may retain this
// information for training purposes. Only enable if you're comfortable with Google storing your transaction memos.
const INCLUDE_MEMO_IN_PROMPT = false; // By default, will only include Plaid's category guess in the prompt sent to Gemini. If true, will also include the transaction memo.


/**
 * Creates a menu entry in the Google Sheets UI when the document is opened.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Budget Tracker')
    .addItem('Connect to Bank', 'showPlaidConnectDialog')
    .addItem('Sync Transactions', 'syncTransactions')
    .addItem('Recategorize Transactions', 'recategorizeAllTransactions')
    .addSeparator()
    .addItem('Reset Spreadsheet', 'confirmResetSpreadsheet')
    .addToUi();
}

/**
 * Handles edits to the spreadsheet and marks category changes as user-categorized.
 * This function is triggered automatically by Google Sheets when any edit is made.
 * 
 * @param {Object} e - The event object containing information about the edit.
 */
function onEdit(e) {
  try {
    const editedRange = e.range;
    const sheet = editedRange.getSheet();
    // Check if this is the transactions sheet
    if (sheet.getSheetId() != TRANSACTIONS_SHEET_ID) {
      return; // Exit if not in the transactions sheet
    }
    
    // Check if the accounts dropdown was edited and update the filter
    if (editedRange.getRow() === 2 && editedRange.getColumn() === 7) { // Dropdown is in G2
      const accountSelection = editedRange.getValue();
      
      // Get the existing filter
      const filter = sheet.getFilter();
      
      if (filter) {
        const accountColumn = 6; // Column F is the Account column
        
        if (accountSelection && accountSelection !== 'All Accounts') {
          // Create filter criteria that matches the selected account
          const criteria = SpreadsheetApp.newFilterCriteria()
            .whenFormulaSatisfied(`=OR(F8="${accountSelection}", F8="")`)
            .build();
          
          filter.setColumnFilterCriteria(accountColumn, criteria);
        } else {
          // Clear the filter criteria for the account column when 'All Accounts' is selected
          filter.removeColumnFilterCriteria(accountColumn);
        }
      } else {
        Logger.log("No filter found on the sheet. A filter should be pre-created on the transactions range.");
      }
      return;
    }
    
    // Set categorization source to "USER" for manually categorized transactions
    const categoriesRange = sheet.getRange(CATEGORIES_RANGE);
    // Get the range of the edited cells that intersects with the categories range
    const FIRSTROW = Math.max(editedRange.getRow(), categoriesRange.getRow());
    const LASTROW = Math.min(editedRange.getLastRow(), categoriesRange.getLastRow());
    const FIRSTCOL = Math.max(editedRange.getColumn(), categoriesRange.getColumn());
    const LASTCOL = Math.min(editedRange.getLastColumn(), categoriesRange.getLastColumn());

    if (FIRSTROW > LASTROW || FIRSTCOL > LASTCOL) return; // Exit if no categories were edited
    
    // For each edited row, set the corresponding source cell to "USER"
    for (let row = FIRSTROW; row <= LASTROW; row++) {
      const category = sheet.getRange(row, FIRSTCOL).getValue();
      if (!category) continue;
      
      // Set the source column to "USER"
      const sourceCell = sheet.getRange(row, 11); // Categorization source is the 11th column
      sourceCell.setValue("USER");
    }
  } catch (error) {
    Logger.log("Error in onEdit: " + error.message);
  }
}

/**
 * Opens a new tab with the Plaid Link interface for connecting to a bank.
 */
function showPlaidConnectDialog() {
  const html = HtmlService.createHtmlOutputFromFile('plaid_connect')
    .setTitle('Connect to Bank')
    .setWidth(400)
    .setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(html, 'Connect to Bank');
}

/**
 * Initiates the Plaid Link flow to connect to a bank.
 * This function creates a link token and returns it to the client-side HTML.
 * @return {Object} Object containing the link token or an error message.
 */
function getLinkToken() {
  try {
    const url = `${PLAID_BASE_URL}/link/token/create`;
    const payload = {
      client_id: CLIENT_ID,
      secret: SECRET,
      client_name: CLIENT_NAME,
      user: {
        client_user_id: CLIENT_USER_ID
      },
      products: PLAID_PRODUCTS,
      country_codes: [PLAID_COUNTRY_CODES],
      language: "en",
      transactions: {
        days_requested: 365 // request a full year of transactions
      }
    };
    
    const options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload)
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const responseData = JSON.parse(response.getContentText());
    
    return { success: true, link_token: responseData.link_token };
  } catch (error) {
    Logger.log(`Error creating link token: ${error}`);
    return { success: false, error: error.toString() };
  }
}

/**
 * Exchanges a public token for an access token and item ID.
 * @param {string} publicToken - The public token received from Plaid Link.
 * @param {Object} metadata - Metadata from Plaid Link including institution information.
 * @return {Object} Object containing success status, access token and item ID or error.
 */
function exchangePublicToken(publicToken, metadata) {
  try {
    // Check if the institution is already linked
    const institutionId = metadata && metadata.institution ? metadata.institution.institution_id : "unknown";
    const institutionName = getInstitutionName(metadata);
    
    // Get existing banks list
    const scriptProperties = PropertiesService.getScriptProperties();
    const existingBanksJson = scriptProperties.getProperty('PLAID_BANKS_LIST');
    
    if (existingBanksJson) {
      try {
        const banksList = JSON.parse(existingBanksJson);
        
        // Check if this institution already exists
        const existingBank = banksList.find(bank => bank.id === institutionId);
        if (existingBank) {
          Logger.log(`Institution ${institutionName} (${institutionId}) is already connected.`);
          return { 
            success: false, 
            error: `${institutionName} is already connected to the budget tracker.` 
          };
        }
      } catch (e) {
        Logger.log('Error parsing existing banks list: ' + e);
      }
    }
    
    // Exchange the public token for an access token
    const url = `${PLAID_BASE_URL}/item/public_token/exchange`;
    const payload = {
      client_id: CLIENT_ID,
      secret: SECRET,
      public_token: publicToken
    };
    
    const options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload)
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const responseData = JSON.parse(response.getContentText());
    
    Logger.log("Successfully exchanged public token for access token. Response: " + JSON.stringify(responseData));

    return { 
      success: true, 
      access_token: responseData.access_token,
      item_id: responseData.item_id
    };
  } catch (error) {
    Logger.log(`Error exchanging public token: ${error}`);
    return { success: false, error: error.toString() };
  }
}

/**
 * Removes a Plaid Item using the access token.
 * This prevents Plaid from continuing to charge for the connection.
 * @param {string} accessToken - The access token for the item to remove.
 * @return {Object} Object containing success status and any error message.
 */
function removePlaidItem(accessToken) {
  try {
    if (!accessToken) {
      Logger.log('No access token provided to removePlaidItem');
      return { success: false, error: 'No access token provided' };
    }
    
    const url = `${PLAID_BASE_URL}/item/remove`;
    const payload = {
      client_id: CLIENT_ID,
      secret: SECRET,
      access_token: accessToken
    };
    
    const options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const responseData = JSON.parse(response.getContentText());
    
    if (response.getResponseCode() === 200) {
      Logger.log(`Successfully removed Plaid item with access token: ${accessToken}. Request ID: ${responseData.request_id}`);
      return { success: true };
    } else {
      Logger.log(`Error removing Plaid item with access token: ${accessToken}:\n${JSON.stringify(responseData)}`);
      return { success: false, error: responseData.error_message || 'Unknown error' };
    }
  } catch (error) {
    Logger.log(`Exception in removePlaidItem: ${error}`);
    return { success: false, error: error.toString() };
  }
}

/**
 * Connects a bank after token exchange and starts syncing transactions.
 * @param {string} accessToken - The access token received from Plaid.
 * @param {string} itemId - The item ID received from Plaid.
 * @param {Object} metadata - Metadata from Plaid Link including institution information.
 * @return {boolean} True if the bank was successfully connected and transactions synced, false otherwise.
 */
function connectBankAndSync(accessToken, itemId, metadata) {
  try {
    // Get institution information
    const institutionName = getInstitutionName(metadata);
    const institutionId = metadata && metadata.institution ? metadata.institution.institution_id : "unknown";
    
    // Get existing banks list or initialize a new one
    const scriptProperties = PropertiesService.getScriptProperties();
    let banksList = [];
    const existingBanksJson = scriptProperties.getProperty('PLAID_BANKS_LIST');
    if (existingBanksJson) {
      try {
        banksList = JSON.parse(existingBanksJson);
      } catch (e) {
        Logger.log('Error parsing existing banks list, initializing new one');
      }
    }
    
    // Create a bank entry with all necessary information
    const bankEntry = {
      name: institutionName,
      id: institutionId,
      access_token: accessToken,
      item_id: itemId,
      date_added: new Date().toISOString(),
      cursor: null // Initialize cursor as null
    };
    banksList.push(bankEntry);
    
    // Save the updated banks list with all bank info
    scriptProperties.setProperty('PLAID_BANKS_LIST', JSON.stringify(banksList));
    
    // Add the accounts to the configuration sheet
    addAccountsToConfigSheet(metadata);
    
    // Immediately fetch transactions after successful connection
    const transactionResult = syncTransactions();
    
    return transactionResult.success;
  } catch (error) {
    Logger.log(`Error connecting bank: ${error}`);
    return false;
  }
}

/**
 * Extracts the institution name from metadata and formats for the Configuration sheet.
 * @param {Object} metadata - Metadata from Plaid Link including institution information.
 * @return {string} The institution name or "Unknown Bank" if not found.
 */
function getInstitutionName(metadata) {
  let institutionName = metadata.institution ? metadata.institution.name : "Unknown Bank";
  // Ensure no dash in institution name (will mess up formatting on config sheet)
  const dashIndex = institutionName.indexOf(" - ");
  if (dashIndex !== -1) {
    institutionName = institutionName.substring(0, dashIndex);
  }
  return institutionName;
}

/**
 * Adds accounts to the configuration sheet in the accounts range.
 * @param {Object} metadata - Metadata from Plaid Link including institution and accounts information.
 */
function addAccountsToConfigSheet(metadata) {
  try {
    // If no metadata or accounts, use institution name as fallback
    const institutionName = getInstitutionName(metadata);
    if (!metadata || !metadata.accounts || metadata.accounts.length === 0) {
      addSingleAccountToConfigSheet(institutionName, false); // Default to bank account column if no account type info is available
      return;
    }
    
    // Add each account with institution name prefix
    metadata.accounts.forEach(account => {
      const accountName = `${institutionName} - ${account.name}`;
      // Determine if this is a credit card based on account type and subtype
      const isCreditCard = (account.type === "credit" && account.subtype === "credit card");
      addSingleAccountToConfigSheet(accountName, isCreditCard);
    });
  } catch (error) {
    Logger.log(`Error adding accounts to config sheet: ${error}`);
  }
}

/**
 * Adds a single account name to the configuration sheet.
 * @param {string} accountName - The name of the account to add.
 * @param {boolean} isCreditCard - Whether this account is a credit card (true) or bank account/cash (false).
 */
function addSingleAccountToConfigSheet(accountName, isCreditCard) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetById(CONFIGURATION_SHEET_ID);
    
    // Get the accounts range
    const accountsRange = configSheet.getRange(ACCOUNTS_RANGE);
    const accountsValues = accountsRange.getValues();
    
    // Determine which column to use (1 for bank accounts, 2 for credit cards)
    const columnIndex = isCreditCard ? 1 : 0;
    
    // Check if this account already exists in the appropriate column
    for (let i = 0; i < accountsValues.length; i++) {
      if (accountsValues[i][columnIndex] === accountName) {
        Logger.log(`Account ${accountName} already exists in column ${columnIndex + 1}, skipping.`);
        return; // If account already exists, don't add it again
      }
    }
    
    // Find the first empty cell in the appropriate column
    let emptyRowIndex = -1;
    for (let i = 0; i < accountsValues.length; i++) {
      if (!accountsValues[i][columnIndex]) {
        emptyRowIndex = i;
        break;
      }
    }
    
    // If we found an empty cell, add the account name
    if (emptyRowIndex !== -1) {
      // Column is 0-based in the array but 1-based in the getCell method
      const cell = accountsRange.getCell(emptyRowIndex + 1, columnIndex + 1);
      cell.setValue(accountName);
    } else {
      const errorMessage = `No empty cells found in accounts range to add the ${isCreditCard ? 'credit card' : 'bank account'} ${accountName}`;
      Logger.log(errorMessage);
      
      // Show an alert to the user
      const ui = SpreadsheetApp.getUi();
      ui.alert('Error Adding Account', errorMessage + '\n\nPlease add more rows to the accounts range in the configuration sheet.', ui.ButtonSet.OK);
    }
  } catch (error) {
    const errorMessage = `Error adding account to config sheet: ${error}`;
    Logger.log(errorMessage);
    
    // Show an alert to the user
    const ui = SpreadsheetApp.getUi();
    ui.alert('Error Adding Account', errorMessage, ui.ButtonSet.OK);
  }
}

/**
 * Synchronizes the banks in script properties with those in the configuration sheet.
 * Removes banks from script properties that are no longer in the configuration sheet.
 * @return {Object} Object containing arrays of active banks and removed banks. The banks are formatted the same as PLAID_BANKS_LIST in script properties.
 */
function syncBanksWithConfigSheet() {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    const banksListJson = scriptProperties.getProperty('PLAID_BANKS_LIST');
    
    // If no banks are stored, nothing to sync
    if (!banksListJson) {
      return { activeBanks: [], removedBanks: [] };
    }
    
    // Get the current banks list from script properties
    const banksList = JSON.parse(banksListJson);
    if (!banksList || banksList.length === 0) {
      return { activeBanks: [], removedBanks: [] };
    }
    
    // Get the banks from the configuration sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetById(CONFIGURATION_SHEET_ID);
    const accountsRange = configSheet.getRange(ACCOUNTS_RANGE);
    const accountsValues = accountsRange.getValues();
    
    // Create a set of institution names from the configuration sheet
    // Extract institution names from account names (e.g., "Chase - Plaid Checking" -> "Chase")
    const configSheetInstitutions = new Set();
    
    // Check both columns (bank accounts and credit cards)
    accountsValues.forEach(row => {
      // Check bank accounts column (first column)
      if (row[0]) {
        extractAndAddInstitution(row[0], configSheetInstitutions);
      }
      
      // Check credit cards column (second column)
      if (row[1]) {
        extractAndAddInstitution(row[1], configSheetInstitutions);
      }
    });
    
    // Filter the banks list to keep only those still in the configuration sheet
    const activeBanks = [];
    const removedBanks = [];
    
    banksList.forEach(bank => {
      if (configSheetInstitutions.has(bank.name)) {
        activeBanks.push(bank);
      } else {
        removedBanks.push(bank);
        // Safely remove the item so Plaid does not continue to charge for it
        const result = removePlaidItem(bank.access_token);
        if (!result.success) {
          return { activeBanks: [], removedBanks: [], error: result.error };
        }
        Logger.log(`Removed bank: ${bank.name} since it was missing from the config sheet.`);
      }
    });
    
    // Update the banks list in script properties if any were removed
    if (removedBanks.length > 0) {
      scriptProperties.setProperty('PLAID_BANKS_LIST', JSON.stringify(activeBanks));
    }
    
    return { activeBanks, removedBanks };
  } catch (error) {
    Logger.log(`Error syncing banks with config sheet: ${error}`);
    return { activeBanks: [], removedBanks: [], error: error.toString() };
  }
}

/**
 * Helper function to extract institution name from an account name and add it to the set.
 * @param {string} accountName - The full account name (e.g., "Chase - Checking")
 * @param {Set} institutionsSet - The set to add the institution name to
 */
function extractAndAddInstitution(accountName, institutionsSet) {
  const dashIndex = accountName.indexOf(' - ');
  if (dashIndex !== -1) {
    const institutionName = accountName.substring(0, dashIndex);
    institutionsSet.add(institutionName);
  } else {
    // If there's no dash, use the whole name as institution name
    institutionsSet.add(accountName);
  }
}

/**
 * Wrapper function for syncPlaidTransactions that displays the result to the user.
 * This is the function that should be called from the menu.
 */
function syncTransactions() {
  const ui = SpreadsheetApp.getUi();
  const result = syncPlaidTransactions();
  
  if (result.success) {
    ui.alert('Success', result.message, ui.ButtonSet.OK);
  } else {
    ui.alert('Error', result.message, ui.ButtonSet.OK);
  }
  
  return result;
}

/**
 * Shows a confirmation dialog before resetting the spreadsheet.
 */
function confirmResetSpreadsheet() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Reset Spreadsheet',
    'This will clear all transactions and remove all connected banks. This action cannot be undone. Are you sure you want to proceed?',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    const result = resetSpreadsheet();
    ui.alert(result.success ? 'Success' : 'Error', result.message, ui.ButtonSet.OK);
  }
}

/**
 * Resets the spreadsheet by clearing all transactions and removing all banks.
 * @return {Object} Object containing success status and message.
 */
function resetSpreadsheet() {
  try {
    // Delete the PLAID_BANKS_LIST property and remove all Plaid items
    const scriptProperties = PropertiesService.getScriptProperties();
    if (scriptProperties.getProperty('PLAID_BANKS_LIST')) {
      const banksList = JSON.parse(scriptProperties.getProperty('PLAID_BANKS_LIST'));
      
      // Remove each Plaid item to prevent Plaid from continuing to charge for the connections
      for (const bank of banksList) {
        if (bank.access_token) {
          const result = removePlaidItem(bank.access_token);
          if (result.success) {
            Logger.log(`Successfully removed Plaid item for institution: ${bank.name || 'Unknown'}`);
          } else {
            Logger.log(`Failed to remove Plaid item for institution: ${bank.name || 'Unknown'}, Error: ${result.error}`);
            return { success: false, message: `Failed to remove Plaid item for institution: ${bank.name || 'Unknown'}, Error: ${result.error}` };
          }
        }
      }

      scriptProperties.deleteProperty('PLAID_BANKS_LIST');
      Logger.log('Deleted PLAID_BANKS_LIST from script properties');
    }

    // Clear all transactions from the sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetById(TRANSACTIONS_SHEET_ID);
    const range = sheet.getRange(TRANSACTIONS_RANGE);
    range.clearContent();
    Logger.log('Cleared all transactions from the sheet');
    
    // Clear all banks from the configuration sheet
    const configSheet = ss.getSheetById(CONFIGURATION_SHEET_ID);
    const banksRange = configSheet.getRange(ACCOUNTS_RANGE);
    banksRange.clearContent();
    Logger.log('Cleared all banks from the configuration sheet');
    
    return {
      success: true,
      message: 'Successfully reset the spreadsheet. All transactions have been cleared and all banks have been removed.'
    };
  } catch (error) {
    Logger.log(`Error resetting spreadsheet: ${error}`);
    return { success: false, message: `Error: ${error.toString()}` };
  }
}

/**
 * Fetches transactions from all connected banks and writes them to the sheet.
 * @return {Object} Object containing success status and message.
 */
function syncPlaidTransactions() {
  try {
    // Check if a sync task is already running
    const scriptProperties = PropertiesService.getScriptProperties();
    const syncRunning = scriptProperties.getProperty('SYNC_TASK_RUNNING');
    
    if (syncRunning === 'true') {
      Logger.log('Sync already in progress. Aborting this execution.');
      return { success: false, message: "Another sync task is already in progress. Please try again later." };
    }
    
    // Set the lock flag
    scriptProperties.setProperty('SYNC_TASK_RUNNING', 'true');
    
    // Protect TRANSACTIONS_RANGE and ACCOUNTS_RANGE to prevent edits during sync
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const transactionsSheet = ss.getSheetById(TRANSACTIONS_SHEET_ID);
    const transactionsRange = transactionsSheet.getRange(TRANSACTIONS_RANGE);
    transactionsRange.protect().setDescription('Sync in progress').setWarningOnly(true);
    
    const configSheet = ss.getSheetById(CONFIGURATION_SHEET_ID);
    const accountsRange = configSheet.getRange(ACCOUNTS_RANGE);
    accountsRange.protect().setDescription('Sync in progress').setWarningOnly(true);
    
    // Sync the banks list with the configuration sheet
    const syncResult = syncBanksWithConfigSheet();
    if (syncResult.error) {
      return { success: false, message: syncResult.error };
    }
    
    let removedBanksMessage = "";
    if (syncResult.removedBanks && syncResult.removedBanks.length > 0) {
      const removedNames = syncResult.removedBanks.map(bank => bank.name).join(", ");
      removedBanksMessage = `Removed ${syncResult.removedBanks.length} banks no longer in configuration sheet: ${removedNames}.\n`;
    }
    
    // Get the active banks list
    const activeBanks = syncResult.activeBanks;
    if (!activeBanks || activeBanks.length === 0) {
      return { success: false, message: "No active banks found. Please connect to a bank first." };
    }
    
    // Prepare to collect transactions
    let addedTransactions = [];
    let modifiedTransactions = [];
    let removedTransactionIds = [];
    let allAccounts = [];
    let successfulBanks = [];
    let failedBanks = [];
    
    // Process each active bank
    for (const bank of activeBanks) {
      try {
        // The bank object from activeBanks already contains all the information we need since it comes from PLAID_BANKS_LIST
        const accessToken = bank.access_token;
        
        // Use the cursor from the bank object if it exists
        let cursor = bank.cursor || null;
        
        // Fetch transactions for this bank with cursor-based pagination
        const url = `${PLAID_BASE_URL}/transactions/sync`;
        let hasMoreTransactions = true;
        let bankAddedTransactions = [];
        let bankModifiedTransactions = [];
        let bankRemovedTransactionIds = [];
        let bankAccounts = [];
        
        // Log the start of transaction fetching for this bank
        Logger.log(`Fetching transactions for ${bank.name}`);
        
        // Loop to handle pagination with cursor
        while (hasMoreTransactions) {
          const payload = {
            client_id: CLIENT_ID,
            secret: SECRET,
            access_token: accessToken,
            cursor: cursor
          };
          
          const options = {
            method: "post",
            contentType: "application/json",
            payload: JSON.stringify(payload)
          };
          
          const response = UrlFetchApp.fetch(url, options);
          const responseData = JSON.parse(response.getContentText());
          
          // Update the cursor for the next request
          cursor = responseData.next_cursor;
          
          // Process added transactions
          if (responseData.added && responseData.added.length > 0) {
            // Add institution name to each transaction for better identification
            responseData.added.forEach(transaction => {
              transaction.institution_name = bank.name;
            });
            
            // Add this batch to our bank's added transactions
            bankAddedTransactions = bankAddedTransactions.concat(responseData.added);

            Logger.log("Added transactions:\n" + JSON.stringify(responseData.added));
          }
          
          // Process modified transactions
          if (responseData.modified && responseData.modified.length > 0) {
            // Add institution name to each transaction for better identification
            responseData.modified.forEach(transaction => {
              transaction.institution_name = bank.name;
            });
            
            // Add this batch to our bank's modified transactions
            bankModifiedTransactions = bankModifiedTransactions.concat(responseData.modified);

            Logger.log("Modified transactions:\n" + JSON.stringify(responseData.modified));
          }
          
          // Process removed transactions
          if (responseData.removed && responseData.removed.length > 0) {
            // Add this batch to our bank's removed transaction IDs
            const removedIds = responseData.removed.map(item => item.transaction_id);
            bankRemovedTransactionIds = bankRemovedTransactionIds.concat(removedIds);

            Logger.log("Removed transactions:\n" + JSON.stringify(responseData.removed));
          }
          
          // Add accounts if present
          if (responseData.accounts && responseData.accounts.length > 0) {
            bankAccounts = bankAccounts.concat(responseData.accounts);
          }
          
          // Check if we need to fetch more transactions
          hasMoreTransactions = responseData.has_more;
          
          // Save the cursor after each successful request
          if (cursor) {       
            // Get the banks list from script properties to update
            const scriptProperties = PropertiesService.getScriptProperties();
            let banksList = [];
            const existingBanksJson = scriptProperties.getProperty('PLAID_BANKS_LIST');
            if (existingBanksJson) {
              try {
                banksList = JSON.parse(existingBanksJson);
                
                // Find and update the bank entry with the new cursor
                const bankIndex = banksList.findIndex(b => b.id === bank.id);
                if (bankIndex >= 0) {
                  banksList[bankIndex].cursor = cursor;
                  
                  // Save the updated banks list
                  scriptProperties.setProperty('PLAID_BANKS_LIST', JSON.stringify(banksList));
                }
              } catch (e) {
                Logger.log('Error updating cursor in banks list: ' + e);
              }
            }
          }
          
          // Add a small delay to avoid hitting rate limits
          Utilities.sleep(100);
        }
        
        // Log the total number of transactions fetched for this bank
        Logger.log(`Fetched transactions for ${bank.name}. ${bankAddedTransactions.length} added, ${bankModifiedTransactions.length} modified, and ${bankRemovedTransactionIds.length} removed.`);
        
        // Add this bank's transactions and accounts to our overall collection
        if (bankAddedTransactions.length > 0 || bankModifiedTransactions.length > 0 || bankRemovedTransactionIds.length > 0) {
          // Filter added transactions to only include those on or after TRANSACTION_DATA_START_DATE
          const startDate = new Date(TRANSACTION_DATA_START_DATE);
          
          // Filter added transactions
          const filteredAddedTransactions = bankAddedTransactions.filter(transaction => {
            const transactionDate = new Date(transaction.date);
            return transactionDate >= startDate;
          });
          
          // Log how many transactions were filtered out
          const numFiltered = bankAddedTransactions.length - filteredAddedTransactions.length;
          if (numFiltered > 0) {
            Logger.log(`Filtered out ${numFiltered} new transactions which occurred before ${TRANSACTION_DATA_START_DATE}`);
          }
          
          // Add transactions from this batch to our collections
          addedTransactions = addedTransactions.concat(filteredAddedTransactions);
          modifiedTransactions = modifiedTransactions.concat(bankModifiedTransactions);
          removedTransactionIds = removedTransactionIds.concat(bankRemovedTransactionIds);
          allAccounts = allAccounts.concat(bankAccounts);
          successfulBanks.push(bank.name);
        }
      } catch (bankError) {
        Logger.log(`Error fetching transactions for bank ${bank.name}: ${bankError}`);
        failedBanks.push(bank.name);
      }
    }
    
    // Process the transactions and combine with existing transactions
    if (addedTransactions.length > 0 || modifiedTransactions.length > 0 || removedTransactionIds.length > 0) {
      // Process transactions with the sync endpoint results
      let updatedTransactions = getUpdatedTransactions(
        addedTransactions, 
        modifiedTransactions, 
        removedTransactionIds, 
        allAccounts
      );
      
      // Write all transactions to the sheet
      if (updatedTransactions.length > 0) {
        const result = writeTransactionsToSheet(updatedTransactions);
        
        if (result.success) {
          const addedCount = addedTransactions.length;
          const modifiedCount = modifiedTransactions.length;
          const removedCount = removedTransactionIds.length;
          
          return {
            success: true, 
            message: `Successfully synced transactions from ${successfulBanks.length} banks: ${successfulBanks.join(", ")}\n` +
                    `${removedBanksMessage}` +
                    `Added: ${addedCount}, Modified: ${modifiedCount}, Removed: ${removedCount}. ` +
                    `${failedBanks.length > 0 ? `Failed to sync ${failedBanks.length} banks: ${failedBanks.join(", ")}` : ''}` 
          };
        } else {
          return result; // Pass through any errors from writeTransactionsToSheet
        }
      } else {
        return {
          success: true, 
          message: `${removedBanksMessage}No transactions found after processing.` 
        };
      }
    } else {
      return {
        success: true, 
        message: `${removedBanksMessage}No new transactions found.` 
      };
    }
  } catch (error) {
    Logger.log(`Error syncing transactions: ${error}`);
    return { success: false, message: `Error: ${error.toString()}` };
  } finally {
    // Always clear the lock flag when done, even if there was an error
    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperty('SYNC_TASK_RUNNING', 'false');
    
    // Unprotect the TRANSACTIONS_RANGE when sync is complete
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetById(TRANSACTIONS_SHEET_ID);
      const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
      
      // Find and remove the protection we added
      for (let i = 0; i < protections.length; i++) {
        const protection = protections[i];
        if (protection.getDescription() === 'Sync in progress') {
          protection.remove();
          break;
        }
      }
    } catch (unprotectError) {
      Logger.log(`Error removing protection: ${unprotectError}`);
    }
  }
}

/**
 * Formats Plaid transactions into the sheet format.
 * @param {Array} transactions - Array of transaction objects from Plaid.
 * @param {Array} accounts - Array of account objects from Plaid.
 * @return {Array} Array of formatted transaction rows ready for the sheet.
 */
function formatTransactionsForSheet(transactions, accounts) {
  // Create a map of account IDs to account names for easier lookup
  const accountMap = {};
  accounts.forEach(account => {
    accountMap[account.account_id] = account.name;
  });
  
  // Prepare the data for writing to the sheet
  return transactions.map(transaction => {
    // Determine if this is an outflow or inflow
    const amount = transaction.amount;
    let outflow = "";
    let inflow = "";
    
    if (amount > 0) {
      outflow = amount; // Positive amount in Plaid means money leaving the account (expense)
    } else {
      inflow = Math.abs(amount); // Negative amount means money coming into the account (income)
    }
    
    // Format the account name to include institution name if available
    let accountName = accountMap[transaction.account_id] || "Unknown Account";
    if (transaction.institution_name) {
      accountName = `${transaction.institution_name} - ${accountName}`;
    }
    
    // Determine status based on pending flag
    const status = transaction.pending ? PENDING_STATUS : CLEARED_STATUS;
    
    return [
      transaction.date, // Date
      outflow, // Outflow
      inflow, // Inflow
      transaction.personal_finance_category ? JSON.stringify(transaction.personal_finance_category) : "", // Category
      accountName, // Account name with institution
      transaction.name, // Memo (transaction description)
      status, // Status based on pending flag
      "", // Empty column for formatting
      transaction.transaction_id, // Transaction ID generated by plaid
      "" // Placeholder value for categorization source
    ];
  });
}

/**
 * Reads existing transactions from the sheet.
 * @return {Array} Array of transaction rows from the sheet.
 */
function readExistingTransactions() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetById(TRANSACTIONS_SHEET_ID);
    
    // Get the transaction range
    const transactionRange = sheet.getRange(TRANSACTIONS_RANGE);
    const startRow = transactionRange.getRow();
    const startCol = transactionRange.getColumn();
    const numCols = transactionRange.getNumColumns();
    
    // Get the last row with content in the sheet
    var lastRow = sheet.getLastRow();
    // Shrink the range even further if needed (to account for the hidden configuration rows)
    const dates = sheet.getRange(`B1:B${lastRow}`).getValues(); // Check column B for date
    lastRow -= dates.reverse().findIndex(c=>c[0]!='');
    const numRows = lastRow - startRow + 1;
    
    // Get only the rows that have data
    return sheet.getRange(startRow, startCol, numRows, numCols).getValues();
  } catch (error) {
    Logger.log(`Error reading existing transactions: ${error}`);
    return [];
  }
}

/**
 * Gets the plaid transaction ID from a transaction row.
 * @param {Array} row - A transaction row from the sheet.
 * @return {string} A unique identifier string.
 */
function getTransactionId(row) {
  // If the Plaid transaction ID is available, use it
  if (row[8]) {
    return row[8];
  }
}

/**
 * Gets the list of valid category options from the sheet.
 * @return {Array} Array of valid category options.
 */
function getCategoryOptions() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetById(TRANSACTIONS_SHEET_ID);
    const categoryRange = sheet.getRange(CATEGORY_OPTIONS_RANGE);
    const categoryValues = categoryRange.getValues();
    
    // Flatten the 2D array and filter out empty values
    return categoryValues.flat().filter(category => category && category.trim() !== "");
  } catch (error) {
    Logger.log(`Error getting category options: ${error}`);
    return [];
  }
}

/**
 * Categorizes transactions based on existing patterns or using Gemini AI.
 * Only processes rows without a valid category.
 * @param {Array} rowsToProcess - Transaction rows to categorize.
 * @param {Array} referenceRows - Reference transaction rows for pattern matching.
 * @param {Array} validCategories - List of valid categories from the sheet.
 * @return {number} Number of transactions categorized.
 */
function categorizeTransactions(rowsToProcess, referenceRows, validCategories) {
  // Create a map of transaction descriptions to categories from reference transactions. Also include the source of the categorization
  const memoToCategoryMap = {};
  let categorizedCount = 0;
  
  // Process in reverse order (from oldest to newest) so newer categories overwrite older ones
  for (let i = referenceRows.length - 1; i >= 0; i--) {
    const row = referenceRows[i];
    if (row[3] && row[5] && validCategories.includes(row[3])) { // If it has a valid category and memo
      const memo = row[5].toLowerCase().trim();
      
      // Only override existing categorization pattern if the transaction has been manually categorized by user
      if (!memoToCategoryMap[memo] || row[9] === "USER") {
        memoToCategoryMap[memo] = {category: row[3], source: row[9]};
      }
    }
  }
  
  // Process each transaction
  for (let i = 0; i < rowsToProcess.length; i++) {
    const row = rowsToProcess[i];
    
    /*
    // Skip if already has a valid category
    if (row[3] && validCategories.includes(row[3])) {
      continue;
    }
    */
    // Skip if already categorized by user
    if (row[9] && row[9] === "USER") {
      continue;
    }
    
    const memo = row[5] ? row[5].toLowerCase().trim() : "";
    
    // First try to match with existing transactions or previously categorized transactions
    if (memo && memoToCategoryMap[memo]) {
      row[3] = memoToCategoryMap[memo].category;
      row[9] = memoToCategoryMap[memo].source;
      Logger.log(`Categorized "${row[5]}" as "${row[3]}" based on existing transactions (source: ${row[9]})`);
      categorizedCount++;
      continue; // Move to next transaction
    }
    
    // If no match found and Gemini API key is provided, use Gemini
    if (GEMINI_API_KEY) {
      try {
        const category = categorizeWithGemini(row, validCategories);
        if (category && validCategories.includes(category)) {
          row[3] = category;
          row[9] = "AI";
          
          // Add this newly categorized transaction to the map for future use
          memoToCategoryMap[memo] = {category: row[3], source: row[9]};
          
          Logger.log(`Categorized "${row[5]}" as "${row[3]}" using Gemini AI`);
          categorizedCount++;
        } else {
          Logger.log(`Gemini returned an invalid category: ${category}`);
          row[9] = "ERROR";
        }
      } catch (error) {
        Logger.log(`Error categorizing with Gemini: ${error}`);
        row[9] = "ERROR";
      }
    }
  }
  
  Logger.log(`Categorized ${categorizedCount} transactions in total`);
  return categorizedCount;
}

/**
 * Recategorizes all existing transactions that don't have a valid category or are categorized by AI.
 * Reads all transactions from the sheet, identifies those without valid categories,
 * and attempts to categorize them using existing patterns or Gemini AI.
 */
function recategorizeAllTransactions() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // Read existing transactions
    const existingRows = readExistingTransactions();
    Logger.log(`Read ${existingRows.length} existing transactions from the sheet`);
    
    // Get valid categories
    const validCategories = getCategoryOptions();
    Logger.log(`Found ${validCategories.length} valid categories`);
    
    // Categorize transactions that don't have a valid category
    const categorizedCount = categorizeTransactions(existingRows, existingRows, validCategories);
    
    if (categorizedCount > 0) {
      // Write the updated transactions back to the sheet
      const result = writeTransactionsToSheet(existingRows);
      
      if (result.success) {
        ui.alert('Success', `Successfully categorized ${categorizedCount} transactions.`, ui.ButtonSet.OK);
      } else {
        ui.alert('Error', `Categorized ${categorizedCount} transactions but failed to write to sheet: ${result.message}`, ui.ButtonSet.OK);
      }
    } else {
      ui.alert('Information', 'No transactions needed categorization.', ui.ButtonSet.OK);
    }
  } catch (error) {
    Logger.log(`Error recategorizing transactions: ${error}`);
    ui.alert('Error', `Failed to recategorize transactions: ${error.toString()}`, ui.ButtonSet.OK);
  }
}

/**
 * Uses Gemini AI to categorize a transaction.
 * @param {Array} row - Transaction row to categorize.
 * @param {Array} validCategories - List of valid categories.
 * @return {string} The suggested category or empty string if unable to categorize.
 */
function categorizeWithGemini(row, validCategories) {
  if (!GEMINI_API_KEY) {
    return "";
  }

  // Add a short delay to avoid hitting API rate limits
  Utilities.sleep(2000);
  
  const memo = row[5] || "";
  const amount = row[1] ? row[1] : (row[2] ? -row[2] : 0);
  const account = row[4] || "";
  const plaid_personal_finance_category = row[3];

  // Prepare the prompt for Gemini
  const prompt = `Categorize this financial transaction into exactly one of the following categories:\n${validCategories.join("\n")}\n\nTransaction details:\n${INCLUDE_MEMO_IN_PROMPT ? `Memo: ${memo}\n` : ""}Personal Finance Category according to Plaid: ${plaid_personal_finance_category}\n\nCategory:`;
  
  // Log the transaction and prompt being sent to Gemini
  Logger.log(`Sending to Gemini - Transaction: "${memo}" Plaid personal finance category: ${plaid_personal_finance_category}`);
  
  // Call Gemini API
  const url = `${GEMINI_API_URL}?key=${GEMINI_API_KEY}`;
  const payload = {
    contents: [{
      role: "user",
      parts: [{
        text: prompt
      }]
    }],
    systemInstruction: {
      parts: [{
        text: "You are a financial transaction categorization tool. Only respond with a single category name that exactly matches one from the provided list, no additional text."
      }]
    },
    generationConfig: {
      temperature: 0.1,
      maxOutputTokens: 10,
    }
  };
  
  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseData = JSON.parse(response.getContentText());
    
    if (responseData.candidates && responseData.candidates.length > 0) {
      const generatedText = responseData.candidates[0].content.parts[0].text.trim();

      // Find the matching category (relevant in case category has emojis)
      const matchingCategory = validCategories.find(category => 
        category.includes(generatedText)
      );
      return matchingCategory || generatedText;
    } else {
      Logger.log(`Gemini returned no candidates for transaction. Gemini response: ${response.getContentText()}`);
      return "";
    }
  } catch (error) {
    Logger.log(`Error calling Gemini API: ${error}`);
    return "";
  }
}

/**
 * Updates the existing transactions in the sheet with the added/modified/deleted transactions from the Plaid endpoint.
 * @param {Array} addedTransactions - Array of added transaction objects from Plaid.
 * @param {Array} modifiedTransactions - Array of modified transaction objects from Plaid.
 * @param {Array} removedTransactionIds - Array of removed transaction IDs from Plaid.
 * @param {Array} accounts - Array of account objects from Plaid.
 * @return {Array} Array of the combined/updated transaction rows.
 */
function getUpdatedTransactions(addedTransactions, modifiedTransactions, removedTransactionIds, accounts) {
  // Read existing transactions from the sheet
  let existingRows = readExistingTransactions();
  // Filter out empty rows
  existingRows = existingRows.filter(row => row[1] !== '' || row[2] !== ''); // Remove rows that have no inflow or outflow
  Logger.log(`Read ${existingRows.length} existing transactions from the sheet`);

  // Format added and modified transactions for the sheet
  const addedRows = formatTransactionsForSheet(addedTransactions, accounts);
  const modifiedRows = formatTransactionsForSheet(modifiedTransactions, accounts);
  
  // Create a map of transaction IDs to row indices for existing transactions
  const existingTransactionMap = new Map();
  
  // Process existing rows to build the map
  for (let i = 0; i < existingRows.length; i++) {
    const row = existingRows[i];
    const transactionId = getTransactionId(row);
    if (transactionId) {
      existingTransactionMap.set(transactionId, i);
    }
  }
  
  // Process removed transactions - remove them from the result rows
  if (removedTransactionIds.length > 0) {
    Logger.log(`Processing ${removedTransactionIds.length} removed transactions`);
    
    // Create a set of removed transaction IDs for faster lookup
    const removedIdsSet = new Set(removedTransactionIds);
    
    // Filter out removed transactions
    const filteredRows = existingRows.filter(row => {
      const transactionId = getTransactionId(row);
      const shouldRemove = transactionId && removedIdsSet.has(transactionId);
      
      if (shouldRemove) {
        Logger.log(`Removing transaction with ID: ${transactionId}`);
      }
      
      return !shouldRemove; // Keep if not in the removed set
    });
    
    // Update existingRows with the filtered rows
    existingRows.length = 0;
    existingRows.push(...filteredRows);
  }
  
  // Process modified transactions - update existing transactions
  if (modifiedRows.length > 0) {
    Logger.log(`Processing ${modifiedRows.length} modified transactions`);
    
    for (const modifiedRow of modifiedRows) {
      const transactionId = getTransactionId(modifiedRow);
      
      if (transactionId && existingTransactionMap.has(transactionId)) {
        // Update the existing transaction
        const index = existingTransactionMap.get(transactionId);
        existingRows[index] = modifiedRow;
        Logger.log(`Updated transaction with ID: ${transactionId}`);
      } else {
        // If not found, treat it as a new transaction
        Logger.log(`Modified transaction not found in existing data, adding as new: ${transactionId}`);
        addedRows.push(modifiedRow);
      }
    }
  }
  
  // Process added transactions
  if (addedRows.length > 0) {
    Logger.log(`Processing ${addedRows.length} added transactions`);
    
    // Categorize the new transactions using existing transactions and Gemini
    const categoryOptions = getCategoryOptions();
    categorizeTransactions(addedRows, existingRows, categoryOptions);
  }
  
  // Combine existing and new transactions
  const allRows = [...existingRows, ...addedRows];
  
  // Sort all transactions by date (newest first)
  allRows.sort((a, b) => {
    if (!a[0]) return -1; // Empty rows go to the top
    if (!b[0]) return 1;
    const dateA = new Date(a[0]);
    const dateB = new Date(b[0]);
    return dateB - dateA; // Descending order (newest first)
  });
  
  return allRows;
}

/**
 * Writes transaction data to the specified sheet.
 * @param {Array} transactionRows - Array of formatted transaction rows ready for the sheet.
 * @return {Object} Object containing success status and message.
 */
function writeTransactionsToSheet(transactionRows) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetById(TRANSACTIONS_SHEET_ID);
    
    // Get the transaction range
    const transactionRange = sheet.getRange(TRANSACTIONS_RANGE);
    const startRow = transactionRange.getRow();
    const startCol = transactionRange.getColumn();
    
    // Check if transaction data is available
    if (transactionRows.length > 0) {
      const rangeWidth = transactionRange.getNumColumns();
      const dataWidth = transactionRows[0].length;
      
      // Check if transaction data matches the expected width
      if (dataWidth !== rangeWidth) {
        return { 
          success: false, 
          message: `Error: Transaction data width (${dataWidth}) doesn't match the expected range width (${rangeWidth}).` 
        };
      }
      
      // Check if transaction data exceeds the available rows
      const rangeHeight = transactionRange.getNumRows();
      if (transactionRows.length > rangeHeight) {
        return { 
          success: false, 
          message: `Error: Too many transactions (${transactionRows.length}) for the available range (${rangeHeight} rows).` 
        };
      }
      
      // Clear the existing data in the range
      const clearRange = sheet.getRange(startRow, startCol, rangeHeight, rangeWidth);
      clearRange.clearContent();
      
      // Write the new data to the sheet
      const range = sheet.getRange(startRow, startCol, transactionRows.length, dataWidth);
      range.setValues(transactionRows);
    }
    
    Logger.log(`Successfully wrote ${transactionRows.length} transactions to the sheet.`);
    return { success: true };
  } catch (error) {
    Logger.log(`Error writing transactions to sheet: ${error}`);
    return { success: false, message: `Error: ${error.toString()}` };
  }
}

/**
 * Adds a new empty transaction row at the top of the transactions list.
 * @return {boolean} True if successful, false otherwise.
 */
function addNewTransactionRow() {
  try {
    // Check if a sync task is already running
    const scriptProperties = PropertiesService.getScriptProperties();
    const syncRunning = scriptProperties.getProperty('SYNC_TASK_RUNNING');
    
    if (syncRunning === 'true') {
      throw new Error('Sync task is already running. Please wait for it to finish before adding a new transaction row.');
    }

    // Get existing transactions
    const existingTransactions = readExistingTransactions();
    
    // Get transaction range to determine number of columns
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetById(TRANSACTIONS_SHEET_ID);
    const transactionRange = sheet.getRange(TRANSACTIONS_RANGE);
    const numCols = transactionRange.getNumColumns();
    
    // Create an empty row
    const emptyRow = Array(numCols).fill('');
    
    // Add the empty row at the beginning of the existing transactions array
    existingTransactions.unshift(emptyRow);
    
    // Write the modified data back to the sheet
    const result = writeTransactionsToSheet(existingTransactions);
    
    // Show error alert if there was a problem
    if (!result.success) {
      SpreadsheetApp.getUi().alert('Error', result.message, SpreadsheetApp.getUi().ButtonSet.OK);
    }
    
    return result.success;
  } catch (error) {
    const errorMessage = `Error adding new transaction row: ${error.message}`;
    Logger.log(errorMessage);
    SpreadsheetApp.getUi().alert('Error', errorMessage, SpreadsheetApp.getUi().ButtonSet.OK);
    return false;
  }
}