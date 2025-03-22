/**
 * NBU Exchange Rate Fetcher for Google Sheets
 * 
 * This script fetches exchange rates from the National Bank of Ukraine (NBU)
 * and updates your Google Sheet with USD and EUR rates.
 */

/**
 * Creates a custom menu in Google Sheets for easy access to functions.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('NBU Exchange Rates')
    .addItem('Fetch Today\'s Rates', 'fetchCurrentRates')
    .addItem('Fetch Rates by Date', 'showDatePrompt')
    .addSeparator()
    .addItem('Add Scheduled Updates', 'createScheduledTrigger')
    .addItem('Remove Scheduled Updates', 'removeScheduledTrigger')
    .addToUi();
}

/**
 * Shows a date picker dialog to fetch rates for a specific date.
 */
function showDatePrompt() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt(
    'Fetch Exchange Rates by Date',
    'Enter date (YYYY-MM-DD):',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() == ui.Button.OK) {
    var dateStr = response.getResponseText();
    if (isValidDateFormat(dateStr)) {
      fetchRatesForDate(dateStr);
    } else {
      ui.alert('Invalid date format. Please use YYYY-MM-DD format.');
    }
  }
}

/**
 * Validates date format (YYYY-MM-DD).
 */
function isValidDateFormat(dateString) {
  var regex = /^\d{4}-\d{2}-\d{2}$/;
  if (!regex.test(dateString)) return false;
  
  var parts = dateString.split('-');
  var year = parseInt(parts[0], 10);
  var month = parseInt(parts[1], 10) - 1; // JS months are 0-based
  var day = parseInt(parts[2], 10);
  
  var date = new Date(year, month, day);
  return date.getFullYear() === year && 
         date.getMonth() === month && 
         date.getDate() === day;
}

/**
 * Fetches current exchange rates from NBU.
 */
function fetchCurrentRates() {
  var today = new Date();
  var dateStr = Utilities.formatDate(today, 'GMT+2', 'yyyy-MM-dd');
  fetchRatesForDate(dateStr);
}

/**
 * Fetches exchange rates for a specific date from NBU API.
 * 
 * @param {string} dateStr - Date in YYYY-MM-DD format
 */
function fetchRatesForDate(dateStr) {
  // Remove hyphens from date for API
  var apiDateStr = dateStr.replace(/-/g, '');
  
  try {
    // Fetch exchange rates from NBU
    var apiUrl = "https://bank.gov.ua/NBUStatService/v1/statdirectory/exchange?date=" + apiDateStr + "&json";
    var response = UrlFetchApp.fetch(apiUrl);
    var data = JSON.parse(response.getContentText());
    
    // Filter for USD and EUR
    var usdData = data.find(item => item.cc === 'USD');
    var eurData = data.find(item => item.cc === 'EUR');
    
    if (!usdData || !eurData) {
      SpreadsheetApp.getUi().alert("Could not find USD or EUR exchange rates in the data.");
      return;
    }
    
    // Update the sheet with the exchange rates
    updateSheet(dateStr, usdData.rate, eurData.rate);
    
    // Log success
    Logger.log("Successfully fetched rates for " + dateStr);
    SpreadsheetApp.getUi().alert("Exchange rates updated for " + dateStr);
    
  } catch (error) {
    Logger.log("Error fetching exchange rates: " + error.toString());
    SpreadsheetApp.getUi().alert("Error fetching exchange rates: " + error.toString());
  }
}

/**
 * Updates the active sheet with the exchange rate data.
 * 
 * @param {string} dateStr - Date in YYYY-MM-DD format
 * @param {number} usdRate - USD to UAH exchange rate
 * @param {number} eurRate - EUR to UAH exchange rate
 */
function updateSheet(dateStr, usdRate, eurRate) {
  var sheet = SpreadsheetApp.getActiveSheet();
  
  // Clear the relevant range where we'll write the data
  sheet.getRange("A1:B4").clearContent();
  
  // Set the values
  sheet.getRange("A1").setValue("Date");
  sheet.getRange("B1").setValue(dateStr);
  sheet.getRange("A2").setValue("Currency");
  sheet.getRange("B2").setValue("Rate (UAH)");
  sheet.getRange("A3").setValue("USD");
  sheet.getRange("B3").setValue(usdRate);
  sheet.getRange("A4").setValue("EUR");
  sheet.getRange("B4").setValue(eurRate);
  
  // Optional: Format the cells
  sheet.getRange("B3:B4").setNumberFormat("0.0000");
}

/**
 * Creates a daily trigger to automatically update exchange rates.
 */
function createScheduledTrigger() {
  // Check if a trigger already exists
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'fetchCurrentRates') {
      SpreadsheetApp.getUi().alert("A scheduled update already exists.");
      return;
    }
  }
  
  // Create a trigger to run at 10:00 AM Kyiv time (GMT+2/3) every weekday
  ScriptApp.newTrigger('fetchCurrentRates')
    .timeBased()
    .atHour(10)
    .everyDays(1)
    .create();
  
  SpreadsheetApp.getUi().alert("Daily update scheduled for 10:00 AM.");
}

/**
 * Removes all scheduled triggers for exchange rate updates.
 */
function removeScheduledTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  var found = false;
  
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'fetchCurrentRates') {
      ScriptApp.deleteTrigger(triggers[i]);
      found = true;
    }
  }
  
  if (found) {
    SpreadsheetApp.getUi().alert("Scheduled updates have been removed.");
  } else {
    SpreadsheetApp.getUi().alert("No scheduled updates were found.");
  }
}

/**
 * Fetches exchange rates for multiple dates and creates a historical record.
 * 
 * @param {string} startDateStr - Start date in YYYY-MM-DD format
 * @param {string} endDateStr - End date in YYYY-MM-DD format
 */
function fetchHistoricalRates(startDateStr, endDateStr) {
  if (!isValidDateFormat(startDateStr) || !isValidDateFormat(endDateStr)) {
    SpreadsheetApp.getUi().alert('Invalid date format. Please use YYYY-MM-DD format.');
    return;
  }
  
  var startDate = new Date(startDateStr);
  var endDate = new Date(endDateStr);
  
  if (startDate > endDate) {
    SpreadsheetApp.getUi().alert('Start date must be before end date.');
    return;
  }
  
  // Create a new sheet for historical data
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var historySheet = ss.getSheetByName('Exchange Rate History');
  
  if (!historySheet) {
    historySheet = ss.insertSheet('Exchange Rate History');
    // Add headers
    historySheet.getRange("A1:C1").setValues([["Date", "USD", "EUR"]]);
    historySheet.getRange("A1:C1").setFontWeight("bold");
  }
  
  // Get existing dates in the history sheet to avoid duplicates
  var existingDates = {};
  var dataRange = historySheet.getDataRange();
  if (dataRange.getNumRows() > 1) {
    var values = dataRange.getValues();
    for (var i = 1; i < values.length; i++) {
      if (values[i][0]) {
        var dateKey = Utilities.formatDate(new Date(values[i][0]), 'GMT', 'yyyy-MM-dd');
        existingDates[dateKey] = true;
      }
    }
  }
  
  // Process each date in the range
  var currentDate = new Date(startDate);
  var rowsToAdd = [];
  
  while (currentDate <= endDate) {
    var dateKey = Utilities.formatDate(currentDate, 'GMT', 'yyyy-MM-dd');
    
    // Skip if this date is already in the history
    if (!existingDates[dateKey]) {
      var apiDateStr = dateKey.replace(/-/g, '');
      
      try {
        // Fetch exchange rates from NBU
        var apiUrl = "https://bank.gov.ua/NBUStatService/v1/statdirectory/exchange?date=" + apiDateStr + "&json";
        var response = UrlFetchApp.fetch(apiUrl);
        var data = JSON.parse(response.getContentText());
        
        // Filter for USD and EUR
        var usdData = data.find(item => item.cc === 'USD');
        var eurData = data.find(item => item.cc === 'EUR');
        
        if (usdData && eurData) {
          rowsToAdd.push([new Date(dateKey), usdData.rate, eurData.rate]);
        }
        
        // Sleep to avoid hitting API limits
        Utilities.sleep(500);
        
      } catch (error) {
        Logger.log("Error fetching data for " + dateKey + ": " + error.toString());
      }
    }
    
    // Move to next day
    currentDate.setDate(currentDate.getDate() + 1);
  }
  
  // Add the rows if we have any data
  if (rowsToAdd.length > 0) {
    // Find the last row
    var lastRow = historySheet.getLastRow();
    historySheet.getRange(lastRow + 1, 1, rowsToAdd.length, 3).setValues(rowsToAdd);
    
    // Format the date columns
    historySheet.getRange(lastRow + 1, 1, rowsToAdd.length, 1).setNumberFormat("yyyy-MM-dd");
    historySheet.getRange(lastRow + 1, 2, rowsToAdd.length, 2).setNumberFormat("0.0000");
    
    SpreadsheetApp.getUi().alert("Added exchange rates for " + rowsToAdd.length + " days.");
  } else {
    SpreadsheetApp.getUi().alert("No new data added. Dates may already exist in the history sheet.");
  }
} 
