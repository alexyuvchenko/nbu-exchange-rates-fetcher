/**
 * NBU Exchange Rate Fetcher for Google Sheets
 *
 * This script fetches exchange rates from the National Bank of Ukraine (NBU)
 * and updates your Google Sheet with USD and EUR rates.
 */

// Configuration constants
const CONFIG = {
  API: {
    BASE_URL: 'https://bank.gov.ua/NBUStatService/v1/statdirectory/exchange',
    CURRENCIES: {
      USD: { code: 'USD', name: 'US Dollar' },
      EUR: { code: 'EUR', name: 'Euro' },
    },
  },
  DATE_FORMAT: {
    API: 'yyyyMMdd',
    DISPLAY: 'yyyy-MM-dd',
    INPUT: 'yyyy-MM-dd',
    INPUT_SLASH: 'dd/MM/yyyy',
    INPUT_DOT: 'dd.MM.yyyy',
  },
  CACHE: {
    KEY_PREFIX: 'NBU_RATES_',
    EXPIRY_DAYS: 90, // Keep rates in cache for 90 days
  },
  SHEET: {
    COLUMN_NAMES: {
      DATE: 'Date',
      USD_NBU: 'USD NBU',
      EUR_NBU: 'EURO NBU',
    },
  },
};

/**
 * Class for date handling and formatting
 */
class DateHandler {
  /**
   * Create a UTC Date object from components
   */
  static createUTCDate(year, month, day) {
    return new Date(Date.UTC(year, month, day));
  }

  /**
   * Format date as YYYY-MM-DD
   */
  static formatYYYYMMDD(date) {
    const year = date.getFullYear();
    const month = date.getMonth() + 1; // JavaScript months are 0-indexed
    const day = date.getDate();

    return year + '-' + (month < 10 ? '0' + month : month) + '-' + (day < 10 ? '0' + day : day);
  }

  /**
   * Format date for the NBU API
   */
  static formatForAPI(date) {
    return Utilities.formatDate(date, 'UTC', CONFIG.DATE_FORMAT.API);
  }

  /**
   * Validate date format (YYYY-MM-DD)
   */
  static isValidDateFormat(dateString) {
    const regex = /^\d{4}-\d{2}-\d{2}$/;
    if (!regex.test(dateString)) return false;

    const parts = dateString.split('-');
    const year = parseInt(parts[0], 10);
    const month = parseInt(parts[1], 10) - 1; // JS months are 0-based
    const day = parseInt(parts[2], 10);

    const date = new Date(year, month, day);
    return date.getFullYear() === year && date.getMonth() === month && date.getDate() === day;
  }

  /**
   * Parse a date string or object into a standardized format
   */
  static parse(dateValue) {
    let dateObj;

    // Handle different possible date formats
    if (typeof dateValue === 'string') {
      // Try to parse string format DD/MM/YYYY
      if (dateValue.match(/^\d{1,2}\/\d{1,2}\/\d{4}$/)) {
        const parts = dateValue.split('/');
        const day = parseInt(parts[0], 10);
        const month = parseInt(parts[1], 10);
        const year = parseInt(parts[2], 10);

        dateObj = this.createUTCDate(year, month - 1, day);
        if (isNaN(dateObj.getTime())) return null;

        // Try to parse string format DD.MM.YYYY
      } else if (dateValue.match(/^\d{1,2}\.\d{1,2}\.\d{4}$/)) {
        const parts = dateValue.split('.');
        const day = parseInt(parts[0], 10);
        const month = parseInt(parts[1], 10);
        const year = parseInt(parts[2], 10);

        dateObj = this.createUTCDate(year, month - 1, day);
        if (isNaN(dateObj.getTime())) return null;
      } else if (dateValue.match(/^\d{4}-\d{2}-\d{2}$/)) {
        // YYYY-MM-DD format
        const parts = dateValue.split('-');
        const year = parseInt(parts[0], 10);
        const month = parseInt(parts[1], 10);
        const day = parseInt(parts[2], 10);

        dateObj = this.createUTCDate(year, month - 1, day);
        if (isNaN(dateObj.getTime())) return null;
      } else {
        // Try to parse other date formats
        dateObj = new Date(dateValue);
        if (isNaN(dateObj.getTime())) return null;

        // Extract year, month, day and recreate using UTC
        const year = dateObj.getFullYear();
        const month = dateObj.getMonth();
        const day = dateObj.getDate();
        dateObj = this.createUTCDate(year, month, day);
      }
    } else if (dateValue instanceof Date) {
      // It's already a date object, but we still need to handle timezone
      const year = dateValue.getFullYear();
      const month = dateValue.getMonth();
      const day = dateValue.getDate();
      dateObj = this.createUTCDate(year, month, day);
    } else {
      return null;
    }

    const apiDateStr = this.formatForAPI(dateObj);
    return { date: dateObj, apiDateStr: apiDateStr };
  }
}

/**
 * Class for cache operations
 */
class RateCache {
  /**
   * Save exchange rates to cache
   */
  static save(dateKey, rates) {
    if (!rates) return;

    try {
      const scriptProps = PropertiesService.getScriptProperties();
      const cacheKey = CONFIG.CACHE.KEY_PREFIX + dateKey;
      const cacheData = JSON.stringify({
        rates: rates,
        timestamp: new Date().getTime(),
      });

      scriptProps.setProperty(cacheKey, cacheData);
      Logger.log('Saved rates to cache for date: ' + dateKey);
    } catch (error) {
      Logger.log('Error saving to cache: ' + error.toString());
    }
  }

  /**
   * Get exchange rates from cache
   */
  static get(dateKey) {
    try {
      const scriptProps = PropertiesService.getScriptProperties();
      const cacheKey = CONFIG.CACHE.KEY_PREFIX + dateKey;
      const cacheData = scriptProps.getProperty(cacheKey);

      if (!cacheData) return null;

      const data = JSON.parse(cacheData);
      const now = new Date().getTime();
      const cacheTime = data.timestamp;
      const expiryTime = CONFIG.CACHE.EXPIRY_DAYS * 24 * 60 * 60 * 1000;

      // Check if cache has expired
      if (now - cacheTime > expiryTime) {
        // Cache expired, clean it up
        scriptProps.deleteProperty(cacheKey);
        return null;
      }

      Logger.log('Retrieved rates from cache for date: ' + dateKey);
      return data.rates;
    } catch (error) {
      Logger.log('Error retrieving from cache: ' + error.toString());
      return null;
    }
  }

  /**
   * Clean up expired cache entries
   */
  static cleanup() {
    try {
      const scriptProps = PropertiesService.getScriptProperties();
      const allProps = scriptProps.getProperties();
      const now = new Date().getTime();
      const expiryTime = CONFIG.CACHE.EXPIRY_DAYS * 24 * 60 * 60 * 1000;
      let cleanedCount = 0;

      // Find and delete expired cache entries
      for (const key in allProps) {
        if (key.startsWith(CONFIG.CACHE.KEY_PREFIX)) {
          try {
            const data = JSON.parse(allProps[key]);
            if (now - data.timestamp > expiryTime) {
              scriptProps.deleteProperty(key);
              cleanedCount++;
            }
          } catch (e) {
            // Invalid JSON, delete this entry
            scriptProps.deleteProperty(key);
            cleanedCount++;
          }
        }
      }

      if (cleanedCount > 0) {
        Logger.log('Cleaned up ' + cleanedCount + ' expired cache entries');
      }
    } catch (error) {
      Logger.log('Error cleaning cache: ' + error.toString());
    }
  }

  /**
   * Clear all cache entries
   */
  static clearAll() {
    try {
      const scriptProps = PropertiesService.getScriptProperties();
      const allProps = scriptProps.getProperties();
      let clearedCount = 0;

      for (const key in allProps) {
        if (key.startsWith(CONFIG.CACHE.KEY_PREFIX)) {
          scriptProps.deleteProperty(key);
          clearedCount++;
        }
      }

      return clearedCount;
    } catch (error) {
      Logger.log('Error clearing cache: ' + error.toString());
      throw error;
    }
  }
}

/**
 * Class for NBU API operations
 */
class NbuApi {
  /**
   * Fetch exchange rates from the NBU API
   */
  static fetchRates(apiDateStr) {
    // First check if we have cached data
    const cachedRates = RateCache.get(apiDateStr);
    if (cachedRates) {
      return cachedRates;
    }

    const apiUrl = `${CONFIG.API.BASE_URL}?date=${apiDateStr}&json`;

    try {
      const response = UrlFetchApp.fetch(apiUrl);
      const data = JSON.parse(response.getContentText());

      // Filter for USD and EUR
      const usdData = data.find(item => item.cc === CONFIG.API.CURRENCIES.USD.code);
      const eurData = data.find(item => item.cc === CONFIG.API.CURRENCIES.EUR.code);

      if (!usdData || !eurData) {
        return null;
      }

      const rates = {
        usd: usdData.rate,
        eur: eurData.rate,
      };

      // Save to cache for future use
      RateCache.save(apiDateStr, rates);

      return rates;
    } catch (error) {
      Logger.log('Error fetching exchange rates: ' + error.toString());
      return null;
    }
  }
}

/**
 * Class for sheet operations
 */
class SheetHandler {
  /**
   * Update the active sheet with exchange rate data
   */
  static updateActiveSheet(dateStr, usdRate, eurRate) {
    const sheet = SpreadsheetApp.getActiveSheet();

    // Clear the relevant range where we'll write the data
    sheet.getRange('A1:B4').clearContent();

    // Set the values
    sheet.getRange('A1').setValue('Date');
    sheet.getRange('B1').setValue(dateStr);
    sheet.getRange('A2').setValue('Currency');
    sheet.getRange('B2').setValue('Rate (UAH)');
    sheet.getRange('A3').setValue('USD');
    sheet.getRange('B3').setValue(usdRate);
    sheet.getRange('A4').setValue('EUR');
    sheet.getRange('B4').setValue(eurRate);

    // Format the cells
    sheet.getRange('B3:B4').setNumberFormat('0.0000');
  }

  /**
   * Update rates based on date column
   */
  static updateRatesByDateColumn() {
    const sheet = SpreadsheetApp.getActiveSheet();
    const data = sheet.getDataRange().getValues();

    // Find the column indices
    const headers = data[0];
    const usdNbuColIndex = headers.indexOf(CONFIG.SHEET.COLUMN_NAMES.USD_NBU);
    const eurNbuColIndex = headers.indexOf(CONFIG.SHEET.COLUMN_NAMES.EUR_NBU);
    const dateColIndex = headers.indexOf(CONFIG.SHEET.COLUMN_NAMES.DATE);

    // Check if all required columns exist
    if (usdNbuColIndex === -1 || eurNbuColIndex === -1 || dateColIndex === -1) {
      SpreadsheetApp.getUi().alert(
        `One or more required columns not found. Please ensure your sheet has '${CONFIG.SHEET.COLUMN_NAMES.USD_NBU}', '${CONFIG.SHEET.COLUMN_NAMES.EUR_NBU}', and '${CONFIG.SHEET.COLUMN_NAMES.DATE}' columns.`
      );
      return;
    }

    const results = this.processDateColumnRows(
      sheet,
      data,
      dateColIndex,
      usdNbuColIndex,
      eurNbuColIndex
    );

    // Format the rate columns
    if (results.updatedRows > 0) {
      sheet.getRange(2, usdNbuColIndex + 1, data.length - 1, 1).setNumberFormat('0.0000');
      sheet.getRange(2, eurNbuColIndex + 1, data.length - 1, 1).setNumberFormat('0.0000');
      SpreadsheetApp.getUi().alert('Updated exchange rates for ' + results.updatedRows + ' rows.');
    } else {
      SpreadsheetApp.getUi().alert('No exchange rates were updated. Check your date format.');
    }
  }

  /**
   * Process rows for date-based rate updates
   */
  static processDateColumnRows(sheet, data, dateColIndex, usdColIndex, eurColIndex) {
    // Keep track of dates we've already looked up to avoid redundant API calls in this session
    const sessionCache = {};
    let updatedRows = 0;

    // Start from row 1 (skipping header row)
    for (let i = 1; i < data.length; i++) {
      const dateValue = data[i][dateColIndex];

      // Skip if no date value
      if (!dateValue) continue;

      const parsedDate = DateHandler.parse(dateValue);
      if (!parsedDate) continue;

      const apiDateStr = parsedDate.apiDateStr;

      // First check our session cache to minimize even PropertiesService calls
      if (apiDateStr in sessionCache) {
        // Update this row with cached values
        sheet.getRange(i + 1, usdColIndex + 1).setValue(sessionCache[apiDateStr].usd);
        sheet.getRange(i + 1, eurColIndex + 1).setValue(sessionCache[apiDateStr].eur);
        updatedRows++;
        continue;
      }

      const rates = NbuApi.fetchRates(apiDateStr);

      if (rates) {
        // Update the row with the exchange rates
        sheet.getRange(i + 1, usdColIndex + 1).setValue(rates.usd);
        sheet.getRange(i + 1, eurColIndex + 1).setValue(rates.eur);

        // Cache the rates for this session
        sessionCache[apiDateStr] = rates;

        updatedRows++;

        // Sleep to avoid hitting API limits (only needed for actual API calls)
        if (!RateCache.get(apiDateStr)) {
          Utilities.sleep(500);
        }
      }
    }

    return { updatedRows: updatedRows };
  }
}

/**
 * Class for trigger management
 */
class TriggerManager {
  /**
   * Create a scheduled trigger
   */
  static createScheduled() {
    // Check if a trigger already exists
    const triggers = ScriptApp.getProjectTriggers();
    for (let i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === 'fetchCurrentRates') {
        SpreadsheetApp.getUi().alert('A scheduled update already exists.');
        return;
      }
    }

    // Create a trigger to run at 10:00 AM every day
    ScriptApp.newTrigger('fetchCurrentRates').timeBased().atHour(10).everyDays(1).create();

    SpreadsheetApp.getUi().alert('Daily update scheduled for 10:00 AM.');
  }

  /**
   * Remove scheduled triggers
   */
  static removeScheduled() {
    const triggers = ScriptApp.getProjectTriggers();
    let found = false;

    for (let i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === 'fetchCurrentRates') {
        ScriptApp.deleteTrigger(triggers[i]);
        found = true;
      }
    }

    if (found) {
      SpreadsheetApp.getUi().alert('Scheduled updates have been removed.');
    } else {
      SpreadsheetApp.getUi().alert('No scheduled updates were found.');
    }
  }
}

// ======================================
// Public Functions (Menu Entry Points)
// ======================================

/**
 * Creates a custom menu in Google Sheets for easy access to functions.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('NBU Exchange Rates')
    .addItem("Fetch Today's Rates", 'fetchCurrentRates')
    .addItem('Fetch Rates by Date', 'showDatePrompt')
    .addSeparator()
    .addItem('Update Rates from Date Column', 'updateRatesByDateColumn')
    .addSeparator()
    .addItem('Add Scheduled Updates', 'createScheduledTrigger')
    .addItem('Remove Scheduled Updates', 'removeScheduledTrigger')
    .addSeparator()
    .addItem('Clear Rate Cache', 'clearRateCache')
    .addToUi();

  // Run cache cleanup on open
  RateCache.cleanup();
}

/**
 * Shows a date picker dialog to fetch rates for a specific date.
 */
function showDatePrompt() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Fetch Exchange Rates by Date',
    'Enter date (DD.MM.YYYY):',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() == ui.Button.OK) {
    const dateStr = response.getResponseText();
    // Check if it's in DD.MM.YYYY format
    if (dateStr.match(/^\d{1,2}\.\d{1,2}\.\d{4}$/)) {
      // Convert from DD.MM.YYYY to DD/MM/YYYY for the parser
      const normalizedDate = dateStr.replace(/\./g, '/');
      const parsedDate = DateHandler.parse(normalizedDate);
      if (parsedDate) {
        // Convert to YYYY-MM-DD for internal processing
        const formattedDate = DateHandler.formatYYYYMMDD(parsedDate.date);
        fetchRatesForDate(formattedDate);
      } else {
        ui.alert('Invalid date. Please use DD.MM.YYYY format.');
      }
    } else {
      ui.alert('Invalid date format. Please use DD.MM.YYYY format.');
    }
  }
}

/**
 * Fetches current exchange rates from NBU.
 */
function fetchCurrentRates() {
  const today = new Date();
  const dateStr = DateHandler.formatYYYYMMDD(today);
  fetchRatesForDate(dateStr);
}

/**
 * Fetches exchange rates for a specific date from NBU API.
 *
 * @param {string} dateStr - Date in YYYY-MM-DD format
 */
function fetchRatesForDate(dateStr) {
  const parsedDate = DateHandler.parse(dateStr);
  if (!parsedDate) {
    SpreadsheetApp.getUi().alert('Invalid date format. Please use YYYY-MM-DD format.');
    return;
  }

  const apiDateStr = parsedDate.apiDateStr;
  const rates = NbuApi.fetchRates(apiDateStr);

  if (!rates) {
    SpreadsheetApp.getUi().alert('Could not find USD or EUR exchange rates in the data.');
    return;
  }

  // Update the sheet with the exchange rates
  SheetHandler.updateActiveSheet(dateStr, rates.usd, rates.eur);

  // Log success
  Logger.log('Successfully fetched rates for ' + dateStr);
  SpreadsheetApp.getUi().alert('Exchange rates updated for ' + dateStr);
}

/**
 * Updates USD NBU and EURO NBU columns based on Date column.
 */
function updateRatesByDateColumn() {
  SheetHandler.updateRatesByDateColumn();
}

/**
 * Creates a daily trigger to automatically update exchange rates.
 */
function createScheduledTrigger() {
  TriggerManager.createScheduled();
}

/**
 * Removes all scheduled triggers for exchange rate updates.
 */
function removeScheduledTrigger() {
  TriggerManager.removeScheduled();
}

/**
 * Clears the entire rate cache on user request
 */
function clearRateCache() {
  try {
    const clearedCount = RateCache.clearAll();
    SpreadsheetApp.getUi().alert(
      'Cache cleared. Removed ' + clearedCount + ' cached exchange rates.'
    );
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error clearing cache: ' + error.toString());
  }
}
