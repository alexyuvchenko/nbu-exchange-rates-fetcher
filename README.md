# NBU Exchange Rate Fetcher

This project provides a solution for fetching exchange rates (USD-Hryvnia, EUR-Hryvnia) from the National Bank of Ukraine (NBU) for a specific date and integrating with Google Sheets.

## Features

- Fetch latest exchange rates from the National Bank of Ukraine
- Fetch exchange rates for a specific date
- Update multiple rows based on a date column
- Persistent caching to reduce API calls and improve performance
- Google Apps Script integration for Google Sheets
- Scheduled updates

## Solution Overview

### Google Sheets Integration

- `nbuExchangeRatesFetcher.gs` - Google Apps Script to fetch rates directly within Google Sheets

## Prerequisites

- Google account
- Access to Google Sheets
- No programming knowledge required

## Installation & Usage

1. Create a new Google Sheet
2. Open the Script Editor (Extensions > Apps Script)
3. Copy the contents of `nbuExchangeRatesFetcher.gs` into the script editor
4. Save the project
5. Refresh your Google Sheet
6. Use the custom menu "NBU Exchange Rates" that appears:
   - Select "Fetch Today's Rates" to get current rates
   - Select "Fetch Rates by Date" to get rates for a specific date in DD.MM.YYYY format
   - Select "Update Rates from Date Column" to fill in rates based on dates in your spreadsheet
   - Select "Add Scheduled Updates" to set up automatic daily updates
   - Select "Remove Scheduled Updates" to disable automatic updates
   - Select "Clear Rate Cache" to remove stored rate data

## Update Rates from Date Column

This powerful feature allows you to automatically populate exchange rates based on dates in your spreadsheet:

1. Create a sheet with columns named exactly:
   - "Date" - for your date values
   - "USD NBU" - will be filled with USD rates
   - "EURO NBU" - will be filled with Euro rates
2. Fill in the "Date" column with dates
3. Select "Update Rates from Date Column" from the menu
4. The script will automatically fill in the exchange rates for all dates

## Date Formats

The script supports various date formats:
- DD.MM.YYYY (e.g., 15.06.2023)
- DD/MM/YYYY (e.g., 15/06/2023)
- YYYY-MM-DD (e.g., 2023-06-15)

## Data Format for "Fetch Today's Rates" and "Fetch Rates by Date"

The output format in your Google Sheet:

| Date     | YYYY-MM-DD |
| -------- | ---------- |
| Currency | Rate (UAH) |
| USD      | [USD rate] |
| EUR      | [EUR rate] |

## Caching

The script implements persistent caching of exchange rates to:
- Improve performance
- Reduce the number of API calls
- Avoid hitting rate limits

Cached rates expire after 90 days to ensure data is kept fresh.

## Troubleshooting

- **HTTP Error**: Check your internet connection or try again later
- **Missing Columns**: When using "Update Rates from Date Column", ensure your columns are named exactly "Date", "USD NBU", and "EURO NBU"
- **Invalid Date Format**: Ensure dates are in one of the supported formats

## API Usage

This script uses the official NBU API at https://bank.gov.ua/NBUStatService/v1/statdirectory/exchange

## License

This project is licensed under the MIT License - see the LICENSE file for details.
