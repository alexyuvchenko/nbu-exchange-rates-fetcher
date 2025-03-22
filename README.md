# NBU Exchange Rate Fetcher

This project provides a solution for fetching exchange rates (USD-Hryvnia, EUR-Hryvnia) from the National Bank of Ukraine (NBU) for a specific date and integrating with Google Sheets.

## Features

- Fetch latest exchange rates from the National Bank of Ukraine
- Fetch exchange rates for a specific date
- Google Apps Script integration for Google Sheets
- Scheduled updates

## Solution Overview

### Google Sheets Integration

- `nbuExchangeRatesFetcher.js` - Google Apps Script to fetch rates directly within Google Sheets

## Prerequisites

- Google account
- Access to Google Sheets
- No programming knowledge required

## Installation & Usage

1. Create a new Google Sheet
2. Open the Script Editor (Extensions > Apps Script)
3. Copy the contents of `nbuExchangeRatesFetcher.js` into the script editor
4. Save the project
5. Refresh your Google Sheet
6. Use the custom menu "NBU Exchange Rates" that appears:
   - Select "Fetch Today's Rates" to get current rates
   - Select "Fetch Rates by Date" to get rates for a specific date
   - Select "Add Scheduled Updates" to set up automatic daily updates

## Data Format

The output format in your Google Sheet:

| Date     | YYYY-MM-DD |
| -------- | ---------- |
| Currency | Rate (UAH) |
| USD      | [USD rate] |
| EUR      | [EUR rate] |

## Troubleshooting

- **HTTP Error**: Check your internet connection or try again later

## License

This project is licensed under the MIT License - see the LICENSE file for details.
