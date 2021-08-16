# Stock Data

This script updates a Google Sheets spreadsheet with stock ticker info from finviz.com. Options are 'Update Selected Stocks', and 'Update All Stocks'.

Usage:
- Set spreadsheet Column 1 to ticker names, Row A to field names in Finviz stock table (eg: 'SMA200', 'EPS Next Q')
- Open Google Script Editor and make a new script with the contents of `integration.gs.example`
- Add Cheerio library
- A 'Finance' tab will be added to the topbar of the spraedsheet, all done!