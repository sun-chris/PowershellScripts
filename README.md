# PowershellScripts

## excelStock.ps1

---

What this script does is:

1. Download a megalist CSV of all active stock tickers and ETFs from alphavantage.co
2. Load that CSV into MS Excel
3. Convert the stock ticker column into a "Linked Data Object", that contains information about stock price, volume, etc.. from excel's data connections
4. Expand that linked data cell's properties into columns
5. Clean the data and save it back to CSV for use in python scripts and other things that aren't confined by excel

This is done because as far as I'm aware, this is the best way of requesting metadata about this many stock all at once in bulk. I'm then going forward and using this basic company data, to filter out a pool of 1500 companies to request full historic data for from alphavantage.co each week (Staying within their API limitation of 500 req/day on their free plan, and capturing data on friday/saturday/sunday)
