# YFinance-Stocks-Data

### The code uses Yahoo Finance's python API (yfinance) to fetch stock data from the portal. 
- User shares the stock tickers and the attributes for which data is needed in a specific excel format (the excel format is available in the 'input' folder). 
- The program asks the user for the date the user needs stock data for, fetches the full data (yfinance library doesnt have a feature to pull data for a specific date), filters the data as per the date and writes the data back to the excel document.
- There are two sheets in the input excel - Stock_Current and Stock_History. User provides the ticker and attribute details in the Stock_Current sheet.
- The data is written back to the Stock_Current sheet as an overwrite. Also the data is appended to the existing data of the Stock_History sheet, so that this sheet maintains a history of all the data fetched.
