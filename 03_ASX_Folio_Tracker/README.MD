# Data Analytics Portfolio

Welcome to my portfolio. I am learning to build clean data models and reports using Excel and Power BI.

---

## Project 3: ASX Stock Tracker and Analysis (Excel)

This project tracks a list of stock trades on the Australian Securities Exchange (ASX).

### How the "Fact Trades" Sheet Works

*   **Columns in this table:** Date, Ticker, Account, Transaction Type, Investment Type, Quantity, Price Per Share, Total Amount, XIRR Cashflow, Investment Occasion.

*   **Data Validation (Dropdown lists):**
    *   The user must pick values from lists for `Account`, `Transaction Type`, `Investment Type`, and `Investment Occasion` so there are no spelling mistakes.
    *   `Quantity` must be at least 1 unit.
    *   `Ticker` can be a maximum of 6 characters.

*   **Formulas Used:**
    1.  **Total Amount:** `=IF(A2="", "", F2*G2)`
        *   *Explanation:* If the Date cell is empty, it does not calculate. Otherwise, it multiplies Quantity (F) by Price Per Share (G).
    2.  **XIRR Cashflow:** `=IFS(E2="Dividend", 0, AND(D2="Buy", E2="Fresh"), -H2, AND(D2="Sell", E2="Fresh"), H2, TRUE, "")`
        *   *Explanation:* This counts the cash flow only when we buy stocks using fresh money (which is a negative value, `-H2`). If we are using money from dividends to buy stocks, we do not count it here.

*   **About the Data:**
    *   *Note: This data is for practice only. These are not stock recommendations.*
    *   I created a 161-row dataset starting from 2025 across 10 random ASX stocks. 
    *   I used Google Sheets and the `=GOOGLEFINANCE` function to get the historical stock prices for those dates.
    *   I connected the Google Sheet to Excel to pull the prices using an `XLOOKUP` formula.

## Fact Trade Table Image
![Fact Trade](images/Fact Trades.png)
