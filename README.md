# VBA-challenge

Module 2 VBA assignment submission for Data Analytics Bootcamp - UoT SCS

Overview
This VBA script analyzes stock data across multiple worksheets in an Excel workbook. For each worksheet, it calculates and summarizes quarterly stock changes, percentage changes, and total stock volume for each stock ticker. Additionally, it identifies the greatest percentage increase, greatest percentage decrease, and greatest total volume. The results are summarized in sections on each worksheet.

Features

Summary Table: Displays Ticker, Quarterly Change, Percentage Change, and Total Stock Volume.

Greatest Values: Identifies and displays the following,
Ticker with the greatest percentage increase.
Ticker with the greatest percentage decrease.
Ticker with the greatest total volume.

Conditional Formatting: Highlights positive quarterly changes in green and negative changes in red.


How It Works
Headers Setup: The script adds headers for the summary table and the greatest values section.

Data Processing:
Loops through the worksheets in the workbook and using the current worksheet - Q1 as the starting sheet.
Loops through rows of stock data, identifying when a new ticker starts or ends.
Tracks the opening price, closing price, and total volume for each ticker.
Calculates quarterly change and percentage change for each ticker.

Greatest Values: Tracks the tickers and their corresponding values for the greatest percentage increase, decrease, and total volume.

Results Output: Populates the summary table and greatest values section with the calculated data.

Conditional Formatting: Applies formatting to the "Quarterly Change" column to visually distinguish positive and negative changes.
