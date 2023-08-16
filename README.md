<picture>
 <source media="(prefers-color-scheme: dark)" srcset="YOUR-DARKMODE-IMAGE">
 <source media="(prefers-color-scheme: light)" srcset="YOUR-LIGHTMODE-IMAGE">
</picture>

# VBA-challenge
## Module 2 VBA Challenge for University of Birmingham Data Analytics Bootcamp
### Description
-------------------------------------------------------------------------------------------------------------------------------------------------
During this project I tested my VBA skills on fictious stock data collected between 2018 and 2020. The VBA modules include loops that will loop through all the stocks for each year (sheet) and output the following information:
  - Ticker symbol
  - Yearly change from opening price at the beginning of a given year to the closing price at the end of that year
  - Percentage change from the opening price at the beginning of a given year to the closing price at the end of that year
  - Total stock volume of the stock for each ticker

There is an additional functionality which returns the stock with the greatest % decrease, greatest % increase, and largest total stock volume for a given year

The decision was made to split the script into separate modules to allow the script to run more efficiently.

### VBA Module descriptions
-------------------------------------------------------------------------------------------------------------------------------------------------
There are 3 modules in total and each module has the following function:
 - Module 1 - extracts the data listed above
 - Module 2 - applies conditional formatting to the yearly change
 - Module 3 - returns the stock with the greatest % decrease, greatest % increase, and largest total stock volume for a given year

### Screenshots
-------------------------------------------------------------------------------------------------------------------------------------------------
The attached screenshots show the summary of the results created by the script for each year, shown on each separate sheet.

### Requirements
-------------------------------------------------------------------------------------------------------------------------------------------------
You will need Developer on your Excel and have macros enabled to run this script.

### Credits
-------------------------------------------------------------------------------------------------------------------------------------------------
Thank you to the bootcamp study groups without which I wouldn't have figured out the "firstrow" conundrum!
