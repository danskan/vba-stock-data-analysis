# vba-stock-data-analysis
VBA Script that Analyzes Stock Data and Outputs Tables Showing Various Stock Price Performance


This vba subroutine traverses stock data and outputs some key statistics, such as Annual Price Change, Percent Change, and Total Volume.  Then it creates a summary table that shows the best and worst performers by annual price percent change, as well as reports the ticker and volume value for the highest volume stock in each year.

Potential Issues with this Data

- If a stock starts trading after the first trading day of the year, the script sets the percentage change to zero, because the opening price shows zero.  Instead, the script should find the opening price of the stock when it does start trading and use that opening price as the real opening price.  This is a complication in the way that the script traverses the data.  Specifically in the case of PLNT, which does not start trading until part way through the year, the script calculates the percent change from zero to it's closing price.  Depending on the who is reading the report, they may want it to read from it's opening price (midway through the year, after about 100 lines of zeros for opening price values) of over 14 dollars, which would produce a much smaller resulting percentage change, and is likely more correct.  Given that there are 980 IPOs every year, this data may need to be cleaned further prior to running the script.
