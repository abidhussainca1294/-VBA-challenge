# VBA-challenge
# Multiple Year Stock Data

In this challenge, a script is created that loops through all the stocks for one year and outputs the following information:

1. The ticker symbols

2. Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.

3. The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.

4. The total stock volume of the stock.

Screenshots have been attached of the results.

This Macro loops through all the worksheets in the Excel File for different year.

Conditional formatting is also performed on Yearly Change and Percent Change.

A summary table is created for getting the greatest total volume, greatest percentage increase and greatest percentage decrease. The Macro provides the values and the corresposnidng ticker in the Summary table for all the worksheets.

# Note
'For Each' is used for looping worksheets.

'For loop' is used for running through the each row in a column.

Nested if statements are used.

An alternative method to derive maximum value in a column is used for greatest total volume and identifying the correspoding row.
