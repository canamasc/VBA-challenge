# VBA-challenge
Georgia Tech Data Science Bootcamp

Prologue: 
  - I accidentally did not commit frequently as I was developing - habit I am still trying to form. I did most of the functionality in the alphabet_testing file, writing a segment, then running it to test and then debug before I moved onto the next segment.

Main differences seen in the multi_year VBA file:
  1. Parent sub whose only job is to call the main Sub for each worksheet in the workbook. Key here is to make sure the current worksheet is activated, so we correctly call the macro on it.
  2. I noticed my original code was coloring in the header for YEARLY CHANGE column in my formatting rules, so I changed the Range definition for this
  3. In the for loop, where I compute percent change, I did not run into a DIV#0! error in alphabet_test file but I did in the Multi year file, so I made this part a bit more robust in the final VBA file. Basically just catching the error before VBA tries to do the math, and filling the cell in with value 0 if there is an error.
  
 - In plain English:
  Given yearly stock info (spanning multiple years), return:
  - Unique tickers
  - Yearly price change
  - Yearly percent change
  - total stock volume
  - Stock with greatest % increase
  - Stock with greatest % decrease
  - Stock with greatest total volume
  
 
 
 PSEUDOCODE:
 First block of code sets up headers.
 Fill in ticker column using column A, then remove duplicates.
 Find out how many rows are in this new column, name this value cell2.
 For each row (ticker) in ticker column, do the following:
    1. Find the last occurrence of that ticker in column A. We can do this because the rows are sorted chronologically, so we know that the last instance of the ticker will be the last date of the year. xlWhole used to make sure we are matching the string exactly. Save this row number as temprow.
    2. Find first occurence of same ticker, and save this row number as temprow1. 
    3. Yearly change = closing value at temprow (column 6) - opening value at temprow1 (col 3).
    4. Percent change = yearly change / opening value at temprow 1. If for some reason, opening value at temprow1 = 0, percent change also  = 0.
    5. Total stock volume is equal to the sum of every row in the <vol> column where Ticker value = tick. SUMIF accomplishes this - if col A = tick, sum over col G. 
  

 The only thing left to do next is conditional formatting of yearly change column. We want to exclude row 1 since this is the header, so we go from row 2 to row = cell2, which we know to be the total number of unique tickers, i.e. how many rows our yearly change column will have.
  We set rules for this range: if cell value < 0, make the cell color red. If cell value > 0, make cell color green. Cell values that equal 0 are left with blank interior color.
  We format percent change column as percent with 2 decimal places to match instructor example.
  
  Bonus:
  Finding the values is straightforward, as we can use MAX/MIN functions on the columns we just created. To find the ticker symbols they correspond to, we could use VLOOKUP, but I found this to execute upsettingly slowly, so I made a loop using a variant array to speed things up. (Found a good example on the internet, it is cited in the code in a comment.)
  Basically, for each max/min value we just found, we loop through the range defined by inarr (includes the 4 new columns we previously made), and check to see if percent change (or total stock volume) is equal to the value we found to be the max/min. If so, then we want to fill in the cell next to the max/min with the value found at the same row, but in the ticker column.
  This is done 3 times for each of the 3 bonus values found.
  
  Application.Screenupdating is set to FALSE at the start of program and to TRUE at end - helps speed things up. 
  I auto-fit the columns at the end because I think it looks neater.
  
  
  
  
  
    
 
 
