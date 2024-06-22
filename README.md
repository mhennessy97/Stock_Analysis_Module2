# Stock_Analysis_Module2


First began by setting variables that will be used throughout code. 

need to add loop so code references each worksheet

then defined values for variables. Used 0 for each except for row = 2. Used this because data begins in row 2 and headers are in row 1. 

needed to insert headers in each worksheet. Important to begin code with ws. 

then needed to determine row count on each sheet to use with "row_index" loop. Named as "rowcount" and referenced cells then rows.count as row index in column A. the XLUP then counted to get the # all the way to last rows with row_index. 

For ticker names, used row_index + 1 so it will search through each row and referenced 1st column. Used <> (does not equal) row_index, 1. This searched through each row to see when the ticker changes in each set of data. 

Then used total to determine the total stock volume in each row index. 

If function determines total for each ticker. Used range formula within if function to populate values in columns I, J, K, L.

Else function after if searches through data to determine if loop should stop or move on. 

Change is calculates by close - open

Then percent change is determined by dividing change / (start(2), 3). Start was defined as row_index + 1 so will search through each row in column 3. 

Then used range formula by specifying the column & the row, + the column index to work through each column. Later specify column index = column index + 1. 

Put quarterly change in column J2 working through each row and set the number format as a decimal. Did the same for percent change, but included percentage symbol as well. 

Next used conditional formatting to select case and highlight Quarterly and Percent change columns. Used google to look up color index #s to highlight interior of cells. 

Needed to reset variables after end if for total, change, and column index. 

Added in else statement to set total if ticker is constant and loop through adding totals in column 7 on each worksheet. 


After next row_index, added in range functions to populate % increase in Q2, % decrease in Q3, and greatest total vol in Q4. 

First uses Q2 with the "%" symbol and worksheet function with Max finds highest value within specified range. Opposite for Min. Same for volume, just a different column and percentage symbol not needed.

Then needed to add values in final table set up in columns O, P, & Q. Used max & min with range formula to search through specified columns and rows. 
