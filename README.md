# VBA-challenge
VBA scripting to analyze real stock market data.

In this repository, the folder named 'resources' contains four files:
  1. alphabetical_testing.xlsx
  2. alphabetical_testing_GI.xlsm
  3. Multiple_year_stock_sata.xlsx
  4. Multiple_year_stock_data_GI.xlsm
  
  The files ending with xlsx are the original files while the ones ending with xlsm contains the macro developed and the subsequent elaboration. Alphabetical_testing was the training data set to develop the script while Multiple_year_stock contains real stock market data for the years 2016, 2015 and 2014.
  
  The macro code is contained in the file 'Sub VBA_Stocks.vbs'. The files 2014.png, 2015.png and 2016.png contains a screen shot of the results for each year on the woorkbook Multiple_year_stock_data_GI.xlsm.
  
  The script developed is reading every worksheet in the workbook and doing the following actions:
creating four columns containing: the ticker symbol, the calculated yearly change, percent change and total volume of the stock. The column containing the percent change is highlithing in green the cells containing a positive percent change and in red the ones containing a negative percent change. If both the opening price and closing price are zero then the macro is setting the cell value to 0 and the cell color to red. In contrast, if the opening price is zero but the closing price is not the cell value is set to N/A and the color is red or green if the closing price is negative or positive, respectively. 
Finally, the script is also returning the stock with the Greatest % increase, Greates % decrease and the Greatest total volume.
  
  
