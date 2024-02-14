# VBA-challenge


## List of items in the repository:
* #### Solution Images Folder:
   Contains screenshots of the results.
* #### Multiple_year_stock_data_vbcode:
   VBA script file.

## Code Description.
Two data sets were provided for this challenge:

* Multiple_year_stock_data:

   This data set had details of the stocks for 3 years 2018,2019 and 2020.
   As per the instruction,these are the steps followed in the code:
  * created 4 columns named "Ticker","Yearly Change","Percent Change" and "Total Stock Volume".
  * Used "IF-Else" condition with in the "FOR LOOP" for comparing each row using of the ticker column to collect the unique ticker values and populated in the result in the "Ticker" column.
  * Find the closing price and opening price of each ticker and assigned the values to seperate variables "cl_price" and "op_price" respectively.
  * Subtracted "op_price" from "cl_price" for calculating "Yearly Change",and populated these values in the "Yearly Change" column.
  * Find the "Percent Change" using the equation "(year_chng / op_price) * 100" and used "ROUND" function for getting the values in 2 decimal places then 
   populated the values in the "Percent Change" column.
  * "Percent Change" column formated using "%" sign.
  * Conditional formating is applied on the column "Yearly Change" column used "Interior.ColorIndex " for filling the rows having negative values with Red color and those with positive values with Green color.
  * Total stock volume is calculated by adding the <vol> values of each <ticker> and populated in the column "Total Stock Volume".
  * Used "IF" condition to compare each row of the "Percent Change" column and "Total Stock Volume" column to find the maximum and minimum values of "Percent Change" and maximum value of "Total Stock Volume" .The 
    maximum values inserted in another column "Value" as "Greatest % Increase" and "Greatest Total Volume". The minimum  value as "Greatest % Decrease".
  * Found the row number of the maximum and minimum values and stored in the variable "row_num".
  * Using "row_num" found the corresponding "Ticker" of the maximum and minimum values and populated in the new "Ticker" column.
  * Finally applied "FOR LOOP" for the entire code for working the code across the worksheets in the workbook.
    
* Alphabetical_test data:
   As per the instruction the same code developed for Multiple_year_stock_data copied to this dataset.
   Since this data set has similar columns as Multiple_year_stock_data the code worked perfectly.

