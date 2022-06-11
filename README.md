# stock-analysis

## Overview/Purpose of Project:
   * The purpose of this project was to help Steve analyze the stock market for his parents by providing/refactoring a VBA code/script that is clean, efficient and minimizes the time to run it in an excel workbook. Moreover, the code would enable his parents to analyze stocks from other years making the code even more useful. Below is a brief snapshot/summary of our results as we compare the stock performance between 2017 and 2018, and demonstrate the decrease in execution times of the refactored script when compared to the original script.

## Results:
   * After refactoring the code, it became obvious that the "**For Loop**" was the critical element that would enable us and Steve's parents to analyze the stock data with minimal effort. Moreover, The For Loops we implemented allowed us to quickly examine stock data associated with an array (list) of 12 total tickers and effortlessly calculate total daily volume and return for each ticker and most importantly apply it to different years of data. Below are examples of two critical For Loops we used to make this possible. 
     * **Looping Over Rows:** In the below example, we created a For Loop that combs through the spreadsheet by looping over all the rows and with the help of conditionals, we were able to determine and store the starting and ending prices for the desired year. See below.
       * ![For_Loop_Over_Rows](Resources/For_Loop_over_rows.png)
     * **Looping through arrays to output Ticker, Total Daily Volume, and Return:** The next critical For Loop we created enables us to ouput the desired result (Total Daily Volume and Return). Following this step, we were able to create a code to that would apply formatting for the worksheet "All Stocks Analysis." Notice towards the end of the code there is an "endTime = Timer" followed by a Msg Box dictating the code run time with the year value. We will get in more detail on this later. For now, reference the below image.
       * ![For_Loop_through arrays](Resources/FOR_LOOP_ARRAYS_OUTPUT.png)

      
