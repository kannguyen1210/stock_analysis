# Stock Analysis

### Project Overview
The Analysis was based on a set of stocks and their performance in 2017 and 2018. The Excel file was included with a VBA script to assist with analyzing the data set due to the large volume of data. The VBA code will provide the total daily volumes and annual returns of selected stocks and color-coded the data to easily identify which stocks provide positive returns in 2017 and 2018. 


### Results
#### Stock Performances 2017 versus 2018
From the screenshots of the stock performances below, we can see that 2017 was a great year as most of the stocks yielded a positive return except for "TERP" which had a negative 7.2% return. However, when going to 2018, we can clearly see it was not a good year for the market as most of the stocks had negative returns except for "ENPH" and "RUN". 
###### Stock Performance 2017
![alt text](https://github.com/kannguyen1210/stock_analysis/blob/main/Resources/Stock_Performance_2017.png)
###### Stock Performance 2018
![alt text](https://github.com/kannguyen1210/stock_analysis/blob/main/Resources/Stock_Performance_2018.png)


#### VBA Performance between Original script and Refactored script
From the screenshots of the script performances below, we can clearly see that the execution time of the refractored script is significantly lower than the original one for both 2017 and 2018. Therefore, we can conclude that the refractoring has been a success and script performances was improved.
###### Execution time 2017 - Original
![alt text](https://github.com/kannguyen1210/stock_analysis/blob/main/Resources/Before_Refactored_2017.png)
###### Execution time 2017 - Refractored
![alt text](https://github.com/kannguyen1210/stock_analysis/blob/main/Resources/VBA_Challenge_2017.png)
###### Execution time 2018 - Original
![alt text](https://github.com/kannguyen1210/stock_analysis/blob/main/Resources/Before_Refactored_2018.png)
###### Execution time 2017 - Refractored
![alt text](https://github.com/kannguyen1210/stock_analysis/blob/main/Resources/VBA_Challenge_2018.png)


### Summary
In general, the advantages of refractoring code is quite clear as it improves the design of the software in general, making it easier to debug as well adding further enhancement. It also help reduce the number of lines of codes to complete a certain features. However, it is not a best practice to refractoring code when you are close to the deadline or your budget is limited as the amount of time and effort going into refractoring could increases your project cost significantly.
For this specific exercise, we can see that the refractored version improves execution time significantly while we only have to maintain much fewer code. However, the new refractored script has a lot more variable that needs to be maintained when adding on more to the code in the future. It will certainly take more time to add additional features to the code to reuse these variables.
