
# stock-analysis


## Overview Project

Steve wanted an analysis for a specific list of stocks available in 2017 and 2018 to help his parents in their investment. 
In order to help him, we create a macro that he can use in the future to run larger datasets for any year. 



## Results

### 2017 

![2017](https://user-images.githubusercontent.com/87447639/130153603-d5591cea-fc4b-413e-8c5c-cebcf032f52d.PNG)

For 2017, most of the shares showed a positive return on investments, the highest being "DQ" with a 199.4% return. Only the "TERP" stock showed a loss with -7.2%.
It can be concluded that there is no direct relationship between the % return on investment and the total daily volume of each share. For example, the stock "SPWR" is the one that has the highest daily volume but its performance return is one of the lowest.

### 2018

![2018](https://user-images.githubusercontent.com/87447639/130153627-8662922a-e257-40d0-ae6c-db04525ed1cd.PNG)

For 2018, only 2 stocks from the list of 12 show a return on investment. The highest named "RUN" with 84% and secondly "ENPH" with 81.9%. Both values showed a high daily volume in their performance, but as in 2017, a direct relationship between both variables could not be determined.

For Steve parent we can concluded that 2017 was a positive year for the market in those specific stocks that showed a positive investment returns. 
To have a complete understanding of the data, it is important to analyze what were the events and variables that could have occurred in both years.

## Original Script vs Refactored Script



For the original VBA script we can see how we create 3 loops to complete the data in the "Total Daily Volume" and "Return" columns. This script is designed to do so column by column using the conditionals "if". 

![code original](https://user-images.githubusercontent.com/87447639/130270595-10722c63-4484-45e9-a4d8-6e108de21713.PNG)

For the refactored scrip when we created the variables "tickerVolumes" "tickerStartingPrices " & "tickerEndingPrices' at the begin of the code, and we created conditionals using those variables the code ran faster. In this case we only need to create two conditionals.

![Variables](https://user-images.githubusercontent.com/87447639/130270642-84673f7b-ee26-4cfa-9311-e0fc104bca43.PNG)

![code refactored](https://user-images.githubusercontent.com/87447639/130270616-3a256101-07c7-480a-afb5-aa2d9f8afca8.PNG)

For the execution of times between the original script and the refactored one, we can see how the time can be optimized and obtain the fastest result for both years.

### 2017 

  - ORIGINAL
  ![2017 original](https://user-images.githubusercontent.com/87447639/130267598-802b0bac-7a19-4aed-993d-fd4d246be0a6.PNG)
 
  - REFACTORED
  ![VBA_Challenge_2017](https://user-images.githubusercontent.com/87447639/130153842-5ca1f68f-c2d3-43e7-8172-b204f3438f0b.PNG)

### 2018

  - ORIGINAL
 ![2018 original](https://user-images.githubusercontent.com/87447639/130267653-38fd3da7-168c-4fcd-928c-b43820a966a1.PNG)

  - REFACTORED
  ![VBA_Challenge_2018](https://user-images.githubusercontent.com/87447639/130153844-5b0ca36d-8362-4b15-8aea-bbc3f95dcf5f.PNG)
  
## Sumary 

For Steve's analysis, refactoring of the code was beneficial since after modifying the code, understanding it and placing it in a simpler way, we could notice how the time to process to run a macro it was significantly shortened. Additionally, by refactoring the code, any other developer can use it and understand it more easily.

On the contrary, when refactoring the code, it can be time consuming and it can take longer than desired. It is possible that errors are made or part of the code is erased without wishing to do so. That is why when modifying a code you must be clear about the objectives to change it and be very carefully to any changes.

Aditionally, another benefit of refactoring the VBA is that Steve could now use this macro in the future to run larger data sets for any year. Although it took a longer time to determine the new variables and change the original script, but the execution of times between the original script and the refactored one, we can see how the time can be optimized and obtain the fastest result for both years. It is important to be very careful when making any changes.

