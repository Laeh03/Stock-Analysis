# Stock-Analysis

## __Overview of Project__
The main goal of this macro-enabled workbook is to provide an analysis tool for any previous or future stocks. This tool aims to provide a quick overview of any stocks' total volume traded within a given period and its percent returns. 
We built two versions of this project, both provide analysis with the second version being more efficient. Our goal is to provide insight on why _Refactoring_ code is important in analyzing data.

![Screenshot (122)](https://user-images.githubusercontent.com/64225504/134261574-fce315ed-8b66-4b49-a90e-b9454aad8057.png)
![2018](https://user-images.githubusercontent.com/64225504/134261343-4d45f423-57a1-4018-b710-29a974d8bfa9.png)



## __Results__

For the stock performance analysis, 2017 was generally better vs. 2018. In 2017, outstanding performances can be seen with 'DQ' & 'SEDG' nearly gaining a 200% increase in return and almost every other stock saw positive increases except for 'TERP.' In terms of volume traded, top performer is 'SPWR' with 'FSLR' running in second, which may mean these were popular with the public. With Steve's parents pestering him about 'DQ's stock, looking at 2017 alone, it might be a good idea to invest in it however a year's worth of data might not be enough.

In 2018, almost every stock had a negative outlook except for 'ENPH' & 'RUN.' Every stock either saw an increase or decrease in traded volumes but these two said stocks had the largest increases which could be attributed to its success. If these stocks had positive outlooks for 2 consecutive years, we can definitely suggest to invest in 'ENPH' & 'RUN' to Steve's parents.

## Original vs. Refactored
The key difference between the two VBA scripts is the use of indices and arrays. On the first code, we loop through each tickers on 1 (one) array to perform the analyis. The code runs smoothly giving accurate results in analyzing returns & total volumes traded on any given stock. 

## __Summary__
- Refactoring code gives the possibilty of running it more efficient, giving shorter runtimes on analyzing data. Refactoring will definitely help on larger sets of data but dealing with a small dataset, it might not be worth a coder's time. The given dataset is a great example, with the original code providing results in less than a second even if the refactored code slashes the runtime by fractions.
- Refactoring an existing _VBA script_ can give you a side-by-side comparison between it & the existing script. One can create a roadmap on which areas of code could be improved increasing efficiency by creating more arrays, cleaner coding & etc. However, if one does not fully understand the syntax of an existing code, there will be a chance of making a perfectly running script into an unstable & unusable one.

