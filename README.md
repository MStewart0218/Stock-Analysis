# Stock-Analysis
VBA analysis of stock data

> ## Overview of Project:
> The purpose of this analysis is to provide end user, Steve, with a comprehensive analysis of Total Volume and Return of multiple stocks/ tickers. Steve is researching the stock data to advise his parents on the best stocks to invest in. Initially, Steve’s parents wanted to invest in “DQ” stock but an analysis of the entire dataset has shown that “DQ” may not be the best investment for Steve’s parents.

## Analysis:

The analysis was performed by creating a VBA script to analyze Total stock Volume and Stock Return percentages. The Total Stock Volume and Stock Return percentage compare 12 stock tickers from 2017 and 2018. 

## Questions to Consider:

####  - While comparing the stock performance between 2017 and 2018, what were the execution times of the original script and the refactored script?
####  - What are the advantages or disadvantages of refactoring code?
####  - How do these pros and cons apply to refactoring the original VBA script?

## Results:

The VBA script highlighted the following ticker performance form 2017 to 2018 while considering the original stock "DQ" Steve's parents wanted to invest in:

•	Ticker "ENPH" Total Daily Return increased by $385,701,400 from 2017 to 2018. "ENPH" return percentage decreased from 129.5% in 2017 to 81.9% in 2018. "ENPH" outperformed ticker "DQ" in 2018 making this a better stock for Steve's parents to invest in and gain positive return on investment.

•	Ticker "RUN" Total Daily Return increased by $235,075,800 from 2017 to 2018. "RUN" return percentage increased from 5.5% in 2017 to 84.0% in 2018. "RUN" outperformed tickers "DQ"  and "ENPH" in 2018 making this the best stock for Steve's parents to invest in and gain positive return on investment.

Advantages of refactoring code are streamlining the overall script and decreasing the script run-time to improve analysis performance. One disadvantage of refactoring code was the possibly of removing functionality of the code: consequently, creating errors and erroneous output.

The pros and cons of refactoring the original VBA script posed challenges such as mixing up the order of operation of the original steps which caused the script to pull partial data instead for providing an output of all 12 tickers in the dataset. Subsequently, once the order of operation was corrected in the refactored VBA script performance improved and the runtime was faster, than the original runtime, when analyzing the data.

**Original Runtime vs Refactored Runtime**

| Year | Original Runtime| Refactored Runtime
| ------ | ----------- | ------| 
| 2017 | 0.65625 | 0.625 |
| 2018 | 0.640625 | 0.6171875 |
    
## Documents:

1. `VBA Challenge Workbook` 
    [VBA_Challenge.xslx](https://github.com/MStewart0218/Stock-Analysis/blob/main/VBA_Challenge.xlsm)
 
