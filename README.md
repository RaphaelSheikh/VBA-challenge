# VBA-challenge

![image](https://github.com/RaphaelSheikh/VBA-challenge/assets/166172978/fe4bee0f-28e6-4f2b-9899-33458a6e8563)

The project involves creating a VBA script to analyze stock market data using two workbooks: Test Data and Stock Data. The Test Data workbook, used during script development, contains six sheets labeled A-F and is smaller in size for testing purposes. The Stock Data workbook, containing the main data, comprises four sheets categorized by quarters (Q1, Q2, Q3, and Q4) and is larger in size. Data is sourced into Microsoft Excel, and VBA scripts are available in both workbook directories. Running the script on the Stock Data workbook may take some time due to its larger size.

Solution:
The script iterates through all the stock data once and displays the following information:
- The ticker symbol
- Quarterly change from opening price at the beginning of a given quarter to the closing price at the end of that quarter
- The percent change from opening price at the beginning of a given quarter to the closing price at the end of that quarter
- The total stock volume of the stock

The script will also identify key performers in the dataset by reporting stocks with the Greatest % Increase, Greatest % Decrease, and Greatest Total Volume. This will offer insights into significant movements within the market over the quarter.

To provide a comprehensive analysis, the script will run across multiple worksheets, enabling analysis of data from different quarters in one execution. This approach aims to offer a holistic view of stock performance trends over time, enhancing decision-making for investors.

References:
- Data for this dataset was generated by edX Boot Camps LLC

- Color Palette, Excel: (http://dmcritchie.mvps.org/excel/colors.htm)

- Percentage change formula: Investopedia (https://www.investopedia.com/terms/p/percentage-change.asp#:~:text=How%20Do%20I%20Calculate%20Percentage,multiply%20that%20number%20by%20100.)

- Additional formatting: Microsoft VBA Documentation (https://learn.microsoft.com/en-us/office/vba/api/overview/)

Screenshots:

Q1:

![image](https://github.com/RaphaelSheikh/VBA-challenge/assets/166172978/00e934c8-7a6d-4e19-8a81-c7718c8ca64d)

Q2:

![image](https://github.com/RaphaelSheikh/VBA-challenge/assets/166172978/2b358ede-374e-45e7-892a-0b9796bab681)

Q3:

![image](https://github.com/RaphaelSheikh/VBA-challenge/assets/166172978/34c78787-38af-4d4d-b81e-242c08c56ba3)

Q4:

![image](https://github.com/RaphaelSheikh/VBA-challenge/assets/166172978/14245253-e92f-4a4d-8eda-c9b23166ed44)
