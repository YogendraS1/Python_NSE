# Python_NSE
Yogendra

Using Python and a combination of libraries such as requests, pandas, openpyxl, and lxml, I developed a script that scrapes data from Moneycontrol's mutual 
fund performance tracker. This script retrieves performance data for various fund categories, including Large Cap, Large and Mid Cap, ELSS, Focused, and more. 📈
During the development process, I encountered several challenges. Firstly, I had to handle the dynamic nature of the website's HTML structure, which required 
me to employ regex and the lxml library to preprocess the HTML source. Additionally, I had to extract data from multiple tables on each webpage and consolidate 
it into a meaningful format for further analysis. 🧩
To overcome these challenges, I leveraged the power of pandas and openpyxl libraries. I used pandas' read_html function to extract tables from the 
HTML source and transformed the data into pandas dataframes. Then, I utilized openpyxl to create a workbook, create sheets for each fund category, 
and populate the data in an organized manner. 💪
The resulting code generates an Excel workbook, "MutualFunds.xlsx," with sheets for each fund category and a separate sheet, "Best Mutual Funds," 
that highlights funds meeting specific criteria. The criteria for selection include a 3-year return greater than 20%, a Crisil rank of 1 or 2, and
filtering out sponsored funds. 📊
By utilizing this code, investors and financial enthusiasts can quickly identify the best-performing mutual funds based on predefined criteria. 
This not only saves time but also enables informed investment decisions. 💼
