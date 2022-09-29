### Description


FinancialAnalyzer is a Tkinter app that produces an excel file with financial statement analysis of chosen publicly traded company by using a python library called xlwings.
The file is then saved in the Downloads folder on your computer.



### Primary Languages

Python

### Technologies used

tkinter

pandas

requests

xlwings

os 

### How to use the app

1. First you will need an API key from financialmodelingprep.com

2. Input the api key.

```
response = requests.get("https://financialmodelingprep.com/api/v3/income-statement/{}?limit=120&apikey=[your api key]".format(company))
response_bs = requests.get("https://financialmodelingprep.com/api/v3/balance-sheet-statement/{}?limit=120&apikey=[your api key]".format(company))
```

3. Run the program and enter a stock ticker.

![image](https://user-images.githubusercontent.com/67132647/192931578-5827817a-cbab-453f-81ac-33bd1ca24371.png)

4. After hitting the Show Analysis button the excel sheet will appear in the Downloads folder as company_Analysis.xlsx










