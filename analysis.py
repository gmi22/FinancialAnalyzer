import tkinter as tk
from tkinter import ttk
import requests
import pandas as pd
import xlwings as xw
import matplotlib.pyplot as plt
from matplotlib import dates as mpl_dates
import os



root = tk.Tk()






def graph():

    
        company = user_name.get()

        response = requests.get("https://financialmodelingprep.com/api/v3/income-statement/{}?limit=120&apikey=[your api key]".format(company))
        response_bs = requests.get("https://financialmodelingprep.com/api/v3/balance-sheet-statement/{}?limit=120&apikey=[your api key]".format(company))

        r = response.json()
        r_bs = response_bs.json()

        app = xw.App(visible=False)
        wb = xw.Book()  
        ws = wb.sheets['Sheet1']

        date = [r[i]['date'] for i in reversed(range(0,len(r))) ]
        revenue = [r[i]["revenue"] for i in reversed(range(0,len(r)))]
        grossProfitRatio = [r[i]["grossProfitRatio"] for i in reversed(range(0,len(r)))]
        researchAndDevelopmentExpenses = [r[i]["researchAndDevelopmentExpenses"] for i in reversed(range(0,len(r)))]
        grossProfit = [r[i]["grossProfit"] for i in reversed(range(0,len(r)))]
        sellingGeneralAndAdministrativeExpenses = [r[i]["sellingGeneralAndAdministrativeExpenses"] for i in reversed(range(0,len(r)))]
        depreciationAndAmortization = [r[i]["depreciationAndAmortization"] for i in reversed(range(0,len(r)))]
        netIncomeRatio = [r[i]["netIncomeRatio"] for i in reversed(range(0,len(r)))]
        netIncome = [r[i]["netIncome"] for i in reversed(range(0,len(r)))]
        operatingIncome = [r[i]["operatingIncome"] for i in reversed(range(0,len(r)))]
        interestExpense = [r[i]["interestExpense"] for i in reversed(range(0,len(r)))]
        cashAndCashEquivalents = [r_bs[i]["cashAndCashEquivalents"] for i in reversed(range(0,len(r_bs)))]

        ws.range('A1').options(index=False).value = company

        #revenue trend
        revenue_data = {'Date':date,
                'Revenue': revenue}

        df_revenue = pd.DataFrame(revenue_data)
        df_revenue['Revenue'] = df_revenue['Revenue'].apply('{:,}'.format)
        df_revenue['Date']= pd.to_datetime(df_revenue['Date'])
        ws.range('A25').options(index=False).value = df_revenue
        chart_revenue = ws.charts.add(left=0, top=100, width=400, height=211)
        chart_revenue.set_source_data(ws.range('A25').expand())
        chart_revenue.chart_type = 'line'
        chart_revenue.name


        gp_data = {'Date':date,
            'Gross Profit Margin': grossProfitRatio}
        df_gpm = pd.DataFrame(gp_data)
        df_gpm['Gross Profit Margin'] = df_gpm['Gross Profit Margin'].astype(float).map("{:.2%}".format)
        ws.range('J25').options(index=False).value = df_gpm
        chart_gpm = ws.charts.add(left=450, top=100, width=400, height=211)
        chart_gpm.set_source_data(ws.range('J25').expand())
        chart_gpm.chart_type = 'line'
        chart_gpm.name


        rd_data = {'Date':date,
            'R&D': researchAndDevelopmentExpenses}

        df_development = pd.DataFrame(rd_data)
        df_development['R&D'] = df_development['R&D'].apply('{:,}'.format)
        ws.range('A52').options(index=False).value = df_development
        chart_rd = ws.charts.add(left=0, top=500, width=400, height=211)
        chart_rd.set_source_data(ws.range('A52').expand())
        chart_rd.chart_type = 'line'
        chart_rd.name


        SGA_as_Percentage_of_GP_data = {'Date':date,'GP': grossProfit,'SGA': sellingGeneralAndAdministrativeExpenses}

        df_SGA_as_Percentage_of_GP_data = pd.DataFrame(SGA_as_Percentage_of_GP_data)

        df_SGA_as_Percentage_of_GP_data['SG&A as a % of Gross Profit'] = df_SGA_as_Percentage_of_GP_data['SGA']/ df_SGA_as_Percentage_of_GP_data['GP']

        df_SGA_as_Percentage_of_GP_data['SG&A as a % of Gross Profit'] = df_SGA_as_Percentage_of_GP_data['SG&A as a % of Gross Profit'].astype(float).map("{:.2%}".format)

        df_SGA_as_Percentage_of_GP_data.drop(columns=['GP', 'SGA'], inplace=True)

        ws.range('J52').options(index=False).value = df_SGA_as_Percentage_of_GP_data

        chart_rd = ws.charts.add(left=450, top=500, width=400, height=211)
        chart_rd.set_source_data(ws.range('J52').expand())
        chart_rd.chart_type = 'line'
        chart_rd.name


        #net earnings as a % of revenue
        net_income_as_perc_rev = {'Date':date,'Net Income as % of Revenue': netIncomeRatio}

        df_net_income_as_perc_rev = pd.DataFrame(net_income_as_perc_rev)

        df_net_income_as_perc_rev['Net Income as % of Revenue'] = df_net_income_as_perc_rev['Net Income as % of Revenue'].astype(float).map("{:.2%}".format)

        ws.range('A79').options(index=False).value = df_net_income_as_perc_rev

        chart_rd = ws.charts.add(left=0, top=900, width=400, height=211)
        chart_rd.set_source_data(ws.range('A79').expand())
        chart_rd.chart_type = 'line'
        chart_rd.name


        dep_costs_as_p_of_gp = {'Date':date,'Depreciation': depreciationAndAmortization, 'Net Income': grossProfit }

        df_dep_as_perc_rev = pd.DataFrame(dep_costs_as_p_of_gp)

        df_dep_as_perc_rev['Deprecation as a % of Gross Profit'] = df_dep_as_perc_rev['Depreciation']/df_dep_as_perc_rev['Net Income']

        df_dep_as_perc_rev.drop(columns=['Depreciation', 'Net Income'], inplace=True)

        df_dep_as_perc_rev['Deprecation as a % of Gross Profit'] = df_dep_as_perc_rev['Deprecation as a % of Gross Profit'].astype(float).map("{:.2%}".format)

        ws.range('J79').options(index=False).value = df_dep_as_perc_rev

        chart_rd = ws.charts.add(left=450, top=900, width=400, height=211)
        chart_rd.set_source_data(ws.range('J79').expand())
        chart_rd.chart_type = 'line'
        chart_rd.name


        #interest expense % of Operating income

        interest_expense_as_perc_gp = {'Date':date,'Interest Expense': interestExpense, 'Operating Income': operatingIncome }
        df_interest_expense_as_perc_gp = pd.DataFrame(interest_expense_as_perc_gp)
        df_interest_expense_as_perc_gp['Interest Expense as % of Operating Income'] = df_interest_expense_as_perc_gp['Interest Expense']/df_interest_expense_as_perc_gp['Operating Income']
        df_interest_expense_as_perc_gp.drop(columns=['Interest Expense', 'Operating Income'], inplace=True)
        df_interest_expense_as_perc_gp['Interest Expense as % of Operating Income'] = df_interest_expense_as_perc_gp['Interest Expense as % of Operating Income'].astype(float).map("{:.2%}".format)
        ws.range('A105').options(index=False).value = df_interest_expense_as_perc_gp

        chart_rd = ws.charts.add(left=0, top=1350, width=400, height=211)
        chart_rd.set_source_data(ws.range('A105').expand())
        chart_rd.chart_type = 'line'
        chart_rd.name

        for ws in wb.sheets:
            ws.autofit(axis="columns")


        wb.sheets[0].name = "Income Statement Analysis"

        ws_bs = wb.sheets['Sheet2']


        cash = {'Date':date,'Cash': cashAndCashEquivalents}
        df_cash = pd.DataFrame(cash)
        df_cash['Cash'] = df_cash['Cash'].apply('{:,}'.format)
        ws_bs.range('A25').options(index=False).value = df_cash


        chart_rd = ws_bs.charts.add(left=0, top=100, width=400, height=211)
        chart_rd.set_source_data(ws_bs.range('A25').expand())
        chart_rd.chart_type = 'line'
        chart_rd.name


        for ws in wb.sheets:
            ws_bs.autofit(axis="columns")


        downloads = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Downloads')

        wb.save(downloads + '\\' + '{}_Analysis.xlsx'.format(company))
        #wb.save('{}_Analysis.xlsx'.format(company))
        #wb.save('Analysis_test.xlsx')
        wb.close()
        app.quit()







    




user_name = tk.StringVar()

name_label = ttk.Label(root,text = "Ticker: ")
name_label.pack(side = "left",padx=(0,10))
name_entry = ttk.Entry(root,width=15,textvariable=user_name)
name_entry.pack(side = "left")
name_entry.focus()


greet_button = ttk.Button(root,text = 'Show Analysis',command = graph)
greet_button.pack(side="left")






root.mainloop()


