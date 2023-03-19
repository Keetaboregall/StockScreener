#importing required libraries
import pandas as pd
import openpyxl

#creating pandas dataframe from csv file containing stocks' data
stocks=pd.read_csv("stocks_data.csv",index_col=0)
print(stocks)

#considering stocks that have positive earnings per share(EPS) growth over the last 3 years
stocks=stocks[(stocks['eps_2022']>0) & (stocks['eps_2021']>0) & (stocks['eps_2020']>0)]
#considering stocks that have an annual cagr of over 10 percent
stocks=stocks[stocks['cagr']>10]
#considering stocks that have a price-to-earnings (P/E) ratio of less than 40
stocks=stocks[stocks['p/e']<40]
#list of potential stocks:
print(stocks)

#ranking stocks based on their growth rate divided by P/E ratio
stocks['growth/pe_ratio']=stocks['cagr']/stocks['p/e']
stocks['rank']=stocks['growth/pe_ratio'].rank(ascending=False)

#arranging stocks rank-wise
stocks=stocks.sort_values('rank')
stocks=stocks.reset_index(drop=True)
print(stocks)

#generating an excel report
stocks.to_excel("FinalReport.xlsx",index=False)
workbook = openpyxl.load_workbook('FinalReport.xlsx')
worksheet = workbook.active
worksheet.column_dimensions['A'].width = 20
worksheet.column_dimensions['B'].width = 10
worksheet.column_dimensions['C'].width = 10
worksheet.column_dimensions['D'].width = 10
worksheet.column_dimensions['E'].width = 10
worksheet.column_dimensions['F'].width = 10
worksheet.column_dimensions['G'].width = 20
worksheet.column_dimensions['H'].width = 10
workbook.save('FinalReport.xlsx')
print('Dataframe converted to Excel Report')