import yfinance as yf
import pandas as pd
from datetime import datetime, timedelta

filePath = "C:\\Yohan_CS_Work\\Python\\project\\topGainer\\TopGainer.xlsx"
filePathOutput = "C:\\Yohan_CS_Work\\Python\\project\\topGainer\\TopGainerOutput.xlsx"

df = pd.read_excel(filePath, engine='openpyxl')
largeListForAllgainers = []

date = '2024-06-12'
date = datetime.strptime(date, '%Y-%m-%d')
sum = 0

for i in range(0,10):
    first_column = df.iloc[:, 0].dropna().astype(str).tolist()

    time_to_add = timedelta(days=1)

    dateThatDontShowTime = date.strftime('%Y-%m-%d')

    endDate = (date + time_to_add).strftime('%Y-%m-%d')

    time_to_add = timedelta(days=8)

    endDateForTopGainers = (date + time_to_add).strftime('%Y-%m-%d')

    time_to_add = timedelta(days=9)

    endDateForEndTopGainers = (date + time_to_add).strftime('%Y-%m-%d')

    try:
        stock_data = yf.download(first_column, start=dateThatDontShowTime, end=endDate, group_by='ticker')
    except Exception as e:
        stock_data = {}

    print(stock_data)
    allInformationWithOpenClose = []
    for ticker in first_column:
        if ticker in stock_data:
            ticker_data = stock_data[ticker]
            if not ticker_data.empty:
                open_price = ticker_data['Open'].iloc[0]
                close_price = ticker_data['Close'].iloc[0]
                allInformationWithOpenClose.append([date,ticker, open_price, close_price])
        

    for j in range(0,len(allInformationWithOpenClose)):
        allInformationWithOpenClose[j].append(100*(allInformationWithOpenClose[j][3] - allInformationWithOpenClose[j][2])/(allInformationWithOpenClose[j][2]))

    allInformationWithOpenClose.sort(key=lambda x: x[4], reverse=False)
    allInformationWithOpenClose.sort(key=lambda x: x[4], reverse=True)


    for j in range(min(5, len(allInformationWithOpenClose))):
        largeListForAllgainers.append(allInformationWithOpenClose[j])
        stock = yf.Ticker(allInformationWithOpenClose[j][1])
        historical_data = stock.history(start=endDateForTopGainers, end=endDateForEndTopGainers)
        if not historical_data.empty:
            largeListForAllgainers[j+sum].append(historical_data['Close'][0])
            largeListForAllgainers[j+sum].append(100*(largeListForAllgainers[j][5] - largeListForAllgainers[j][3])/largeListForAllgainers[j][3])

    sum += min(5, len(allInformationWithOpenClose))
    


    date += timedelta(days=1)

print(largeListForAllgainers)

df_output = pd.DataFrame(largeListForAllgainers, columns=["Date","CompanyName", "Open", "Close","PercentGain","7 Days After Price","Percent Gain or Loss after 7 days"])
df_output.to_excel(filePathOutput, index=False, engine='openpyxl')
