import yfinance as yf
import pandas as pd
from datetime import datetime, timedelta
from operator import itemgetter
from multiprocessing import Pool

filePath = "your file path"
filePathOutput = "your file path"

def load_data():
    return pd.read_excel(filePath, engine='openpyxl')

def process_date(date_str):
    df = load_data()  # Load data inside the worker function
    date = datetime.strptime(date_str, '%Y-%m-%d')
    first_column = df.iloc[:, 0].dropna().astype(str).tolist()

    time_to_add = timedelta(days=1)
    dateThatDontShowTime = date.strftime('%Y-%m-%d')
    endDate = (date + time_to_add).strftime('%Y-%m-%d')
    time_to_add = timedelta(days=7)
    endDateForTopGainers = (date + time_to_add).strftime('%Y-%m-%d')
    time_to_add = timedelta(days=8)
    endDateForEndTopGainers = (date + time_to_add).strftime('%Y-%m-%d')

    try:
        stock_data = yf.download(first_column, start=dateThatDontShowTime, end=endDate, group_by='ticker')
    except Exception as e:
        stock_data = {}

    allInformationWithOpenClose = []
    for ticker in first_column:
        if ticker in stock_data:
            ticker_data = stock_data[ticker]
            if not ticker_data.empty:
                open_price = ticker_data['Open'].iloc[0]
                close_price = ticker_data['Close'].iloc[0]
                if pd.notna(open_price) and open_price > 1:
                    allInformationWithOpenClose.append([date, ticker, open_price, close_price])

    for j in range(len(allInformationWithOpenClose)):
        allInformationWithOpenClose[j].append(100 * (allInformationWithOpenClose[j][3] - allInformationWithOpenClose[j][2]) / (allInformationWithOpenClose[j][2]))
    allInformationWithOpenClose = sorted(allInformationWithOpenClose, key=itemgetter(4), reverse=True)

    results = []
    for j in range(min(5, len(allInformationWithOpenClose))):
        row = allInformationWithOpenClose[j]
        stock = yf.Ticker(row[1])
        historical_data = stock.history(start=endDateForTopGainers, end=endDateForEndTopGainers)
        if not historical_data.empty:
            row.append(historical_data['Close'][0])
            row.append(100 * (row[5] - row[3]) / row[3])
        results.append(row)

    return results

def main():
    start_date = '2024-02-12'
    dates = [(datetime.strptime(start_date, '%Y-%m-%d') + timedelta(days=i)).strftime('%Y-%m-%d') for i in range(150)]
    
    with Pool(processes=4) as pool:
        all_results = pool.map(process_date, dates)
    
    # Flatten the list of results
    largeListForAllgainers = [item for sublist in all_results for item in sublist]

    totalPercentGain = 0
    averagePercentGain = 0
    for row in largeListForAllgainers:
        if len(row) > 6:
            totalPercentGain += row[6]
            averagePercentGain += 1

    largeListForAllgainers.append(["TotalPercentGain","","","","","",totalPercentGain])
    largeListForAllgainers.append(["AveragePercentGain","","","","","",totalPercentGain / averagePercentGain])

    df_output = pd.DataFrame(largeListForAllgainers, columns=["Date","CompanyName", "Open", "Close","PercentGain","Price after 7 days","Percent gain after 7 days"])
    df_output.to_excel(filePathOutput, index=False, engine='openpyxl')

if __name__ == '__main__':
    main()
