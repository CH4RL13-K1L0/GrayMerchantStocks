import yfinance as yf
from openpyxl import load_workbook
from openpyxl.chart import LineChart, Reference, AreaChart
stockData = "empty"

print("--------------------------------------------------------------------")
print(" _____                  ___  ___              _                 _   ")
print("|  __ \\                 |  \\/  |             | |               | |  ")
print("| |  \\/_ __ __ _ _   _  | .  . | ___ _ __ ___| |__   __ _ _ __ | |_ ")
print("| | __| '__/ _` | | | | | |\\/| |/ _ \\ '__/ __| '_ \\ / _` | '_ \\| __|")
print("| |_\\ \\ | | (_| | |_| | | |  | |  __/ | | (__| | | | (_| | | | | |_ ")
print(" \\____/_|  \\__,_|\\__, | \\_|  |_/\\___|_|  \\___|_| |_|\\__,_|_| |_|\\__|")
print("----------------------Your personal advisor-------------------------\n")

company = input("What company's stocks do you want to look at? (Enter their index)\n")
while True:
    dateOrPeriod = input("Choose whether you want to see results for a specific time range (r) or for a period (p)\n")
    if dateOrPeriod == "p":
        period = input("Select a period: (Available periods are 1d, 5d, 1mo, 3mo, 6mo, 1y, 2y, 5y, 10y, ytd, max)\n")
        intervalHour = "1d"
        if period== "1d":
            intervalHour = "1h"
        stockData = yf.download(company, period=period, auto_adjust=True, interval=intervalHour)
    elif dateOrPeriod == "r":
        startDate = input("Select a start date (format: Year-Month-Date)\n")
        endDate = input("Select an end date (format: Year-Month-Date)\n")
        stockData = yf.download(company, start=startDate, end=endDate, auto_adjust=True,)
    else:
        print("Invalid input, please try again")

    print(stockData.head())
    fileName = f"{company}StockData.xlsx"
    stockData.to_excel(fileName, engine='openpyxl')

    wb = load_workbook(fileName)
    ws = wb.active

    chart = LineChart()
    chart.smooth = False
    chart.style = 12
    chart.height = 10.7
    chart.width = 22
    chart.title = "Stock price trend at market close"
    chart.x_axis.title = "Date"
    chart.y_axis.title = "Price"

    data = Reference(ws, min_col=2, min_row=4, max_row=ws.max_row)
    dates = Reference(ws, min_col=1, min_row=4, max_row=ws.max_row)

    chart.add_data(data, titles_from_data=True)
    chart.set_categories(dates)

    for series in chart.series:
        series.smooth = False

    ws.add_chart(chart, "H2")

    wb.save(fileName)
    print(f"Data successfully saved to {fileName}")
    break
