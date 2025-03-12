import yfinance as yf
from openpyxl import load_workbook
from openpyxl.chart import LineChart, Reference, AreaChart

stockData = "empty" #Declare stockData outside an if statement

print("--------------------------------------------------------------------")
print(" _____                  ___  ___              _                 _   ")
print("|  __ \\                 |  \\/  |             | |               | |  ")
print("| |  \\/_ __ __ _ _   _  | .  . | ___ _ __ ___| |__   __ _ _ __ | |_ ")
print("| | __| '__/ _` | | | | | |\\/| |/ _ \\ '__/ __| '_ \\ / _` | '_ \\| __|")
print("| |_\\ \\ | | (_| | |_| | | |  | |  __/ | | (__| | | | (_| | | | | |_ ")
print(" \\____/_|  \\__,_|\\__, | \\_|  |_/\\___|_|  \\___|_| |_|\\__,_|_| |_|\\__|")
print("----------------------Your personal advisor-------------------------\n")

company = input("What company's stocks do you want to look at? (Enter their index)\n") #Select company
validPeriods = {"5d", "1mo", "3mo", "6mo", "1y", "2y", "5y", "10y", "ytd", "max"} #Creates an array of valid periods for error handling later

while True:
    dateOrPeriod = input("Choose whether you want to see results for a specific time range (r) or for a period (p)\n") #Choose the data display mode (by period or range)
    if dateOrPeriod == "p": #For period selected
        while True:
            period = input("Select a period: (Available periods are 5d, 1mo, 3mo, 6mo, 1y, 2y, 5y, 10y, ytd, max)\n") #Select period
            if period in validPeriods:
                stockData = yf.download(company, period=period, auto_adjust=True) #Grab financial data
                break
            else:
                print("Invalid input, please try again\n")  # loops the code if an invalid statement is given
    elif dateOrPeriod == "r": #For range selected
        startDate = input("Select a start date (format: Year-Month-Date)\n")
        endDate = input("Select an end date (format: Year-Month-Date)\n")
        stockData = yf.download(company, start=startDate, end=endDate, auto_adjust=True)
    else:
        print("Invalid input, please try again") #loops the code if an invalid statement is given

    fileName = f"{company}StockData.xlsx"

    if stockData.empty:
        print("Error: No data found for the given stock ticker. Please check the ticker and try again. If the error persists, check your internet connection.\n")
    else:
        print(stockData.head()) #prints the stock data
        stockData.to_excel(fileName, engine='openpyxl') #Send the data to an Excel sheet

    wb = load_workbook(fileName) #establish the workbook
    ws = wb.active #make the worksheet the active workbook

    chart = LineChart() #Creates the line chart
    chart.smooth = False
    chart.style = 12
    chart.height = 10.7
    chart.width = 22
    chart.title = "Stock price trend at market close"
    chart.x_axis.title = "Date"
    chart.y_axis.title = "Price"

    data = Reference(ws, min_col=2, min_row=4, max_row=ws.max_row) #Takes the data needed for the chart from the inputted data for the data
    dates = Reference(ws, min_col=1, min_row=4, max_row=ws.max_row) #same for the dates

    chart.add_data(data, titles_from_data=True) #Inputs the data

    for series in chart.series: #Makes the line appear smooth
        series.smooth = False

    ws.add_chart(chart, "H2") #Moves the chart to cell H2

    wb.save(fileName) #saves the current workbook under the above filename
    print(f"Data successfully saved to {fileName}")

    break
    #What to do next: Add conditional formatting automatically, start brainstorming new features
