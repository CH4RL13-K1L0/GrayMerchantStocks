#Imports
from typing import Pattern
import yfinance as yf
from numpy.ma.core import choose
from openpyxl import load_workbook
from openpyxl.chart import LineChart, Reference, AreaChart
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import PatternFill
from openpyxl.styles import Font
from openpyxl.cell import cell
from copy import copy
from openpyxl.styles.builtins import styles
from openpyxl.chart.series import SeriesLabel
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.colors import ColorChoice

#Variables
stockData = "empty" #Declare stockData outside an if statement
period = 0
startDate = 0
endDate = 0
validPeriods = {"5d", "1mo", "3mo", "6mo", "1y", "2y", "5y", "10y", "ytd", "max"} #Creates a matrix of valid periods for error handling later
maxLength = 0
movingAverageOption = 0


print("""--------------------------------------------------------------------
 _____                  ___  ___              _                 _   
|  __ \\                 |  \\/  |             | |               | |  
| |  \\/_ __ __ _ _   _  | .  . | ___ _ __ ___| |__   __ _ _ __ | |_ 
| | __| '__/ _` | | | | | |\\/| |/ _ \\ '__/ __| '_ \\ / _` | '_ \\| __|
| |_\\ \\ | | (_| | |_| | | |  | |  __/ | | (__| | | | (_| | | | | |_ 
 \\____/_|  \\__,_|\\__, | \\_|  |_/\\___|_|  \\___|_| |_|\\__,_|_| |_|\\__|
----------------------Your personal advisor-------------------------
""")

company = input("What company's stocks do you want to look at? (Enter their ticker)\n") #Select company

while True:
    dateOrPeriod = input("Choose whether you want to see results for a specific time range or for a period. r/p\n") #Choose the data display mode (by period or range)
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
        print(f"{stockData.head()}\n") #prints the stock data
        while True:
            movingAverageOption = input("Do you want Exponential or Simple moving averages? E/S\n")
            if movingAverageOption == "S":
                stockData["SMA_10"] = stockData["Close"].rolling(window=10).mean() #Uses panda to calculate a simple moving average (SMA)
                stockData["SMA_50"] = stockData["Close"].rolling(window=50).mean()
                stockData["SMA_200"] = stockData["Close"].rolling(window=200).mean()
                break
            elif movingAverageOption == "E":
                stockData["EMA_10"] = stockData["Close"].ewm(span=10, adjust=False).mean() #Uses panda to calculate an exponential moving average (EMA)
                stockData["EMA_50"] = stockData["Close"].ewm(span=50, adjust=False).mean()
                stockData["EMA_200"] = stockData["Close"].ewm(span=200, adjust=False).mean()
                break
            else:
                print("Invalid argument, try again")
        stockData.to_excel(fileName, engine='openpyxl') #Send the data to an Excel sheet

    wb = load_workbook(fileName) #establish the workbook
    ws = wb.active #make the worksheet the active workbook

    for col in ws.columns:
        colLetter = col[0].column_letter
        for cell in col:
            if isinstance(cell.value, (int, float)): #Reformats decimal points if the cell contains a number
                cell.number_format = "0.00"
                maxLength = max(maxLength, len(f"{cell.value:.2f}")) #Track the max length of formatted numbers
            else:
                maxLength = max(maxLength, len(str(cell.value))) #Track the max length of non-numeric values

        ws.column_dimensions[colLetter].width = maxLength + 1 #Extends the cells to fit the contents

    ws["J1"] = "Change"
    ws["J1"].font = copy(ws["F1"].font)
    ws["J1"].alignment = copy(ws["F1"].alignment)
    ws["J1"].fill = copy(ws["F1"].fill)
    ws["G1"].border = copy(ws["F1"].border)
    ws.column_dimensions["J"].width = maxLength + 1

    if dateOrPeriod == "p":
        ws["B3"] = f"Data for the past {period}"
        ws["B3"].font = copy(ws["F1"].font)
        ws["B3"].alignment = copy(ws["F1"].alignment)
        ws["B3"].fill = copy(ws["F1"].fill)
        ws["B3"].border = copy(ws["F1"].border)
    else:
        ws["B3"] = f"Data from {startDate} to {endDate}"
        ws["B3"].font = copy(ws["F1"].font)
        ws["B3"].alignment = copy(ws["F1"].alignment)
        ws["B3"].fill = copy(ws["F1"].fill)
        ws["B3"].border = copy(ws["F1"].border)

    green = PatternFill(start_color="568203", end_color="568203", fill_type="solid")
    red = PatternFill(start_color="FF2800", end_color="FF2800", fill_type="solid")
    redFont = Font(color="FF2800", bold=True)
    greenFont = Font(color="568203", bold=True)

    column = "B"

    ws.conditional_formatting.add( #Adds conditional formatting
        f"{column}4:{column}{ws.max_row}",
        CellIsRule(operator="greaterThan", formula=["B3"], stopIfTrue=False, font=greenFont)
    )
    ws.conditional_formatting.add(
        f"{column}4:{column}{ws.max_row}",
        CellIsRule(operator="lessThan", formula=["B3"], stopIfTrue=False, font=redFont)
    )

    ws["G5"].value = "=B5-B4" #Calculates the change in stock price
    i = 5 #Starts from cell G5
    while i <= ws.max_row: #Loops through each cell in the column
        ws[f"J{i}"] = f"=B{i}-B{i - 1}"
        i += 1

    column2 = "J" #Reformatting values again here because data written by us is formatted separately from that of the program
    for row in ws[f"{column2}2:{column2}{ws.max_row}"]:
        for cell in row:
            cell.number_format = "0.00"

    ws.conditional_formatting.add(
        f"{column2}2:{column2}{ws.max_row}",
        CellIsRule(operator="greaterThan", formula=["0"], stopIfTrue=False, font=greenFont)
    )
    ws.conditional_formatting.add(
        f"{column2}2:{column2}{ws.max_row}",
        CellIsRule(operator="lessThan", formula=["0"], stopIfTrue=False, font=redFont)
    )

    chart = LineChart() #Creates the line chart
    chart.smooth = False
    chart.style = 12
    chart.height = 30
    chart.width = 21
    chart.title = "Stock prices with moving average"
    chart.x_axis.title = "Date"
    chart.y_axis.title = "Price"

    data = Reference(ws, min_col=2, min_row=4, max_row=ws.max_row) #Takes the data needed for the chart from the inputted data for the data
    movingAverageReference10 = Reference(ws, min_col=7, min_row=4, max_row=ws.max_row) #Creates a chart reference for the moving averages from before
    movingAverageReference50 = Reference(ws, min_col=8, min_row=4, max_row=ws.max_row)
    movingAverageReference200 = Reference(ws, min_col=9, min_row=4, max_row=ws.max_row)
    dates = Reference(ws, min_col=1, min_row=4, max_row=ws.max_row) #same for the dates

    chart.add_data(data, titles_from_data=False) #Inputs the data
    chart.add_data(movingAverageReference10, titles_from_data=False) #Inputs the moving averages onto the chart
    chart.add_data(movingAverageReference50, titles_from_data=False)
    chart.add_data(movingAverageReference200, titles_from_data=False)

    chart.series[0].tx = SeriesLabel(v="Stock Close") #Labels the legend
    chart.series[1].tx = SeriesLabel(v=f"10-day {movingAverageOption}MA")
    chart.series[2].tx = SeriesLabel(v=f"50-day {movingAverageOption}MA")
    chart.series[3].tx = SeriesLabel(v=f"200-day {movingAverageOption}MA")

    colors = ["000000", "FF2800", "008000", "FFA500"]
    for i, series in enumerate(chart.series): #Chart lines get visual properties
        series.smooth = False
        series.graphicalProperties.line.solidFill = colors[i]
        series.graphicalProperties.line.width = 17000

    ws["G5"].value = None #Removes an annoying bug

    ws.add_chart(chart, "K2") #Moves the chart to cell L2

    wb.save(fileName) #Saves the current workbook under the above filename
    print(f"Data successfully saved to {fileName}")

    break
    #What to do next: Add bollinger bands and other features 
