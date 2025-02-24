from fileinput import filename
import yfinance as yf
import os
import pandas as pd
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
        print(stockData.head())
        file_name = f"{company}StockData.xlsx"
        stockData.to_excel(file_name, engine='openpyxl')
        print(f"Data successfully saved to {file_name}")
        break
    elif dateOrPeriod == "r":
        startDate = input("Select a start date (format: Year-Month-Date)\n")
        endDate = input("Select an end date (format: Year-Month-Date)\n")
        stockData = yf.download(company, start=startDate, end=endDate, auto_adjust=True,)
        print(stockData.head())
        file_name = f"{company}StockData.xlsx"
        stockData.to_excel(file_name, engine='openpyxl')
        print(f"Data successfully saved to {file_name}")
        break
    else:
        print("Invalid input, please try again")
