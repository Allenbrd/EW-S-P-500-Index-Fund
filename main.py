import numpy as np
import pandas as pd
import requests

import xlsxwriter
import math

# Importing list of stocks
stocks = pd.read_csv('ressources\sp_500_stocks.csv')

from mysecrets import IEX_CLOUD_API_TOKEN

# initializing the final dataframe to return
my_columns = ['Ticker', 'Stock Price', 'Market Capitalisation', 'Index Percentage', 'Number of Shares to Buy']
final_df = pd.DataFrame(columns = my_columns)

# Dividing stocks into groups
def chuncks(lst, n):
    for i in range(0, len(lst), n):
        yield lst[i:i + n]

symbol_groups = list(chuncks(stocks['Ticker'], 100))
symbol_strings = []


# Making batch API calls to get stocks data
for i in range(0, len(symbol_groups)):
    symbol_strings.append(','.join(symbol_groups[i]))

# Initializing the fund's global market capitalisation variable
total_market_cap = 0 

# Getting data by batch requests
for i in range(len(symbol_groups)):
    # getting data batch from API
    api_url = f"https://sandbox.iex.cloud/v1/data/core/quote/{symbol_strings[i]}?token={IEX_CLOUD_API_TOKEN}"
    data = requests.get(api_url).json()
    # Appending each stock's data to final_df
    for stock in data:
        if stock['marketCap'] is not None:
            total_market_cap += stock['marketCap']
        final_df.loc[-1] =  [stock['symbol'], stock['latestPrice'], stock['marketCap'], 'N/A', 'N/A']
        final_df.index = final_df.index + 1  # shifting index
        final_df = final_df.sort_index()  # sorting by index
        index = final_df.index

# Getting the user's protfolio value
portfolio_size = input("Enter the value of your portoflio: ")
try:
    val = float(portfolio_size)
except ValueError:
    print("That is not a number, please try again.")
    portfolio_size = input("Enter the value of your portoflio: ")
    val = float(portfolio_size)

# Computing the number of shares to buy per company
# This calculation is weighted by each stock's percentage in the fund's total marketcap
for i in range (0, len(final_df.index)):
    if final_df.loc[i, 'Market Capitalisation'] is not None:
        # Computing the stock's percentage of total fund market cap
        position_percent = final_df.loc[i, 'Market Capitalisation']/total_market_cap
        # Computing stock postion size is USD
        position_size = val*position_percent
        # Appending values to dataframe
        final_df.loc[i, 'Index Percentage'] = position_percent
        final_df.loc[i, 'Number of Shares to Buy'] = math.floor(position_size/final_df.loc[i, 'Stock Price'])

# Saving trades on an excel sheet
writer = pd.ExcelWriter('recommended trades.xlsx', engine = 'xlsxwriter')
final_df.to_excel(writer, "Recommended Trades", index = False)



# Formatting the sheet
bg_color = '#0a0a23'
font_color = "#ffffff"

string_format = writer.book.add_format(
    {
        'font_color': font_color,
        'bg_color': bg_color,
        'border': 1
    }
)

dollar_format = writer.book.add_format(
    {
        'num_format': '$0.00',
        'font_color': font_color,
        'bg_color': bg_color,
        'border': 1
    }
)

integer_format = writer.book.add_format(
    {
        'num_format': 0,
        'font_color': font_color,
        'bg_color': bg_color,
        'border': 1
    }
)

percent_format = writer.book.add_format(
    {
        'num_format': '00.00%',
        'font_color': font_color,
        'bg_color': bg_color,
        'border': 1
    }
)

column_formats = { 
                    'A': ['Ticker', string_format],
                    'B': ['Price', dollar_format],
                    'C': ['Market Capitalization', dollar_format],
                    'D': ['Index Percentage', percent_format],
                    'E': ['Number of Shares to Buy', integer_format]
                    }

for column in column_formats.keys():
    writer.sheets['Recommended Trades'].set_column(f'{column}:{column}', 20, column_formats[column][1])
    writer.sheets['Recommended Trades'].write(f'{column}1', column_formats[column][0], string_format)

writer._save()