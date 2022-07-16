import numpy as np
import pandas as pd
import requests
import xlsxwriter
import math
from secret_keys import IEX_CLOUD_API_TOKEN

stocks = pd.read_csv('sp_500_stocks.csv')

def chunks(lst, n):
    for i in range(0, len(lst), n):
        yield lst[i:i + n]


symbol_groups = list(chunks(stocks['Ticker'], 100))
symbol_strings = []

for i in range(0, len(symbol_groups)):
    symbol_strings.append(','.join(symbol_groups[i]))

dataframe_columns = ['Ticker', 'Stock Price',  'Market Capitalisation', 'Number of Shares to Buy']

final_dataframe = pd.DataFrame(columns=dataframe_columns)

for symbol_string in symbol_strings:
    batch_api_call_url = f"https://sandbox.iexapis.com/stable/stock/market/batch?symbols={symbol_string}&types=quote&token={IEX_CLOUD_API_TOKEN}"
    data = requests.get(batch_api_call_url).json()
    for symbol in symbol_string.split(','):
        # append is deprecated (don't shoot me) if I were intended to future proof I would probably use pandas concat here instead
        final_dataframe = final_dataframe.append(
            pd.Series([
                symbol,
                data[symbol]['quote']['latestPrice'],
                data[symbol]['quote']['marketCap'],
                'N/A'
            ],
            index=dataframe_columns
            ),
            ignore_index=True
        )

portfolio_size = input('Enter the value of your portfolio: ')

try: 
    val = float(portfolio_size)
except ValueError:
    print("That's not a number, please try again")
    # Will break if continues to input anything that isn't a number
    portfolio_size = input('Enter the value of your portfolio: ')
    val = float(portfolio_size)


position_size = val/len(final_dataframe.index)

for i in range(0, len(final_dataframe.index)):
    final_dataframe.loc[i, 'Number of Shares to Buy'] = math.floor(position_size/final_dataframe.loc[i, 'Stock Price'])

writer = pd.ExcelWriter('Advised Trades.xlsx', engine='xlsxwriter')
final_dataframe.to_excel(writer, 'Advised Trades', index=False)

# Formatting XLSX

background_colour = '#0a0a23'
font_color = '#ffffff'

string_format = writer.book.add_format(
    {
        'font_color': font_color,
        'bg_color': background_colour,
        'border': 1
    }
)

dollar_format = writer.book.add_format(
    {
        'num_format': '$0.00',
        'font_color': font_color,
        'bg_color': background_colour,
        'border': 1
    }
)

integer_format = writer.book.add_format(
    {
        'num_format': '0',
        'font_color': font_color,
        'bg_color': background_colour,
        'border': 1
    }
)

column_formats = {
    'A': ['Ticker', string_format],
    'B': ['Stock Price', dollar_format],
    'C': ['Market Capitalisaion', dollar_format],
    'D': ['Number of shares to buy', integer_format]
}

for column in column_formats.keys():
    writer.sheets['Advised Trades'].set_column(f'{column}:{column}', 18, column_formats[column][1])
    writer.sheets['Advised Trades'].write(f'{column}1', column_formats[column][0], column_formats[column][1])

# Output XLSX
writer.save()