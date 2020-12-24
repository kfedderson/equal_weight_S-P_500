import numpy as np
import pandas as pd
import requests
import xlsxwriter
import math
import os

SB_TOKEN = os.getenv('SB_TOKEN')
stocks = pd.read_csv('sp_500_stocks.csv')
my_columns = ['Ticker', 'Stock Price', 'Market Capitalization', 'Shares to buy']
PORTFOLIO_SIZE = 1000000

# # Add stocks to pandas dataframe, single API call
# my_columns = ['Ticker', 'Stock Price', 'Market Capitalization', 'Shares to buy']
# final_dataframe = pd.DataFrame(columns = my_columns)
# for stock in stocks['Ticker'][:5]:
#   api_url = f'https://sandbox.iexapis.com/stable/stock/{stock}/quote/?token={SB_TOKEN}'
#   data = requests.get(api_url).json()
#   final_dataframe = final_dataframe.append(
#     pd.Series(
#       [
#         stock,
#         data['latestPrice'],
#         data['marketCap'],
#         'N/A'
#       ],
#       index = my_columns
#     ),
#     ignore_index = True
#   )

def chunks(lst, n):
  for i in range(0,len(lst), n):
    yield lst[i:i + n]

# Add stocks to pandas dataframe, batch API call
symbol_groups = list(chunks(stocks['Ticker'], 100))
symbol_strings = []

for i in range(0, len(symbol_groups)):
  symbol_strings.append(','.join(symbol_groups[i]))

final_dataframe = pd.DataFrame(columns = my_columns)

for symbol_string in symbol_strings:
  api_url = f'https://sandbox.iexapis.com/stable/stock/market/batch?symbols={symbol_string} &types=quote&token={SB_TOKEN}'
  data = requests.get(api_url).json()
  for symbol in symbol_string.split(','):
    final_dataframe = final_dataframe.append(
      pd.Series(
        [
          symbol,
          data[symbol]['quote']['latestPrice'],
          data[symbol]['quote']['marketCap'],
          'N/A'
        ],
        index = my_columns
      ),
      ignore_index = True
    )

position_size = PORTFOLIO_SIZE/len(final_dataframe.index)

# Calculate number of shares to buy for each stock and write to dataframe
for i in range(0, len(final_dataframe.index)):
  final_dataframe.loc[i, 'Shares to buy'] = math.floor(position_size/final_dataframe.loc[i, 'Stock Price'])

# Write results out to excel file with formatting
writer = pd.ExcelWriter('recommended_trades.xlsx', engine = 'xlsxwriter')
final_dataframe.to_excel(writer, 'Recommended Trades', index = False)

bg_color = '#0a0a23'
ft_color = '#ffffff'

string_format = writer.book.add_format(
  {
    'font_color': ft_color,
    'bg_color': bg_color,
    'border': 1
  }
)

dollar_format = writer.book.add_format(
  {
    'num_format': '$0.00',
    'font_color': ft_color,
    'bg_color': bg_color,
    'border': 1
  }
)

int_format = writer.book.add_format(
  {
    'num_format' : '0',
    'font_color' : ft_color,
    'bg_color' : bg_color,
    'border' : 1
  }
)

column_formats = {
  'A': ['Ticker', string_format],
  'B': ['Stock Price', dollar_format],
  'C': ['Market Capitalization', dollar_format],
  'D': ['Shares to buy', int_format]
}

for column in column_formats.keys():
  writer.sheets['Recommended Trades'].set_column(f'{column}:{column}', 18, column_formats[column][1])
  writer.sheets['Recommended Trades'].write(f'{column}1', column_formats[column][0], column_formats[column][1])

writer.save()