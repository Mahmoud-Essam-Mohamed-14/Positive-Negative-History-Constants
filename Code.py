import numpy as np
import pandas as pd
import matplotlib as mlp
import matplotlib.pyplot as plt
from openpyxl import load_workbook


def main():
    # Loading Workbook sheets into pandas DataFrames
    df_1 = pd.read_excel('basic_data.xlsx', sheet_name='spreadsheet 1')
    df_2 = pd.read_excel('basic_data.xlsx', sheet_name='spreadsheet 2')
    df_3 = pd.read_excel('basic_data.xlsx', sheet_name='output')
    df_4 = pd.read_excel('basic_data.xlsx', sheet_name='function_history')
    # Checking current mode of functions (negative sequence '-' or positive sequence '+')
    df_3['result'] = np.where(df_1[263] > df_2[262], '+', '-')
    # Checking functions history sequence (negative sequence '0' or positive sequence '1')
    for i in range(1, 262):
        df_4[i] = np.where(df_2[i + 1] > df_2[i], 1, 0)
    df_4[262] = np.where(df_1[263] > df_2[262], 1, 0)

    # Calculate Positive history constant and Negative history constant
    df_4.set_index('Function', inplace=True)
    df_4['positive_sequence_total'] = np.nan
    df_4['positive_sequence_interval'] = np.nan
    df_4['negative_sequence_total'] = np.nan
    df_4['negative_sequence_interval'] = np.nan
    function_list = list(df_4.index.values)

    for function in function_list:
        positive_sequence_total = 0
        negative_sequence_total = 0
        positive_sequence_intervals = 0
        negative_sequence_intervals = 0
        for i in range(1, 263):
            if df_4.loc[function, i] == 1:
                positive_sequence_total += 1
                if (i == 1) or ((i > 1) and (df_4.loc[function, i] != df_4.loc[function, i - 1])):
                    positive_sequence_intervals += 1
            else:
                negative_sequence_total += 1
                if (i == 1) or ((i > 1) and (df_4.loc[function, i] != df_4.loc[function, i - 1])):
                    negative_sequence_intervals += 1

        df_4.loc[function, 'positive_sequence_total'] = positive_sequence_total
        df_4.loc[function, 'positive_sequence_interval'] = positive_sequence_intervals
        df_4.loc[function, 'negative_sequence_total'] = negative_sequence_total
        df_4.loc[function, 'negative_sequence_interval'] = negative_sequence_intervals

    df_4['positive history constant'] = df_4['positive_sequence_total'] / df_4['positive_sequence_interval']
    df_4['negative history constant'] = df_4['negative_sequence_total'] / df_4['negative_sequence_interval']
    df_4.reset_index(inplace=True)

    # Writing Positive history constant and Negative history constant in output Sheet
    df_3['negative history constant'] = df_4['positive history constant']
    df_3['positive history constant'] = df_4['negative history constant']

    # Create new pandas DataFrame including current negative functions only and setting function as index
    df_5 = df_3[df_3['result'] == '-']
    df_5.set_index('Function', inplace=True)

    # Create new pandas DataFrame and importing functions history, then setting function as index
    df_6 = df_4
    df_6.set_index('Function', inplace=True)

    # Create new pandas DataFrame that merge DataFrame No 5 and DataFrame No 6
    df_7 = pd.merge(df_5, df_6, how='inner', right_index=True, left_index=True)

    # Calculating current decline steps
    negative_function_list = list(df_7.index.values)
    df_3.set_index('Function', inplace=True)
    for function in negative_function_list:
        decline_count = 0
        for i in range(1, 263):
            if df_7.loc[function, 263 - i] == 1:
                break
            else:
                decline_count += 1
        df_7.loc[function, 'decline count'] = decline_count
        df_3.loc[function, 'decline count'] = decline_count
    df_3.reset_index(inplace=True)

    # Creating new DataFrame that include functions that have (current negative mode & decline steps > negative const)
    df_8 = df_3[df_3['result'] == '-']
    df_8 = df_8[df_8['decline count'] >= df_8['negative history constant']]
    df_8.set_index('Function', inplace=True)

    # Creating new DataFrame that include functions history from 1 to 263
    df_9 = pd.merge(df_1, df_2, how='inner', left_on='Function', right_on='Function')
    df_9.set_index('Function', inplace=True)

    # Merge DataFrame No 8 and DataFrame No 9, So now we can plot the Criteria Functions
    df_10 = pd.merge(df_9, df_8, how='inner', left_index=True, right_index=True)

    # Ploting the Criteria Functions
    x = list(np.arange(1, 264))
    criteria_function = list(df_10.index.values)
    for function in criteria_function:
        y = []
        title = function.upper()
        for i in range(1, 264):
            y.append(df_10.loc[function, i])
        plt.plot(x, y)
        plt.title(title)
        plt.grid()
        plt.savefig(title)

        plt.clf()
        plt.cla()
        plt.close()

    # Editing DataFrame No 4 with ('+' and '-') signs instead of (1 and 0)
    for i in range(1, 262):
        df_4[i] = np.where(df_2[i + 1] > df_2[i], '+', '-')
    df_4[262] = np.where(df_1[263] > df_2[262], '+', '-')
    df_4.reset_index(inplace=True)

    # Writing our DataFrames to an Excel file
    with pd.ExcelWriter('Basic Data.xlsx') as writer:
        df_1.to_excel(writer, sheet_name='spreadsheet 1', index=False, startcol=1)
    book = load_workbook('Basic Data.xlsx')
    writer = pd.ExcelWriter('Basic Data.xlsx')
    writer.book = book
    writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
    df_2.to_excel(writer, sheet_name='spreadsheet 2', index=False, startcol=1)
    df_3.to_excel(writer, sheet_name='output', index=False, startcol=1)
    df_4.to_excel(writer, sheet_name='history', index=False, startcol=1)
    df_8.to_excel(writer, sheet_name='graph', index=False, startcol=1)
    writer.save()


if __name__ == '__main__':
    main()
