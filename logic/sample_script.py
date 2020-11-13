import os

import pandas as pd, traceback


# this is a place holder script for testing the GUI
# its purpose is to add the first columns of provided excel files
# this proves that the user can upload two files through the GUI
# and that the system can process those files

def run(filepath_1, filepath_2):
    # Creates the two dataframes for each file
    file_1_df = df = pd.read_excel(filepath_1, index_col=None, na_values=['NA'], usecols="A")
    file_2_df = pd.read_excel(filepath_2, index_col=None, na_values=['NA'], usecols="A")

    # process data frames and output stuff into excel file
    temp_filepath = process(file_1_df, file_2_df)
    return temp_filepath


def process(data_frame_1, data_frame_2):
    data_arr_1 = data_frame_1.to_numpy()
    data_arr_2 = data_frame_2.to_numpy()
    process_arr = data_arr_1 + data_arr_2
    try:
        with pd.ExcelWriter('test_file.xlsx', mode='A') as writer:
            pd.DataFrame(process_arr).to_excel(writer, sheet_name = 'Summation', startcol=0)
    except:
        traceback.print_exc()
    filepath = 'test_file.xlsx'

    return filepath
