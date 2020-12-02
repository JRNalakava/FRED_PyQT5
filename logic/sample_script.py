import pandas as pd, traceback


# this is a place holder script for testing the GUI
# its purpose is to add the first columns of provided excel files
# this proves that the user can upload two files through the GUI
# and that the system can process those files


# filepath_1: path of first file
# filepath_2: path of second file
# output: file path of processed file
# function takes in file paths and creates data frames with them.
# function then processes the data_frame and returns the filepath of processed
# file from process()
def run(filepath_1, filepath_2):
    # Creates the two dataframes for each file
    file_1_df = df = pd.read_excel(filepath_1, index_col=None, na_values=['NA'], usecols="A")
    file_2_df = pd.read_excel(filepath_2, index_col=None, na_values=['NA'], usecols="A")

    # process data frames and output result into excel file
    temp_filepath = process(file_1_df, file_2_df)
    return temp_filepath


# data_frame_1: Pandas data frame for file 1
# data_frame_2: Pandas data frame for file 2
# output: file path of processed file
# function takes in two data frames and adds up the cells in each column
# creates an excel file and returns the filepath
def process(data_frame_1, data_frame_2):
    # Changes data frame to numpy array for easier manipulation
    data_arr_1 = data_frame_1.to_numpy()
    data_arr_2 = data_frame_2.to_numpy()
    # Adds up each a[i] + b[i], i = a.length
    process_arr = data_arr_1 + data_arr_2
    # Try catch clause
    try:
        with pd.ExcelWriter('test_file.xlsx', mode='A') as writer:
            pd.DataFrame(process_arr).to_excel(writer, sheet_name = 'Summation', startcol=0)
        filepath = 'test_file.xlsx'

        return filepath
    except:
        traceback.print_exc()