import pandas as pd
def run(rawDataFilePath, terminatedPatientFilePath, date_start, date_end):
    
    # Creates the two dataframes for each file
    Raw_Data_DF = pd.read_excel(rawDataFilePath)
    Terminated_Patient_DF = pd.read_excel(terminatedPatientFilePath)
    
    #TODO Develop date functionality
    year_needed = date_start.year()
    
    # Calls the filterInactive function that labels each patient as inactive(Terminated) or active
    A_or_I_DataFrame = Add_ActiveOrInactive(Raw_Data_DF, Terminated_Patient_DF)
    
    # Calls the filter_DF function and passes in the raw dataframe with the Active/Inactive column filled out for each patient
    filtered_DataFrames = filter_RawData_DF(A_or_I_DataFrame[0], year_needed, A_or_I_DataFrame[1])
    
    # Calls the function that counts number of times each patient has a therapy session
    # Passes in the two filtered dataframes based on the user's input
    patientAttendanceDicts = countAttendance(filtered_DataFrames[0], filtered_DataFrames[1])
    
    # Calls function that passes in the two dictionaries with the patients names and the number of times they had a therapy session
    return RangeTotals(patientAttendanceDicts[0], patientAttendanceDicts[1])

# Function that combines the first and last names of each patient, and labels each patient as inactive(Terminated) or active
def Add_ActiveOrInactive(RawData_DF, TermiantedPatient_DF):

    # Combines the first and last names for each patient
    # First Dataframe that contains all patients
    RawData_DF['Full Name'] = RawData_DF['First Name'].str.cat(RawData_DF['Last Name'],sep=" ")
    # Second dataframe that only contains the terminated patients
    TermiantedPatient_DF['Full Name'] = TermiantedPatient_DF['First Name'].str.cat(TermiantedPatient_DF['Last Name'],sep=" ")

    #Creates a list of the active and terminated patient names
    Active_Names = RawData_DF['Full Name'].unique()
    Inactive_names = TermiantedPatient_DF['Full Name'].unique()

    # Creates a column in the raw data to fill out if the patient is terminated or not
    RawData_DF['Active/Inactive'] = " "

    # Checks if the patients name in the raw dataframe is in the list of terminated patients
    for name in Active_Names:
        if name in Inactive_names:
            # If the patient's name is in the list of terminated patients, labels them as inactive in the Active/Inactive column
            RawData_DF.loc[(RawData_DF['Full Name'] == name), 'Active/Inactive'] = 'Inactive'
        else:
            # Fills out the Active/Inactive column as active if the patient's name is not in the list of terminated patients
            RawData_DF.loc[(RawData_DF['Full Name'] == name), 'Active/Inactive'] = 'Active'
    
    return [RawData_DF, Inactive_names]

# Function that filters out all the group therapy patients and non appointment rows in the patient dataframe
def filter_RawData_DF(RawData_DF, year_needed, TerminatedNames):
    
    TerminatedNames_List = TerminatedNames

    # Changes the date column in the dataframe into an easier format to parse
    RawData_DF['Date'] = pd.to_datetime(RawData_DF['Date'])

    # If statement checking that the user wants all time statsitics
    if year_needed == 0:

        # Creates a filtered dataframe that only includes appointments that are therapy sessions or therapy intake
        Therapy_Sessions_DF = RawData_DF.loc[(RawData_DF['Type'] == 'Appointment') & (RawData_DF['Appointment Type'] == 'Therapy Session') | (RawData_DF['Appointment Type'] == 'Therapy Intake')]

    else:

        # Creates a filtered dataframe that only includes appointments that are therapy sessions or therapy intake from the user's inputted year
        Therapy_Sessions_DF = RawData_DF.loc[(RawData_DF['Type'] == 'Appointment') & (RawData_DF['Date'].dt.year == year_needed) & (RawData_DF['Appointment Type'] == 'Therapy Session') | (RawData_DF['Appointment Type'] == 'Therapy Intake')]

    return [Therapy_Sessions_DF, TerminatedNames_List]

# Function that counts the number of times  each patient has a therapy session
def countAttendance(TherapySessions_DF, TerminatedNamesList):

    # Creates a dictionary with the key as the name of the patient and the value as the number of times they had a therapy session for all patients
    countPatientTS_Dict = dict(TherapySessions_DF['Full Name'].value_counts())

    countTerminatedTS_Dict = {}
    for key, value in countPatientTS_Dict.items():
        if key in TerminatedNamesList:
            countTerminatedTS_Dict[key] = value

    return [countPatientTS_Dict, countTerminatedTS_Dict]

# Function that calculates the number of patients that had a total number of therapy sessions attended within a certain range
def RangeTotals(countPatientTS, countTerminatedTS):
#--------------------------------------
    # Creates the variables for each range
    ZeroToThree = 0
    FourToSeven = 0
    EightToTen = 0
    ElevenToFourteen = 0
    FifteenPlus = 0
    Total = 0
#----------------------------------------
    # Iterates through each value in the dictionary and checks in what range the number of therapy sessions attended is within
    # Adds one to the range variable if within the range
    for value in countTerminatedTS.values():
        Total += 1
        if value >= 0 and value <= 3:
            ZeroToThree += 1
        elif value >= 4 and value <= 7:
            FourToSeven += 1
        elif value >= 8 and value <= 10:
            EightToTen += 1
        elif value >= 11 and value <=14:
            ElevenToFourteen += 1
        else:
            FifteenPlus += 1
#-------------------------------------------------------------
    # All the if statements check if the range variable is not zero to make sure there is no division by zero
    if ZeroToThree != 0:
        # Divides the number of times a patient had total therapy sessions attended within the 0-3 range by the total amount of therapy sessions attended
        ZerotoThreePercent = round((ZeroToThree/Total), 2)
    else:
        # Sets the number equal to the range variable since the number of times is zero
        ZerotoThreePercent = ZeroToThree

    if FourToSeven != 0:
        # Divides the number of times a patient had total therapy sessions attended within the 4-7 range by the total amount of therapy sessions attended
        FourtoSevenPercent = round((FourToSeven/Total), 2)
    else:
        # Sets the number equal to the range variable since the number of times is zero
        FourtoSevenPercent = FourToSeven

    if EightToTen != 0:
        # Divides the number of times a patient had total therapy sessions attended within the 8-10 range by the total amount of therapy sessions attended
        EighttoTenPercent = round((EightToTen/Total), 2)
    else:
        # Sets the number equal to the range variable since the number of times is zero
        EighttoTenPercent = EightToTen

    if ElevenToFourteen != 0:
        # Divides the number of times a patient had total therapy sessions attended within the 11-14 range by the total amount of therapy sessions attended
        ElevenToFourteenPercent = round((ElevenToFourteen/Total), 2)
    else:
        # Sets the number equal to the range variable since the number of times is zero
        ElevenToFourteenPercent = ElevenToFourteen

    if FifteenPlus != 0:
        # Divides the number of times a patient had total therapy sessions where they attended 15+ by the total amount of therapy sessions attended
        FifteenPlusPercent = round((FifteenPlus/Total), 2)
    else:
        # Sets the number equal to the range variable since the number of times is zero
        FifteenPlusPercent = FifteenPlus
#---------------------------------------------------------------
    # Creates a dictionary of the number ranges as the keys and the number of times a patient's total therapy sessions attended was within that range
    Range_Totals = {'0-3':ZeroToThree, '4-7':FourToSeven, '8-10':EightToTen, '11-14':ElevenToFourteen, '15+':FifteenPlus}
    # Creates a dictionary of the number ranges as the keys and the percentage of the number of times a patient's total therapy sessions attended was within that range by the total for all ranges
    Range_Total_Percents = {'0-3':ZerotoThreePercent, '4-7':FourtoSevenPercent, '8-10':EighttoTenPercent, '11-14':ElevenToFourteenPercent, '15+':FifteenPlusPercent}

    # Calls function that creates the excel sheet and passes in the dictionary of range values and percentages as well as the dictionary with the patient name and the number of therapy sessions attended
    return create_Excel(countPatientTS, Range_Totals, Range_Total_Percents, countTerminatedTS)

# Function that creates the output excel file that has the retention report
def create_Excel(countPatientTS_Dict, Range_Totals_Dict, Range_Total_Percents_Dict, countTerminatedTS_Dict):

    # Creates a list of terminated patient names
    TerminatedPatientNames_List = []
    for key in countTerminatedTS_Dict.keys():
        TerminatedPatientNames_List.append(key)

    # Prepares the passed in dictionary to allow for easier creation of the DataFrame
    # Creates a DataFrame from the dictionary that has all the patient names and the total number of sessions they attended
    prepare_PatientTS_Dict = {i: x for i, x in enumerate(countPatientTS_Dict.items())}
    PatientTS_DF = pd.DataFrame.from_dict(prepare_PatientTS_Dict, orient='index', columns=["Patient Name", "Sessions Attended"])

    # Prepares the passed in dictionary to allow for easier creation of the DataFrame
    # Creates a dataframe from the dictionary that has all the range bins and the number of times a patient's total therapy sessions attended was within that range
    prepare_RangeTotals_Dict = {i: x for i, x in enumerate(Range_Totals_Dict.items())}
    RangeTotal_DF = pd.DataFrame.from_dict(prepare_RangeTotals_Dict, orient='index', columns=["Ranges", "Sessions Attended"])


    # Prepares the passed in dictionary to allow for easier creation of the DataFrame
    # Creates a dataframe from the dictionary that has all the range bins and the percentage of the number of times a patient's total therapy sessions attended was within that range by the total for all ranges
    prepare_RangeTotalPercents_Dict = {i: x for i, x in enumerate(Range_Total_Percents_Dict.items())}
    RangeTotalPercents_DF = pd.DataFrame.from_dict(prepare_RangeTotalPercents_Dict, orient='index', columns=["Ranges", "Percentages"])

    # Creates a string holding the file name 'Retention Report'
    file_name = 'Retention_Report.xlsx'

    # Statement that takes each dataframe and puts it into an excel file with the sheet name being 'Retention Report'
    writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
    PatientTS_DF.to_excel(writer, sheet_name = 'Retention Report', startcol=0, index=False)
    RangeTotal_DF.to_excel(writer, sheet_name='Retention Report', startcol=4, index=False)
    RangeTotalPercents_DF.to_excel(writer, sheet_name='Retention Report', startcol= 7, index=False)

    #Opens the workbook that was just created called 'Retention Report'
    workbook = writer.book
    worksheet = writer.sheets['Retention Report']

    # Creates the color and text format that each cell will be changed too based on what range a patient's total number of sessions attended is in
    Red_Range_Bin = workbook.add_format({'bg_color': '#FFCCCB', 'bold': True})
    Orange_Range_Bin = workbook.add_format({'bg_color': '#FFB214', 'bold': True})
    Yellow_Range_Bin = workbook.add_format({'bg_color': '#FFD070', 'bold': True})
    Blue_Range_Bin = workbook.add_format({'bg_color': '#8CCBFF', 'bold': True})
    Green_Range_Bin = workbook.add_format({'bg_color': '#94ffb8', 'bold': True})

    # Creates the color and text format that each cell will be changed too based on if the patient is active or inactive
    Red_InactivePatient = workbook.add_format({'bg_color': '#ff0000', 'bold': True})
    Green_ActivePatient = workbook.add_format({'bg_color': '#00FF00', 'bold': True})

    # Creates the format that every percentage will be changed too
    formatPercentages = workbook.add_format({'num_format': '0%'})

    # Determines where to start to look for the total sessions attended for each patient to be able to color them
    startRow = 1
    endRow = startRow + PatientTS_DF.shape[0] - 1
    startColumn = 1
    endColumn = 1

    # Adding color to each patient's total attended sessions based on what range they are in using conditional formatting
    worksheet.conditional_format(startRow, startColumn, endRow, endColumn, {'type':'cell', 'criteria':'between', 'minimum': 0,'maximum': 3, 'format':Red_Range_Bin})
    worksheet.conditional_format(startRow, startColumn, endRow, endColumn, {'type':'cell', 'criteria':'between', 'minimum': 4,'maximum': 7, 'format':Orange_Range_Bin})
    worksheet.conditional_format(startRow, startColumn, endRow, endColumn, {'type':'cell', 'criteria':'between', 'minimum': 8,'maximum': 10, 'format':Yellow_Range_Bin})
    worksheet.conditional_format(startRow, startColumn, endRow, endColumn, {'type':'cell', 'criteria':'between', 'minimum': 11,'maximum': 14, 'format':Blue_Range_Bin})
    worksheet.conditional_format(startRow, startColumn, endRow, endColumn, {'type':'cell', 'criteria':'>=', 'value': 15, 'format':Green_Range_Bin})

    # Adding the same corresponding color to the labels of each range total
    worksheet.conditional_format('E2:E2', {'type': 'text','criteria':'containing', 'value':'0-3', 'format':   Red_Range_Bin})
    worksheet.conditional_format('E3:E3', {'type': 'text','criteria':'containing', 'value':'4-7', 'format':   Orange_Range_Bin})
    worksheet.conditional_format('E4:E4', {'type': 'text','criteria':'containing', 'value':'8-10', 'format':   Yellow_Range_Bin})
    worksheet.conditional_format('E5:E5', {'type': 'text','criteria':'containing', 'value':'11-14', 'format':   Blue_Range_Bin})
    worksheet.conditional_format('E6:E6', {'type': 'text','criteria':'containing', 'value':'15+', 'format':   Green_Range_Bin})

    # Adding the same corresponding color to the labels of each range total percentage
    worksheet.conditional_format('H2:H2', {'type':'text','criteria':'containing', 'value':'0-3', 'format':   Red_Range_Bin})
    worksheet.conditional_format('H3:H3', {'type':'text','criteria':'containing', 'value':'4-7', 'format':   Orange_Range_Bin})
    worksheet.conditional_format('H4:H4', {'type':'text','criteria':'containing', 'value':'8-10', 'format':   Yellow_Range_Bin})
    worksheet.conditional_format('H5:H5', {'type':'text','criteria':'containing', 'value':'11-14', 'format':   Blue_Range_Bin})
    worksheet.conditional_format('H6:H6', {'type':'text','criteria':'containing', 'value':'15+', 'format':   Green_Range_Bin})

    # Formats the range total percentages by rounding them to zero decimal points (whole number)
    worksheet.set_column('I2:I6', 10, formatPercentages)

    # Loops through each patient name and checks if they are in the list of terminated patient names
    for index in range(len(PatientTS_DF)):
        # Figures out what cell to add the color too
        RowNumber = str(index+2)
        ExactCell = 'A'+RowNumber
        
        # If they are termianted then it colors their name red
        if (PatientTS_DF['Patient Name'].iloc[index] in TerminatedPatientNames_List):
            worksheet.conditional_format(ExactCell,{'type': 'no_blanks', 'format': Red_InactivePatient})
    
        # Otherwise it colors their name green to signify they are still active
        else:
            worksheet.conditional_format(ExactCell,{'type': 'no_blanks', 'format': Green_ActivePatient})

    # Saves the workbook with the newly made changes to the format
    writer.save()

    # Statement that prints out if it made it to the end of the program
    print("Success")

    #Returns the file name
    return file_name
