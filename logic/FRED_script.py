import pandas as pd
def run(rPath, iPath, date_start, date_end):
    # Creates the two dataframes for each file
    DF = pd.read_excel(rPath)
    DF2 = pd.read_excel(iPath)
    #TODO Develop date functionality
    year_needed = date_start.year()
    # Calls the filterInactive function that labels each patient as inactive(Terminated) or active
    DFS = filterInactive(DF, DF2)
    # Calls the filter_DF function and passes in the raw dataframe with the Active/Inactive column filled out for each patient
    filtered_DFs = filter_DF(DFS[0], year_needed, DFS[1])
    # Calls the function that counts number of times each patient has a therapy session
    # Passes in the two filtered dataframes based on the user's input
    attendance_DFs = countAttendance(filtered_DFs[0], filtered_DFs[1])
    # Calls function that passes in the two dictionaries with the patients names and the number of times they had a therapy session
    return RangesTotals(attendance_DFs[0], attendance_DFs[1])

# Function that combines the first and last names of each patient, and labels each patient as inactive(Terminated) or active
def filterInactive(DF, DF2):

    # Combines the first and last names for each patient
    # First Dataframe that contains all patients
    DF['Full Name'] = DF['First Name'].str.cat(DF['Last Name'],sep=" ")
    # Second dataframe that only contains the terminated patients
    DF2['Full Name'] = DF2['First Name'].str.cat(DF2['Last Name'],sep=" ")

    #Creates a list of the terminated patient names
    Active_Names = DF['Full Name'].unique()
    Inactive_names = DF2['Full Name'].unique()

    # Creates a column in the raw data to fill out if the patient is terminated or not
    DF['Active/Inactive'] = " "

    # For loop that iterates through all the rows in the raw dataframe
    #for i, row in DF.iterrows():
    '''
    for i, row in DF.iterrows():
            # Checks if the patients name in the raw dataframe is in the list of terminated patients
        if row['Full Name'] in Inactive_names:
            # If the patient's name is in the list of terminated patients, labels them as inactive in the Active/Inactive column
            DF.at[i, 'Active/Inactive'] = 'Inactive'
        else:
            # Fills out the Active/Inactive column as active if the patient's name is not in the list of terminated patients
            DF.at[i, 'Active/Inactive'] = 'Active'
    '''
    for name in Active_Names:
        if name in Inactive_names:
            #print(name)
        # Checks if the patients name in the raw dataframe is in the list of terminated patients
        #if row['Full Name'] in Inactive_names:
            # If the patient's name is in the list of terminated patients, labels them as inactive in the Active/Inactive column
            DF.loc[(DF['Full Name'] == name), 'Active/Inactive'] = 'Inactive'
        else:
            # Fills out the Active/Inactive column as active if the patient's name is not in the list of terminated patients
            DF.loc[(DF['Full Name'] == name), 'Active/Inactive'] = 'Active'

    return [DF, Inactive_names]

# Function that filters out all the group therapy patients and non appointment rows in the patient dataframe
def filter_DF(DF, year_needed, INames):
    I = INames
    ###DF['Full Name'] = DF['First Name'].str.cat(DF['Last Name'],sep=" ")

    # Changes the date column in the dataframe into an easier format to parse
    DF['Date'] = pd.to_datetime(DF['Date'])

    # If statement checking that the user wants all time statsitics
    if year_needed == 0:

        # Creates a filtered dataframe that only includes appointments that are therapy sessions or therapy intake
        Therapy_Sessions = DF.loc[(DF['Type'] == 'Appointment') & (DF['Appointment Type'] == 'Therapy Session') | (DF['Appointment Type'] == 'Therapy Intake')]
        # Creates a filtered dataframe that only includes appointments that are therapy sessions from patients that are terminated
        #Inactive_Clients = DF.loc[(DF['Type'] == 'Appointment') & (DF['Appointment Type'] == 'Therapy Session') | (DF['Appointment Type'] == 'Therapy Intake') & (DF['Active/Inactive'] == 'Inactive')]
    else:

        # Creates a filtered dataframe that only includes appointments that are therapy sessions or therapy intake from the user's inputted year
        Therapy_Sessions = DF.loc[(DF['Type'] == 'Appointment') & (DF['Date'].dt.year == year_needed) & (DF['Appointment Type'] == 'Therapy Session') | (DF['Appointment Type'] == 'Therapy Intake')]
        # Creates a filtered dataframe that only includes appointments that are therapy sessions or therapy intake from patients that are terminated in the user's inputted year
        #Inactive_Clients = DF.loc[(DF['Type'] == 'Appointment') & (DF['Appointment Type'] == 'Therapy Session') | (DF['Appointment Type'] == 'Therapy Intake') & (DF['Date'].dt.year == year_needed) & (DF['Active/Inactive'] == 'Inactive')]

    return [Therapy_Sessions, I]

# Function that counts the number of times  each patient has a therapy session
def countAttendance(Therapy, INM):

    # Creates a dictionary with the key as the name of the patient and the value as the number of times they had a therapy session for all patients
    Dict_Values = dict(Therapy['Full Name'].value_counts())

    Dict_Terminated = {}
    for key, value in Dict_Values.items():
        if key in INM:
            Dict_Terminated[key] = value

    return [Dict_Values, Dict_Terminated]

# Function that calculates the number of patients that had a total number of therapy sessions attended within a certain range
def RangesTotals(Dict1, Dict2):
#--------------------------------------
    # Creates the variables for each range
    ZeroToThree2 = 0
    FourToSeven2 = 0
    EightToTen2 = 0
    ElevenToFourteen2 = 0
    FifteenPlus2 = 0
    Total2 = 0
#----------------------------------------
    # Iterates through each value in the dictionary and checks in what range the number of therapy sessions attended is within
    # Adds one to the range variable if within the range
    for value in Dict2.values():
        Total2 += 1
        if value >= 0 and value <= 3:
            ZeroToThree2 += 1
        elif value >= 4 and value <= 7:
            FourToSeven2 += 1
        elif value >= 8 and value <= 10:
            EightToTen2 += 1
        elif value >= 11 and value <=14:
            ElevenToFourteen2 += 1
        else:
            FifteenPlus2 += 1
#-------------------------------------------------------------
    # All the if statements check if the range variable is not zero to make sure there is no division by zero
    if ZeroToThree2 != 0:
        # Divides the number of times a patient had total therapy sessions attended within the 0-3 range by the total amount of therapy sessions attended
        # Multiplies the number by 100 and rounds to two decimal points in order to be a percentage
        ZtoThreeP2 = round((ZeroToThree2/Total2), 2)
    else:
        # Sets the number equal to the range variable since the number of times is zero
        ZtoThreeP2 = ZeroToThree2

    if FourToSeven2 != 0:
        # Divides the number of times a patient had total therapy sessions attended within the 4-7 range by the total amount of therapy sessions attended
        # Multiplies the number by 100 and rounds to two decimal points in order to be a percentage
        FtoSevenP2 = round((FourToSeven2/Total2), 2)
    else:
        # Sets the number equal to the range variable since the number of times is zero
        FtoSevenP2 = FourToSeven2

    if EightToTen2 != 0:
        # Divides the number of times a patient had total therapy sessions attended within the 8-10 range by the total amount of therapy sessions attended
        # Multiplies the number by 100 and rounds to two decimal points in order to be a percentage
        EtoTenP2 = round((EightToTen2/Total2), 2)
    else:
        # Sets the number equal to the range variable since the number of times is zero
        EtoTenP2 = EightToTen2

    if ElevenToFourteen2 != 0:
        # Divides the number of times a patient had total therapy sessions attended within the 11-14 range by the total amount of therapy sessions attended
        # Multiplies the number by 100 and rounds to two decimal points in order to be a percentage
        ElToFourteenP2 = round((ElevenToFourteen2/Total2), 2)
    else:
        # Sets the number equal to the range variable since the number of times is zero
        ElToFourteenP2 = ElevenToFourteen2

    if FifteenPlus2 != 0:
        # Divides the number of times a patient had total therapy sessions where they attended 15+ by the total amount of therapy sessions attended
        # Multiplies the number by 100 and rounds to two decimal points in order to be a percentage
        FifteenPlusP2 = round((FifteenPlus2/Total2), 2)
    else:
        # Sets the number equal to the range variable since the number of times is zero
        FifteenPlusP2 = FifteenPlus2
#---------------------------------------------------------------
    # Creates a dictionary of the number ranges as the keys and the number of times a patient's total therapy sessions attended was within that range
    Range_Vals2 = {'0-3':ZeroToThree2, '4-7':FourToSeven2, '8-10':EightToTen2, '11-14':ElevenToFourteen2, '15+':FifteenPlus2}
    # Creates a dictionary of the number ranges as the keys and the percentage of the number of times a patient's total therapy sessions attended was within that range by the total for all ranges
    Range_Val_Percents2 = {'0-3':ZtoThreeP2, '4-7':FtoSevenP2, '8-10':EtoTenP2, '11-14':ElToFourteenP2, '15+':FifteenPlusP2}

    # Calls function that creates the excel sheet and passes in the dictionary of range values and percentages as well as the dictionary with the patient name and the number of therapy sessions attended
    return create_Excel(Dict1, Range_Vals2, Range_Val_Percents2, Dict2)

# Function that creates the output excel file that has the retention report
def create_Excel(Dict1, Range_Vals2, Range_Val_Percents2, Dict2):

    # Creates a dataframe from the dictionary that has all the patient names and the number of therapy sessions attended
    Terminated_Patients = []
    for key in Dict2.keys():
        #print(key)
        Terminated_Patients.append(key)

    prepared_dict = {i: x for i, x in enumerate(Dict1.items())}
    DictNames = pd.DataFrame.from_dict(prepared_dict, orient='index', columns=["Patient Name", "Sessions Attended"])

    # Creates a dataframe from the dictionary that has all the range bins and the number of times a patient's total therapy sessions attended was within that range
    prepared_dict2 = {i: x for i, x in enumerate(Range_Vals2.items())}
    RangeVal2 = pd.DataFrame.from_dict(prepared_dict2, orient='index', columns=["Ranges", "Sessions Attended"])


    # Creates a dataframe from the dictionary that has all the range bins and the percentage of the number of times a patient's total therapy sessions attended was within that range by the total for all ranges
    prepared_dict3 = {i: x for i, x in enumerate(Range_Val_Percents2.items())}
    RangeValP2 = pd.DataFrame.from_dict(prepared_dict3, orient='index', columns=["Ranges", "Percentages"])

    # Creates a string holding the file name 'Retention Report'
    file_name = 'Retention_Report.xlsx'

    # Statement that takes each dataframe and puts it into an excel file with the sheet name being 'Retention Report'
    writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
    #with pd.ExcelWriter('Retention_Report.xlsx', mode='A') as writer:
    DictNames.to_excel(writer, sheet_name = 'Retention Report', startcol=0, index=False)
    RangeVal2.to_excel(writer, sheet_name='Retention Report', startcol=4, index=False)
    RangeValP2.to_excel(writer, sheet_name='Retention Report', startcol= 7, index=False)

    workbook = writer.book
    worksheet = writer.sheets['Retention Report']
    format1 = workbook.add_format({'bg_color': '#FFCCCB', 'bold': True})
    format2 = workbook.add_format({'bg_color': '#FFB214', 'bold': True})
    format3 = workbook.add_format({'bg_color': '#FFD070', 'bold': True})
    format4 = workbook.add_format({'bg_color': '#8CCBFF', 'bold': True})
    format5 = workbook.add_format({'bg_color': '#94ffb8', 'bold': True})

    format6 = workbook.add_format({'bg_color': '#ff0000', 'bold': True})
    format7 = workbook.add_format({'bg_color': '#00FF00', 'bold': True})

    formatPercentages = workbook.add_format({'num_format': '0%'})

    startR = 1
    endR = startR + DictNames.shape[0] - 1
    start_col = 1
    end_col = 1

    # Adding color to each patient's total attended sessions
    worksheet.conditional_format(startR, start_col, endR, end_col, {'type':'cell', 'criteria':'between', 'minimum': 0,'maximum': 3, 'format':format1})
    worksheet.conditional_format(startR, start_col, endR, end_col, {'type':'cell', 'criteria':'between', 'minimum': 4,'maximum': 7, 'format':format2})
    worksheet.conditional_format(startR, start_col, endR, end_col, {'type':'cell', 'criteria':'between', 'minimum': 8,'maximum': 10, 'format':format3})
    worksheet.conditional_format(startR, start_col, endR, end_col, {'type':'cell', 'criteria':'between', 'minimum': 11,'maximum': 14, 'format':format4})
    worksheet.conditional_format(startR, start_col, endR, end_col, {'type':'cell', 'criteria':'>=', 'value': 15, 'format':format5})

    worksheet.conditional_format('E2:E2', {'type': 'text','criteria':'containing', 'value':'0-3', 'format':   format1})
    worksheet.conditional_format('E3:E3', {'type': 'text','criteria':'containing', 'value':'4-7', 'format':   format2})
    worksheet.conditional_format('E4:E4', {'type': 'text','criteria':'containing', 'value':'8-10', 'format':   format3})
    worksheet.conditional_format('E5:E5', {'type': 'text','criteria':'containing', 'value':'11-14', 'format':   format4})
    worksheet.conditional_format('E6:E6', {'type': 'text','criteria':'containing', 'value':'15+', 'format':   format5})

    worksheet.conditional_format('H2:H2', {'type':'text','criteria':'containing', 'value':'0-3', 'format':   format1})
    worksheet.conditional_format('H3:H3', {'type':'text','criteria':'containing', 'value':'4-7', 'format':   format2})
    worksheet.conditional_format('H4:H4', {'type':'text','criteria':'containing', 'value':'8-10', 'format':   format3})
    worksheet.conditional_format('H5:H5', {'type':'text','criteria':'containing', 'value':'11-14', 'format':   format4})
    worksheet.conditional_format('H6:H6', {'type':'text','criteria':'containing', 'value':'15+', 'format':   format5})

    worksheet.set_column('I2:I6', 10, formatPercentages)

    for index in range(len(DictNames)):
        b = str(index+2)
        a = 'A'+b
        if (DictNames['Patient Name'].iloc[index] in Terminated_Patients):
            worksheet.conditional_format(a,{'type': 'no_blanks', 'format': format6})
        else:
            worksheet.conditional_format(a,{'type': 'no_blanks', 'format': format7})

    writer.save()

    # Statement that prints out if it made it to the end of the program
    print("Success")


    return file_name
