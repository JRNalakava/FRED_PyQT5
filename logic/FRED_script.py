import pandas as pd
def run(rPath, iPath, date_start, date_end):
    # Creates the two dataframes for each file
    DF = pd.read_excel(rPath)
    DF2 = pd.read_excel(iPath)
    #TODO Develop date functionality
    year_needed = date_start.year()
    # Calls the filterInactive function that labels each patient as inactive(Terminated) or active
    DF = filterInactive(DF, DF2)
    # Calls the filter_DF function and passes in the raw dataframe with the Active/Inactive column filled out for each patient
    filtered_DFs = filter_DF(DF, year_needed)
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
    Inactive_names = list(DF2['Full Name'])

    # Creates a column in the raw data to fill out if the patient is terminated or not
    DF['Active/Inactive'] = " "

    # For loop that iterates through all the rows in the raw dataframe
    for i, row in DF.iterrows():
        # Checks if the patients name in the raw dataframe is in the list of terminated patients
        if row['Full Name'] in Inactive_names:
            # If the patient's name is in the list of terminated patients, labels them as inactive in the Active/Inactive column
            DF.at[i, 'Active/Inactive'] = 'Inactive'
        else:
            # Fills out the Active/Inactive column as active if the patient's name is not in the list of terminated patients
            DF.at[i, 'Active/Inactive'] = 'Active'

    return DF

# Function that filters out all the group therapy patients and non appointment rows in the patient dataframe
def filter_DF(DF, year_needed):
    ###DF['Full Name'] = DF['First Name'].str.cat(DF['Last Name'],sep=" ")

    # Changes the date column in the dataframe into an easier format to parse
    DF['Date'] = pd.to_datetime(DF['Date'])

    # If statement checking that the user wants all time statsitics
    if year_needed == 0:

        # Creates a filtered dataframe that only includes appointments that are therapy sessions or therapy intake
        Therapy_Sessions = DF.loc[(DF['Type'] == 'Appointment') & (DF['Appointment Type'] == 'Therapy Session') | (DF['Appointment Type'] == 'Therapy Intake')]
        # Creates a filtered dataframe that only includes appointments that are therapy sessions from patients that are terminated
        Inactive_Clients = DF.loc[(DF['Type'] == 'Appointment') & (DF['Appointment Type'] == 'Therapy Session') | (DF['Appointment Type'] == 'Therapy Intake') & (DF['Active/Inactive'] == 'Inactive')]
    else:

        # Creates a filtered dataframe that only includes appointments that are therapy sessions or therapy intake from the user's inputted year
        Therapy_Sessions = DF.loc[(DF['Type'] == 'Appointment') & (DF['Appointment Type'] == 'Therapy Session') | (DF['Appointment Type'] == 'Therapy Intake') & (DF['Date'].dt.year == year_needed)]
        # Creates a filtered dataframe that only includes appointments that are therapy sessions or therapy intake from patients that are terminated in the user's inputted year
        Inactive_Clients = DF.loc[(DF['Type'] == 'Appointment') & (DF['Appointment Type'] == 'Therapy Session') | (DF['Appointment Type'] == 'Therapy Intake') & (DF['Date'].dt.year == year_needed) & (DF['Active/Inactive'] == 'Inactive')]

    return [Therapy_Sessions, Inactive_Clients]

# Function that counts the number of times  each patient has a therapy session
def countAttendance(Therapy, Inactive):

    # Creates a dictionary with the key as the name of the patient and the value as the number of times they had a therapy session for all patients
    Dict_Values = dict(Therapy['Full Name'].value_counts())

    # Creates a dictionary with the key as the name of the patient and the value as the number of times they had a therapy session from only terminated patients
    Dict2_Values = dict(Inactive['Full Name'].value_counts())

    return [Dict_Values, Dict2_Values]

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
        ZtoThreeP2 = round((ZeroToThree2/Total2)*100, 2)
    else:
        # Sets the number equal to the range variable since the number of times is zero
        ZtoThreeP2 = ZeroToThree2

    if FourToSeven2 != 0:
        # Divides the number of times a patient had total therapy sessions attended within the 4-7 range by the total amount of therapy sessions attended
        # Multiplies the number by 100 and rounds to two decimal points in order to be a percentage
        FtoSevenP2 = round((FourToSeven2/Total2)*100, 2)
    else:
        # Sets the number equal to the range variable since the number of times is zero
        FtoSevenP2 = FourToSeven2

    if EightToTen2 != 0:
        # Divides the number of times a patient had total therapy sessions attended within the 8-10 range by the total amount of therapy sessions attended
        # Multiplies the number by 100 and rounds to two decimal points in order to be a percentage
        EtoTenP2 = round((EightToTen2/Total2)*100, 2)
    else:
        # Sets the number equal to the range variable since the number of times is zero
        EtoTenP2 = EightToTen2

    if ElevenToFourteen2 != 0:
        # Divides the number of times a patient had total therapy sessions attended within the 11-14 range by the total amount of therapy sessions attended
        # Multiplies the number by 100 and rounds to two decimal points in order to be a percentage
        ElToFourteenP2 = round((ElevenToFourteen2/Total2)*100, 2)
    else:
        # Sets the number equal to the range variable since the number of times is zero
        ElToFourteenP2 = ElevenToFourteen2

    if FifteenPlus2 != 0:
        # Divides the number of times a patient had total therapy sessions where they attended 15+ by the total amount of therapy sessions attended
        # Multiplies the number by 100 and rounds to two decimal points in order to be a percentage
        FifteenPlusP2 = round((FifteenPlus2/Total2)*100, 2)
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
    DictNames = pd.DataFrame.from_dict(Dict1, orient='index')

    # Creates a dataframe from the dictionary that has all the range bins and the number of times a patient's total therapy sessions attended was within that range
    RangeVal2 = pd.DataFrame.from_dict(Range_Vals2, orient='index')

    # Creates a dataframe from the dictionary that has all the range bins and the percentage of the number of times a patient's total therapy sessions attended was within that range by the total for all ranges
    RangeValP2 = pd.DataFrame.from_dict(Range_Val_Percents2, orient='index')

    # Creates a string holding the file name 'Retention Report'
    file_name = 'Retention_Report.xlsx'

    # Statement that takes each dataframe and puts it into an excel file with the sheet name being 'Retention Report'
    with pd.ExcelWriter(file_name, mode='A') as writer:
        DictNames.to_excel(writer, sheet_name = 'Retention Report', startcol=0)
        RangeVal2.to_excel(writer, sheet_name='Retention Report', startcol=4)
        RangeValP2.to_excel(writer, sheet_name='Retention Report', startcol= 7)

    # Statement that prints out if it made it to the end of the program
    print("Success")

    return file_name