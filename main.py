"""
Lilit Yengoian
Vidrio Python Test
"""

import numpy as np
import pandas as pd
from _datetime import datetime
# Needed packages: et-xmlfile, openpyxl, xlrd

now = datetime.now()

# Asking the user for the location of the Bank Activity file
bank_file = input("Please enter location of the Bank Activity file: ")

# Loading both Mapping file and Bank Activity file using Pandas into individual Dataframes
bank_activity_df = pd.read_excel(bank_file)
mapping_df = pd.read_excel("mapping/Cash_Rec_Mapping.xlsx")

# Replacing any Nan fields in the Bank Activity file with blank string
bank_activity_df = bank_activity_df.replace(np.nan, '', regex=True)

# Creating blank dataframe to hold exceptions
exceptions_df = pd.DataFrame()

# Creating boolean variable to keep track of whether there are any exceptions
exceptions = False

# Adding columns to bank_activity_df
bank_activity_df["Bank Reference ID"] = bank_activity_df["Reference Number"]
bank_activity_df["Post Date"] = pd.to_datetime(bank_activity_df["Cash Post Date"])
bank_activity_df['Value Date'] = bank_activity_df['Cash Value Date']
bank_activity_df['Value Date'] = pd.to_datetime(bank_activity_df['Value Date'])
bank_activity_df['Amount'] = bank_activity_df['Transaction Amount Local']
bank_activity_df["Description"] = bank_activity_df["Transaction Description 1"] + " " + bank_activity_df[
    "Transaction Description 2"] + " " + bank_activity_df["Transaction Description 3"] + " " + bank_activity_df[
                                      "Transaction Description 4"] + " " + bank_activity_df[
                                      "Transaction Description 5"] + " " + bank_activity_df[
                                      "Transaction Description 6"] + bank_activity_df[
                                      "Detailed Transaction Type Name"] + bank_activity_df["Transaction Type"]
bank_activity_df['Bank Account'] = bank_activity_df['Cash Account Number']
bank_activity_df['Closing Balance'] = bank_activity_df['Closing Balance Local']
bank_activity_df['Filename'] = (
            bank_activity_df['Cash Account Number'].to_string() + now.strftime("%d/%m/%Y %H:%M:%S") + ".csv")

# Creating dataframe to hold all Bank Reference ID from Mapping file
refID_df = pd.DataFrame()
refID_df["Bank Reference ID"] = (mapping_df["Bank Ref ID"]).copy()

# Creating single dataframe to hold all the Bank Reference ID and Starting Balance values from the mapping file
refID_StartingBalance_df = pd.DataFrame()
refID_StartingBalance_df["Bank Reference ID"] = (mapping_df["Bank Ref ID"]).copy()
refID_StartingBalance_df["Starting_Balance"] = (mapping_df["Starting_Balance"]).copy()

# Using a for loop, looping through the Bank Reference IDs in refID_df
for refID in refID_df["Bank Reference ID"]:
    # Getting starting balance value from refID_StartingBalance_df
    starting_balance = refID_StartingBalance_df['Starting_Balance'].where(refID_df['Bank Reference ID'] == refID)

    # Creating an output dataframe that contains all the columns from Bank Activity dataframe where the
    # "Cash Account Number" equals the Bank Reference ID
    output_df = pd.DataFrame()
    output_df = bank_activity_df.copy()
    output_df = output_df.drop(output_df.index[output_df['Cash Account Number'] != refID])

    # Creating an MM dataframe that contains all MM activity in the output_df for the current Bank Reference ID
    mm_df = pd.DataFrame()
    mm_df = output_df.copy(deep=True)
    mm_df = mm_df.drop(mm_df.index[~mm_df['Description'].str.contains('STIF')])

    # Removing all MM activity from output dataframe
    output_df = output_df.drop(output_df.index[output_df['Description'].str.contains('STIF')])

    # Creating a write_file_df from the output_df containing all rows and specific columns
    write_file_df = output_df[['Bank Reference ID', 'Post Date', 'Value Date', 'Amount', 'Description', 'Bank Account',
                               'Closing Balance']].copy()

    # Removing all rows from the write_file DataFrame, where there is NA in the value of column 'Bank Reference ID'
    write_file_df = write_file_df.drop(write_file_df.index[write_file_df['Bank Reference ID'] == "NA"])

    # Represents MM investment value
    overnight_investment = 0
    # Represents calculated closing balance
    calc_closing_balance = 0
    # Represents bank closing balance
    bank_closing_balance = 0

    # If the write_file DataFrame is empty print the 'Bank Reference ID' + “ has no activity”
    if write_file_df.empty:
        print(refID, " has no activity")
    else:
        # Looking up the bank_closing_balance from the Bank Activity DataFrame for the given Bank Reference ID from
        # the “Closing_Balance” column
        for index,row in bank_activity_df.iterrows():
            if row["Bank Reference ID"] == refID:
                 bank_closing_balance = row["Closing Balance"]

        # Calculating the overnight MM investment from the MM dataframe (Assuming MM is equal to Transaction Amount
        # Local)
        for index, row in mm_df.iterrows():
            if row["Bank Reference ID"] == refID:
                overnight_investment = row["Closing Balance"]

        # Adding the starting balance row to the write_file DataFrame, populating the columns with given/calculated
        # values:
        write_file_df["Bank Reference ID"] = "Starting Balance"
        write_file_df["Post Date"] = '2020-01-01'
        write_file_df["Value Date"] = '2020-01-01'
        write_file_df["Amount"] = starting_balance
        write_file_df["Description"] = "Starting Balance"
        write_file_df["Bank Account"] = refID
        write_file_df["Closing Balance"] = 0

        # Calculating the closing balance from the write_file dataFrame by summing the “Amount” column
        calc_closing_balance = write_file_df['Amount'].sum()

        # Comparing calc_closing_balance, overnight_investment and bank_closing_balance
        if calc_closing_balance == overnight_investment and calc_closing_balance == bank_closing_balance:
            print("All values equal")
        else:
            # Adding a row to the exception dataframe, indicating the Bank Reference ID, bank_closing_balance, MM value,
            # and calculated_closing_balance
            exceptions_df['Bank Reference ID'] = [refID]
            exceptions_df['bank_closing_balance'] = [bank_closing_balance]
            exceptions_df['MM Value'] = [overnight_investment]
            exceptions_df['calculated_closing_balance'] = [calc_closing_balance]
            exceptions = True

        # Saving the write_file DataFrame to the Output sub folder as an Excel file
        filename = "output/" + str(refID) + " " + now.strftime("%d-%m-%Y %H:%M:%S") + ".xlsx"
        write_file_df.to_excel(filename, sheet_name="Bank Transactions")
        
        # Updating the mapping Dataframe for the relevant Bank Reference ID with the calculated closing balance
        mapping_df.loc[mapping_df['Bank Ref ID'] == refID, "Calculated Closing Balance"] = calc_closing_balance

# Saving the mapping Dataframe to excel, replacing the existing mapping file
mapping_df.to_excel("mapping/Cash_Rec_Mapping.xlsx", index=False)

# Writing the exception DataFrame to excel in the Output sub folder.
if not exceptions_df.empty:
    exceptions_df.to_excel("output/exceptions.xlsx")
