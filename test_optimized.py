from datetime import datetime, date
import pandas as pd
import openpyxl
from dateutil.relativedelta import relativedelta

# Age in years for each student.
def get_age_in_years(df):
    """
    Function to return age in years
    It will skip the invalid values from the 'Date of Birth' column.
    :return:
    """
    #df = pd.DataFrame(data)
    # The below line will read the 'Date of Birth' column and replace empty value with NaT
    df['Date of Birth'] = pd.to_datetime(df['Date of Birth'], errors='coerce')

    # The below line is to drop NaT values from 'Date of Birth' column
    df = df.dropna(subset=['Date of Birth'])
    #------------------------------------------------------------------------------
    df["Age"] = df['Date of Birth'].apply(lambda x: (datetime.now().year - x.year))
    # 3 different execution time taken to execute above line
    # 1. 0:00:00.068631    2. 0:00:00.069022    3.0:00:00.078125
    # ------------------------------------------------------------------------------

    # ------------------------------------------------------------------------------
    # df["Age"] = [relativedelta(pd.to_datetime('now'), d).years for d in df['Date of Birth']]
    # 3 different execution time taken to execute above line
    # 1. 0:00:00.078129    2. 0:00:00.078124    3.0:00:00.084643
    # ------------------------------------------------------------------------------

    print(df[['Name', 'Date of Birth', 'Age']])

# get top 3 marks in each section:
def get_top3_marks_each_section(df):
    """
    Function to get top3 marks of each section.
    :return:
    """
    # Drop the rows with Marks value as empty.
    df1 = df.dropna(subset=['Marks'])

    # This line will sort the columns by Section and Marks then it group by Section
    # and returns last 3 entries i.e. top3 marks of each Section.
    output = df1.sort_values(['Section', 'Marks'], ascending=True).groupby('Section').tail(3)
    print(output[['Name', 'Marks']])


if __name__ == "__main__":
    """
    Main routine
    """
    # Get time at the time of execution start
    begin_time = datetime.now()
    pd.set_option('mode.chained_assignment', None)
    xls = pd.ExcelFile('Book1.xlsx')
    df = pd.read_excel(xls, 'Sheet1')

    # to get the age in years
    print("Get Age of the students:")
    get_age_in_years(df)

    # to get the top 3 marks in each section
    print("")
    print("Get top 3 marks of each section.")
    get_top3_marks_each_section(df)

    # Calculate time taken for execution
    time_taken = datetime.now() - begin_time
    print(f"Time taken to execute : {time_taken}")
