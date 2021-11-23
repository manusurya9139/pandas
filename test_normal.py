import openpyxl
from datetime import *

workbook = openpyxl.load_workbook("Book1.xlsx")
workbook.active
sheet = workbook.worksheets[0]

# Age in years for each student.
def get_age_in_years():
    """
    Function to get the age of students in years.
    :return:
    """
    today = date.today()
    for value in sheet.values:
        age = 0
        if value[4] and value[4] != "Date of Birth":
            if isinstance(value[4], int):
                age = today.year - value[4]
            else:
                birth_date = datetime.strptime(str(value[4]), "%Y-%m-%d %H:%M:%S")
                age = today.year - birth_date.year
            if age < 0: continue
            print(f"Name: {value[0]} Age: {age} years")

# get top 3 marks in each section:
def get_top3_marks_each_section():
    """
    Function to get top3 marks of each section.
    :return:
    """
    sections = []
    for value in sheet.values:
        if value[2] and value[2] != "Section":
            sections.append(value[2])
    sections = list(set(sections))

    section_marks = {}
    for section in sections:
        marks = []
        for value in sheet.values:
            if section == value[2] and value[3]:
                marks.append(value[3])
                section_marks[section] = marks

    for k, v in section_marks.items():
        v = sorted(v, reverse=True)
        print(f"Section '{k}' top 3 marks are :{v[:3]}")


if __name__ == "__main__":
    begin_time = datetime.now()

    print("Get Age of the students:")
    get_age_in_years()

    print("")
    print("Get top 3 marks of each section.")
    get_top3_marks_each_section()
    # Calculate time taken for execution
    time_taken = datetime.now() - begin_time
    print(f"Time taken to execute : {time_taken}")

