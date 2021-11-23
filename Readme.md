Problem Statement:
===============================================================
Given Excel sheet: Book1.xlsx 
The above file contains the data with Student Name, Date of Birth, Section, Marks etc.

Read the above xlsx file and calculate 
1. Age in years. 
2. Top 3 marks in each section.


Following is the output of the file :  test_normal.py
================================================================
C:\Users\user\AppData\Local\Programs\Python\Python39\python.exe C:/Users/user/Desktop/python-scripts/test_normal.py
Get Age of the students:
Name: Abc Age: 11 years
Name: abc2 Age: 10 years
Name: abc4 Age: 22 years
Name: abc9 Age: 11 years
Name: abc4 Age: 11 years
Name: Abc Age: 11 years
Name: abc4 Age: 11 years
Name: abc6 Age: 11 years
Name: abc10 Age: 10 years

Get top 3 marks of each section.
Section 'a' top 3 marks are :[99, 88, 88]
Section 'b' top 3 marks are :[88, 60, 55]
Section 'c' top 3 marks are :[77, 66, 55]
Time taken to execute : 0:00:00.069093

Process finished with exit code 0



Following is the output of the file :  test_optimized.py
================================================================
C:\Users\user\AppData\Local\Programs\Python\Python39\python.exe C:/Users/user/Desktop/python-scripts/test_normal.py

Get Age of the students:
     Name Date of Birth  Age
0     Abc    2010-07-10   11
2    abc2    2011-07-20   10
9    abc9    2010-07-10   11
15   abc4    2010-07-10   11
22    Abc    2010-07-10   11
26   abc4    2010-07-10   11
28   abc6    2010-07-10   11
32  abc10    2011-07-20   10

Get top 3 marks of each section.
    Name  Marks
4   abc4   88.0
9   abc9   88.0
8   abc8   99.0
15  abc4   55.0
19  abc8   60.0
18  abc7   88.0
28  abc6   55.0
27  abc5   66.0
23  abc1   77.0
Time taken to execute : 0:00:00.084644

Process finished with exit code 0