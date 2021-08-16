import pandas as pd

# Open Excel Files
studentsDataFile_Name = 'data\MOCK_Students.xlsx'
coursesDataFile_Name = 'data\MOCK_Courses.xlsx'

studentsDataFile = pd.read_excel(studentsDataFile_Name, sheet_name=0)
coursesDataFile = pd.read_excel(coursesDataFile_Name, sheet_name=1)

# Read Column data
studentFirstChoices = studentsDataFile['Preference 1'].tolist()
studentSecondChoices = studentsDataFile['Preference 2'].tolist()
studentThirdChoices = studentsDataFile['Preference 3'].tolist()
courseMins = coursesDataFile['Minimum/25'].tolist()
courseMaxs = coursesDataFile['Maximum/25'].tolist()
numStudents = len(studentFirstChoices)
numCourses = len(courseMins)


