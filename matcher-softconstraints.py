import pandas as pd
from pulp import *

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
S = len(studentFirstChoices)
C = len(courseMins)


model = LpProblem("TAS Matching", LpMaximize)

studentPreferences = [[0 for c in range(C)] for s in range(S)]
for s in range(S):
    c1 = (studentFirstChoices[s])-1
    c2 = (studentSecondChoices[s])-1
    c3 = (studentThirdChoices[s])-1
    studentPreferences[s][c1] = 5
    studentPreferences[s][c2] = 3   
    studentPreferences[s][c3] = 1

#variables (x, y, z)
studentAssignments = [[LpVariable(("Student %d is in Class %d" % (s, c)), 0, 1, LpInteger) 
                       for c in range(C)] for s in range(S)]
classIsEmpty = [LpVariable(("Class %d is empty" % c), 0, 1, LpInteger) 
                       for c in range(C)]
classDoesntMeetMinimum = [LpVariable(("Class %d doesn't meet the minimum" % c), 0, 1, LpInteger) 
                       for c in range(C)]

#Sum of students in each class
sumStudentsInClass = [LpAffineExpression() for c in range(C)]
for c in range(C):
    for s in range(S):
        sumStudentsInClass[c] += studentAssignments[s][c]

#Constraints for each class
constraints = [[LpConstraint() for c in range(C)] for i in range(3)]
for c in range(C):
    constraints[0][c] = sumStudentsInClass[c] + classIsEmpty[c] >= 1
    constraints[1][c] = sumStudentsInClass[c] + S*classIsEmpty[c] <= S
    constraints[2][c] = sumStudentsInClass[c] + courseMins[c]*classDoesntMeetMinimum[c] >= courseMins[c]

maxAssignmentConstraint = [LpConstraint() for s in range(S)]
for s in range(S):
    sumOfAssignments = LpAffineExpression()
    for c in range(C):
        sumOfAssignments += studentAssignments[s][c]
    maxAssignmentConstraint[s] = sumOfAssignments <= 1
    
#Objective function
objective = LpAffineExpression()
for c in range(C):
    placement = LpAffineExpression()
    for s in range(S):
        placement += studentPreferences[s][c]*studentAssignments[s][c]
    penalty = 3*courseMins[c]
    objective += placement
    objective += penalty*classIsEmpty[c]
    objective += -penalty*classDoesntMeetMinimum[c]

#send data to model
model += objective
for c in range(C):
    for i in range(len(constraints)):
        model += constraints[i][c]
for s in range(S):
    model += maxAssignmentConstraint[s]

model.writeMPS("TAS.mps")

model.writeLP("TAS.lp")
model.solve()
print("Status:", LpStatus[model.status])
print("Objective value: ", value(model.objective))