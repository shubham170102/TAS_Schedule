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
courseMins = coursesDataFile['Test Min'].tolist()
courseMaxs = coursesDataFile['Test Max'].tolist()
S = len(studentFirstChoices)
C = len(courseMins)


model = LpProblem("TAS Matching", LpMaximize)
solver = getSolver('PULP_CBC_CMD')

#variables (x, y, z, q)
studentAssignments = [[LpVariable(("S%d:C%d" % (s, c)), 0, 1, LpInteger) 
                       for c in range(C)] for s in range(S)]
classIsEmpty = [LpVariable(("C%dE" % c), 0, 1, LpInteger) 
                       for c in range(C)]
classDoesntMeetMinimum = [LpVariable(("C%dLM" % c), 0, 1, LpInteger) 
                       for c in range(C)]
classOverMaximum = [LpVariable(("C%dGM" % c), 0, 1, LpInteger) 
                       for c in range(C)]

#Sum of students in each class
sumStudentsInClass = [LpAffineExpression() for c in range(C)]
for c in range(C):
    for s in range(S):
        sumStudentsInClass[c] += studentAssignments[s][c]

#Constraints for each class
constraints = [[LpConstraint() for c in range(C)] for i in range(4)]
for c in range(C):
    constraints[0][c] = sumStudentsInClass[c] + classIsEmpty[c] >= 1
    constraints[1][c] = sumStudentsInClass[c] + S*classIsEmpty[c] <= S
    constraints[2][c] = sumStudentsInClass[c] + courseMins[c]*classDoesntMeetMinimum[c] >= courseMins[c]
    constraints[3][c] = sumStudentsInClass[c] - S*classOverMaximum[c] + courseMaxs[c]*classOverMaximum[c] <= courseMaxs[c]

maxAssignmentConstraint = [LpConstraint() for s in range(S)]
for s in range(S):
    sumOfAssignments = LpAffineExpression()
    for c in range(C):
        sumOfAssignments += studentAssignments[s][c]
    maxAssignmentConstraint[s] = sumOfAssignments <= 1
    
    
#Student preference vector (for the objective function)
studentPreferences = [[0 for c in range(C)] for s in range(S)]
for s in range(S):
    c1 = (studentFirstChoices[s])-1
    c2 = (studentSecondChoices[s])-1
    c3 = (studentThirdChoices[s])-1
    studentPreferences[s][c1] = 5
    studentPreferences[s][c2] = 3   
    studentPreferences[s][c3] = 1
    
#Objective function
objective = LpAffineExpression()
for c in range(C):
    placement = LpAffineExpression()
    for s in range(S):
        placement += studentPreferences[s][c]*studentAssignments[s][c]
    penalty = 10
    objective += placement
    objective += penalty*classIsEmpty[c]
    objective += (-penalty)*classDoesntMeetMinimum[c]
    objective += (-penalty)*classOverMaximum[c]

#send data to model
model += objective
for s in range(S):
    model += maxAssignmentConstraint[s]
for c in range(C):
    for i in range(len(constraints)):
        model += constraints[i][c]
model.writeMPS("TAS.mps")

#Solve
model.writeLP("TAS.lp")
model.solve(solver)
print("Status:", LpStatus[model.status])
print("Objective value: ", value(model.objective))

#Convert the variables to their values
for c in range(C):
    for s in range(S):    
        studentAssignments[s][c] = int(studentAssignments[s][c].varValue)
    classIsEmpty[c] = classIsEmpty[c].varValue
    classDoesntMeetMinimum[c] = classDoesntMeetMinimum[c].varValue
    classOverMaximum[c] = classOverMaximum[c].varValue

#Output Results
numFirstChoiceAssignment = 0
numSecondChoiceAssignment = 0
numThirdChoiceAssignment = 0
numNoChoiceAssignment = 0
numNoAssignment = 0
numMultiAssignment = 0
studentIdAssignments = [[-1] for s in range(S)]
courseSizes = [0 for c in range(C)]
courseFirstChoices = [0 for c in range(C)]
courseSecondChoices = [0 for c in range(C)]
courseThirdChoices = [0 for c in range(C)]
for s in range(S):
    firstChoice = studentFirstChoices[s]
    secondChoice = studentSecondChoices[s]
    thirdChoice = studentThirdChoices[s]
    sumAssignments = 0
    for c in range(C):        
        #Courses Stats
        courseId = c+1
        if (studentAssignments[s][c] == 1): 
            courseSizes[c] += 1
        if (courseId == firstChoice): 
            courseFirstChoices[c] += 1
        elif (courseId == secondChoice):
            courseSecondChoices[c] += 1
        elif (courseId == thirdChoice):
            courseThirdChoices[c] += 1  
            
        #Students Stats
        if (studentAssignments[s][c] == 1): 
            if (courseId == firstChoice): 
                numFirstChoiceAssignment+=1
            elif (courseId == secondChoice):
                numSecondChoiceAssignment+=1
            elif (courseId == thirdChoice):
                numThirdChoiceAssignment+=1
            else:
                numNoChoiceAssignment+=1
            sumAssignments+=1
            if studentIdAssignments[s][0] == -1:
                studentIdAssignments[s][0] = courseId
            else:
                studentIdAssignments[s].append(courseId)
    if sumAssignments > 1:
        numMultiAssignment+=1
    elif sumAssignments == 0:
        numNoAssignment+=1
        
#Print some stats            
prefobj = 5*numFirstChoiceAssignment+3*numSecondChoiceAssignment+numThirdChoiceAssignment
print('Placement Objective Value: %d' % prefobj)
print('First Choice Assignments: %.5f (%d/%d)' 
      % (float(numFirstChoiceAssignment)/S, numFirstChoiceAssignment, S))
print('Second Choice Assignments: %.5f (%d/%d)' 
      % (float(numSecondChoiceAssignment)/S, numSecondChoiceAssignment, S))
print('Third Choice Assignments: %.5f (%d/%d)'
      % (float(numThirdChoiceAssignment)/S, numThirdChoiceAssignment, S))
print('No Choice Assignments: %.5f (%d/%d)'
      % (float(numNoChoiceAssignment)/S, numNoChoiceAssignment, S))
print('Multi Assignments: %.5f (%d/%d)'
      % (float(numMultiAssignment)/S, numMultiAssignment, S))
print('No Assignments: %.5f (%d/%d)'
      % (float(numNoAssignment)/S, numNoAssignment, S))
