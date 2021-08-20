# -*- coding: utf-8 -*-
"""
Created on Mon Jul 12 17:14:06 2021

@author: Saahil Bhatia
"""

import pandas as pd
from ortools.linear_solver import pywraplp

# list multiplication function
def multiply(l1, l2):
    result = 0
    for i in range(len(l1)):
        result += l1[i]*l2[i]
    return result

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

# Initialize the Mixed-Integer Programming solver with the SCIP backend.
solver = pywraplp.Solver.CreateSolver('SCIP')

# Create Student Preference Vector from excel data
def getPreferences():
    preference = []
    for student in range(numStudents):
        for course in range(numCourses):
            if (studentFirstChoices[student] == (course + 1)):
                preference.append(5)
            elif (studentSecondChoices[student] == (course + 1)):
                preference.append(3)
            elif (studentThirdChoices[student] == (course + 1)):
                preference.append(1)
            else:
                preference.append(0)

    for course in range(numCourses):
        preference.append(int(5*numStudents*5/numCourses))
    return preference

# Matrix 'A' and Matrix 'B' combined (described in Google Doc)
def getConstraintMatrix():
    constraint_coeffs = []
    #A
    for student in range(numStudents):
        row = []
        for i in range(student*numCourses):
            row.append(0)
        for i in range(numCourses):
            row.append(1)
        for i in range(numStudents*numCourses-student*numCourses-numCourses):
            row.append(0)
        for course in range(numCourses):
            row.append(0)
        constraint_coeffs.append(row)
    #B
    for course in range(numCourses):
        row = []
        for j in range(course):
            row.append(0)
        row.append(1)
        for j in range(numCourses-course-1):
            row.append(0)
        row = row*numStudents
        
        for j in range(course):
            row.append(0)
        courseMax = courseMaxs[course]
        row.append(-courseMax)
        for j in range(numCourses-course-1):
            row.append(0)
        
        constraint_coeffs.append(row)
    #B2
    for course in range(numCourses):
        row = []
        for j in range(course):
            row.append(0)
        row.append(-1)
        for j in range(numCourses-course-1):
            row.append(0)
        row = row*numStudents
        
        for j in range(course):
            row.append(0)
        courseMin = courseMins[course]
        row.append(courseMin)
        for j in range(numCourses-course-1):
            row.append(0)
        
        constraint_coeffs.append(row)
    return constraint_coeffs

# Bounds of Matrix 'A' and Matrix 'B' combined (described in Google Doc)
def getConstraintBounds():
    #A
    bounds = [1]*numStudents
    #B and B2
    for i in range(2*numCourses):
        bounds.append(0)
    return bounds

# Consolidate the problem data
def create_data_model():
    numStudents = len(studentFirstChoices)
    numCourses = len(courseMins)
    data = {}
    data['obj_coeffs'] = getPreferences()
    data['constraint_coeffs'] = getConstraintMatrix()
    data['bounds'] = getConstraintBounds()
    data['num_vars'] = numCourses*numStudents+numCourses
    data['num_constraints'] = numStudents+(2*numCourses)
    return data

# Instantiate the data
data = create_data_model()



# Define the variable (schedule/input) vector for the solver
x = {}
for j in range(data['num_vars']):
    x[j] = solver.IntVar(0, 1, 'x[%i]' % j)
print('Number of variables =', solver.NumVariables())



# Define Constraints for the solver from the data (-1000 is a lazy way to bound this, change it later)
for i in range(data['num_constraints']):
    constraint = solver.RowConstraint(-1000, data['bounds'][i], '')
    for j in range(data['num_vars']):
        constraint.SetCoefficient(x[j], data['constraint_coeffs'][i][j])
print('Number of constraints =', solver.NumConstraints())



# Define the Objective for the solver from the data
objective = solver.Objective()
for j in range(data['num_vars']):
    objective.SetCoefficient(x[j], data['obj_coeffs'][j])
objective.SetMaximization()



# Solve the Problem
status = solver.Solve()
if status == pywraplp.Solver.OPTIMAL:
    print('Objective value =', solver.Objective().Value())
    print()
    print('Problem solved in %f milliseconds' % solver.wall_time())
    print('Problem solved in %d iterations' % solver.iterations())
    print('Problem solved in %d branch-and-bound nodes' % solver.nodes())
    print()
    
    solution = []
    for j in range(data['num_vars']):
        solution.append(int(x[j].solution_value()))
    
    
    # Statistics
    coursesRunning = []
    for course in range(numCourses):
        coursesRunning.append(solution[numCourses*numStudents+course])
    
    studentAssignments = [0]*numStudents
    for student in range(numStudents):
        for course in range(numCourses):
            if (solution[student*numCourses+course] == 1): 
                studentAssignments[student] = course + 1

    courseSizes = [0]*numCourses
    courseFirstChoices = [0]*numCourses
    courseSecondChoices = [0]*numCourses
    courseThirdChoices = [0]*numCourses
    for student in range(numStudents):
        for course in range(numCourses):
            assignment = studentAssignments[student]
            firstChoice = studentFirstChoices[student]
            secondChoice = studentSecondChoices[student]
            thirdChoice = studentThirdChoices[student]
            if ((course + 1) == firstChoice): 
                courseFirstChoices[course] += 1
            elif ((course + 1) == secondChoice):
                courseSecondChoices[course] += 1
            elif ((course + 1) == thirdChoice):
                courseThirdChoices[course] += 1
            if (solution[student*numCourses+course] == 1): 
                courseSizes[course] += 1

    numFirstChoiceAssignment = 0
    numSecondChoiceAssignment = 0
    numThirdChoiceAssignment = 0
    numNoAssignment = 0
    for student in range(numStudents):
        assignment = studentAssignments[student]
        firstChoice = studentFirstChoices[student]
        secondChoice = studentSecondChoices[student]
        thirdChoice = studentThirdChoices[student]
        if (assignment == firstChoice): 
            numFirstChoiceAssignment+=1
        elif (assignment == secondChoice):
            numSecondChoiceAssignment+=1
        elif (assignment == thirdChoice):
            numThirdChoiceAssignment+=1
        else:
            numNoAssignment+=1    
            
    prefobj = 5*numFirstChoiceAssignment+3*numSecondChoiceAssignment+numThirdChoiceAssignment
    print('Placement Objective Value: %d' % prefobj)
    print('First Choice Assignments: %.5f (%d/%d)' 
          % (float(numFirstChoiceAssignment)/numStudents, numFirstChoiceAssignment, numStudents))
    print('Second Choice Assignments: %.5f (%d/%d)' 
          % (float(numSecondChoiceAssignment)/numStudents, numSecondChoiceAssignment, numStudents))
    print('Third Choice Assignments: %.5f (%d/%d)'
          % (float(numThirdChoiceAssignment)/numStudents, numThirdChoiceAssignment, numStudents))
    print('No Assignments: %.5f (%d/%d)'
          % (float(numNoAssignment)/numStudents, numNoAssignment, numStudents))
else:
    print('The problem does not have an optimal solution.')