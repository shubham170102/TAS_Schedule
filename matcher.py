''' 
What does a matcher do?
    Takes 2 files for input and outputs 4 files
        -takes input describing students and courses
            -Whats the input data we need to solve the problem?
                -For making a solution, we need:
                    -student First Choices
                    -student Second Choices
                    -student Third Choices
                    -course Mins
                    -course Maxs
                -For outputting results formatted, we need the above plus:
                    -student first name
                    -student last name
                    -student id
                    -course name
                    -course id
        -use the input data to come up with a solution
        -outputs stats, course data w/ sizes, student data w/ assignments, student unassignments

    What are the steps for a matcher in general???
        -take in input
        -solve
        -output results
        
    What is common for all matchers???
        -inputs are all the same format
        -each matcher has to match/solve the problem
        -outputs are all the same format

    What are the steps for a matcher based on PuLP linear programming???
        -read from excel files for students and courses and put columns into arrays/lists
        
        -solve the model
            -initialize a model and a solver
            -initialize variables
            -create constraints from the variables and input excel data
            -create an objective equation from the variables and input excel data
            -add this the constraints and objective to the model
            -use the solver to solve the model
        
        -output results
    
    Whats the difference between the hard constraints matcher and a soft constraints matcher?
        -initialize two sets of penalty variables corrisponding to each course
        -add a penalty variable to each course/class constraints
        -add a penalty term to the objective for each penalty variable
'''

import pandas as pd
import csv
from pulp import LpProblem, LpMaximize, getSolver, LpVariable, LpInteger, LpAffineExpression, LpConstraint, LpStatus, value

class Matcher:
    """
    Example Usage:
        import matcher
        
        m = matcher.HardConstraintMatcher(
            students_FileLocation = 'Students.xlsx',
            students_SheetName = 0,
            students_Columns = {
                "First_Name": "First Name",
                "Last_Name": "Last Name",
                "P1": "Preference 1",
                "P2": "Preference 2",
                "P3": "Preference 3",
            },
            courses_FileLocation = 'Courses.xlsx',
            courses_SheetName = 0,
            courses_Columns = {
                "Name": "Course Name",
                "Min": "Minimum Size",
                "Max": "Maximum Size"
            }
        )
        m.solve() #this is the step that takes a long time
        m.outputResults()
    """
    
    def __init__(self,
            # Default Inputs
            students_FileLocation = 'data\MOCK_Students.xlsx',
            students_SheetName = 0,
            students_Columns = {
                "First_Name": "first_name",
                "Last_Name": "last_name",
                "P1": "Preference 1",
                "P2": "Preference 2",
                "P3": "Preference 3",
            },
            courses_FileLocation = 'data\MOCK_Courses.xlsx',
            courses_SheetName = 1,
            courses_Columns = {
                "Name": "Course Name",
                "Min": "Test Min",
                "Max": "Test Max"
            }
        ):
        
        # Open Excel File
        studentsData = pd.read_excel(students_FileLocation, 
                                     sheet_name = students_SheetName)
        coursesData = pd.read_excel(courses_FileLocation, 
                                    sheet_name = courses_SheetName)
        
        # Read Column data
        self.studentFirstName = studentsData[students_Columns["First_Name"]].tolist()
        self.studentLastName = studentsData[students_Columns["Last_Name"]].tolist()
        self.courseNames = coursesData[courses_Columns["Name"]].tolist()
        self.S = len(self.studentFirstName)
        self.C = len(self.courseNames)
        
        self.studentFirstChoices = studentsData[students_Columns["P1"]].tolist()
        self.studentSecondChoices = studentsData[students_Columns["P2"]].tolist()
        self.studentThirdChoices = studentsData[students_Columns["P3"]].tolist()
        self.courseMins = coursesData[courses_Columns["Min"]].tolist()
        self.courseMaxs = coursesData[courses_Columns["Max"]].tolist()
    
    ##### The folowing 3 functions should be expanded upon for each type of Matcher #####
    def initVariables(self):
        self.studentAssignments = [[LpVariable("S%dC%d"%(s,c), 0, 1, LpInteger) 
                       for c in range(self.C)] for s in range(self.S)]
    
    def makeObjective(self):
        #Student preference vector (for the objective function)         
        self.studentPreferences = [[0 for c in range(self.C)] for s in range(self.S)]
        
        for s in range(self.S):
            coursePreferences = []
            for c in range(self.C):
                if (self.studentFirstChoices[s] == (c + 1)):
                    coursePreferences.append(5)
                elif (self.studentSecondChoices[s] == (c + 1)):
                    coursePreferences.append(3)
                elif (self.studentThirdChoices[s] == (c + 1)):
                    coursePreferences.append(1)
                else:
                    coursePreferences.append(0)
            self.studentPreferences[s] = coursePreferences
        # for s in range(self.S):
        #     c1 = (self.studentFirstChoices[s])-1
        #     c2 = (self.studentSecondChoices[s])-1
        #     c3 = (self.studentThirdChoices[s])-1
        #     #students who bullet vote will have their single preference value be 1
        #     self.studentPreferences[s][c1] = 5
        #     self.studentPreferences[s][c2] = 3   
        #     self.studentPreferences[s][c3] = 1
    
    def makeConstraints(self):
        self.sumStudentsInClass = [LpAffineExpression() for c in range(self.C)]
        for c in range(self.C):
            for s in range(self.S):
                self.sumStudentsInClass[c] += self.studentAssignments[s][c]
    
    def initProblem(self):
        self.model = LpProblem("TAS Matching", LpMaximize)
        self.initVariables()
        self.makeObjective()
        self.makeConstraints()
    
    def solve(self, solver = "PULP_CBC_CMD"):
        self.initProblem()
        self.model.writeLP("TAS.lp")
        self.model.solve(getSolver(solver))
        print("Status:", LpStatus[self.model.status])
        print("Objective value: ", value(self.model.objective))
    
    def outputResults(self):
        #Convert the variables to their values
        for c in range(self.C):
            for s in range(self.S):    
                self.studentAssignments[s][c] = int(self.studentAssignments[s][c].varValue)
        
        #Output Results
        numFirstChoiceAssignment = 0
        numSecondChoiceAssignment = 0
        numThirdChoiceAssignment = 0
        numNoChoiceAssignment = 0
        numNoAssignment = 0
        numMultiAssignment = 0
        studentIdAssignments = [[-1] for s in range(self.S)]
        courseSizes = [0 for c in range(self.C)]
        courseFirstChoices = [0 for c in range(self.C)]
        courseSecondChoices = [0 for c in range(self.C)]
        courseThirdChoices = [0 for c in range(self.C)]
        for s in range(self.S):
            firstChoice = self.studentFirstChoices[s]
            secondChoice = self.studentSecondChoices[s]
            thirdChoice = self.studentThirdChoices[s]
            sumAssignments = 0
            for c in range(self.C):        
                #Courses Stats
                courseId = c+1
                if (courseId == firstChoice): 
                    courseFirstChoices[c] += 1
                elif (courseId == secondChoice):
                    courseSecondChoices[c] += 1
                elif (courseId == thirdChoice):
                    courseThirdChoices[c] += 1  
                if (self.studentAssignments[s][c] == 1): 
                    courseSizes[c] += 1
                    
                #Students Stats
                if (self.studentAssignments[s][c] == 1): 
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
              % (float(numFirstChoiceAssignment)/self.S, numFirstChoiceAssignment, self.S))
        print('Second Choice Assignments: %.5f (%d/%d)' 
              % (float(numSecondChoiceAssignment)/self.S, numSecondChoiceAssignment, self.S))
        print('Third Choice Assignments: %.5f (%d/%d)'
              % (float(numThirdChoiceAssignment)/self.S, numThirdChoiceAssignment, self.S))
        print('No Choice Assignments: %.5f (%d/%d)'
              % (float(numNoChoiceAssignment)/self.S, numNoChoiceAssignment, self.S))
        print('Multi Assignments: %.5f (%d/%d)'
              % (float(numMultiAssignment)/self.S, numMultiAssignment, self.S))
        print('No Assignments: %.5f (%d/%d)'
              % (float(numNoAssignment)/self.S, numNoAssignment, self.S))
        
        
        #Output to CSV
        with open('Output_Courses.csv', mode='w') as course_file:
            course_writter = csv.writer(course_file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL, lineterminator = '\n')
            #write column names
            columnNames = []
            columnNames.append('Course number')
            columnNames.append('Course name')
            columnNames.append('First choice')
            columnNames.append('Second choice')
            columnNames.append('Third choice')
            columnNames.append('Weight')
            columnNames.append('Minimum class size')
            columnNames.append('Maximum class size')
            columnNames.append('Students assigned')
            course_writter.writerow(columnNames)
            
            #write row data
            for c in range(self.C):
                rowData = []
                c1 = courseFirstChoices[c]
                c2 = courseSecondChoices[c]
                c3 = courseThirdChoices[c]
                
                rowData.append(c+1) #'Course number'
                rowData.append(self.courseNames[c]) #'Course name'
                rowData.append(c1) #'First choice'
                rowData.append(c2) #'Second choice'
                rowData.append(c3) #'Third choice'
                rowData.append(5*c1+3*c2+c3) #'Weight'
                rowData.append(self.courseMins[c]) #'Minimum class size'
                rowData.append(self.courseMaxs[c]) #'Maximum class size'
                rowData.append(courseSizes[c]) #'Students assigned'
                course_writter.writerow(rowData)
        
        with open('Output_Assigned_Students.csv', mode='w') as course_file:
            student_writer = csv.writer(course_file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL, lineterminator = '\n')
            #write column names
            columnNames = []
            columnNames.append('Student ID')
            columnNames.append('First Name')
            columnNames.append('Last Name')
            columnNames.append('Course Assignment')
            student_writer.writerow(columnNames)
            
            #write row data
            for s in range(self.S):
                if (studentIdAssignments[s][0] != -1):
                    ca = "%d"%studentIdAssignments[s][0]
                    for i in range(1, len(studentIdAssignments[s])):
                        ca += ", %d"%studentIdAssignments[s][i]
                    
                    rowData = []
                    rowData.append(s+1)#'Student ID'
                    rowData.append(self.studentFirstName[s])#'First Name'
                    rowData.append(self.studentLastName[s])#'Last Name'
                    rowData.append(ca)#'Course Assignment'
                    student_writer.writerow(rowData)
                
        with open('Output_Unassigned_Students.csv', mode='w') as course_file:
            student_writer = csv.writer(course_file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL, lineterminator = '\n')
            #write column names
            columnNames = []
            columnNames.append('Student ID')
            columnNames.append('First Name')
            columnNames.append('Last Name')
            student_writer.writerow(columnNames)
            
            #write row data
            for s in range(self.S):
                if (studentIdAssignments[s][0] == -1):            
                    rowData = []
                    rowData.append(s+1)#'Student ID'
                    rowData.append(self.studentFirstName[s])#'First Name'
                    rowData.append(self.studentLastName[s])#'Last Name'
                    student_writer.writerow(rowData)
                    
        with open('Output_Stats.csv', mode='w') as course_file:
            stats_writer = csv.writer(course_file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL, lineterminator = '\n')
            #write column names
            stats_writer.writerow(['First Choice Assignments: %.5f (%d/%d)' 
                  % (float(numFirstChoiceAssignment)/self.S, numFirstChoiceAssignment, self.S)])
            stats_writer.writerow(['Second Choice Assignments: %.5f (%d/%d)' 
                  % (float(numSecondChoiceAssignment)/self.S, numSecondChoiceAssignment, self.S)])
            stats_writer.writerow(['Third Choice Assignments: %.5f (%d/%d)'
                  % (float(numThirdChoiceAssignment)/self.S, numThirdChoiceAssignment, self.S)])
            stats_writer.writerow(['No Choice Assignments: %.5f (%d/%d)'
                  % (float(numNoChoiceAssignment)/self.S, numNoChoiceAssignment, self.S)])
            stats_writer.writerow(['Multi Assignments: %.5f (%d/%d)'
                  % (float(numMultiAssignment)/self.S, numMultiAssignment, self.S)])
            stats_writer.writerow(['No Assignments: %.5f (%d/%d)'
                  % (float(numNoAssignment)/self.S, numNoAssignment, self.S)])

class HardConstraintMatcher(Matcher):
    def initVariables(self):
        super(HardConstraintMatcher, self).initVariables()
        self.classWillRun = [LpVariable("C%dr"%c, 0, 1, LpInteger) 
                       for c in range(self.C)]
    
    def makeObjective(self):
        super(HardConstraintMatcher, self).makeObjective()
        objective = LpAffineExpression()
        for c in range(self.C):
            for s in range(self.S):
                objective += self.studentPreferences[s][c]*self.studentAssignments[s][c]
            #objective += (5*self.S/self.C)*self.classWillRun[c]
            
        #Add objective to model
        self.model += objective
    
    def makeConstraints(self):
        super(HardConstraintMatcher, self).makeConstraints()
        
        #Constraints described by Dr. Miller in email
        # runConstraints = [[LpConstraint() for s in range(self.S)] for c in range(self.C)]
        # for c in range(self.C):
        #     for s in range(self.S):  
        #         constraint = self.studentAssignments[s][c] - self.classWillRun[c]
        #         runConstraints[c][s] = LpConstraint(e=constraint, sense=-1, name="C%dS%dr"%(c, s), rhs=0)
        # sizeConstraints = [LpConstraint() for c in range(self.C)]    
        # for c in range(self.C):
        #     constraint = self.classWillRun[c] - self.sumStudentsInClass[c]
        #     sizeConstraints[c] = LpConstraint(e=constraint, sense=-1, name="C%ds"%c, rhs=0)
        
        #Constraints for each class
        classConstraints = [[LpConstraint() for c in range(self.C)] for i in range(2)]
        for c in range(self.C):    
            hardMin = self.sumStudentsInClass[c] - self.courseMins[c]*self.classWillRun[c]
            hardMax = self.sumStudentsInClass[c] - self.courseMaxs[c]*self.classWillRun[c]
            cMin = LpConstraint(e=hardMin, sense=1, name="C%dm"%c, rhs=0)
            cMax = LpConstraint(e=hardMax, sense=-1, name="C%dM"%c, rhs=0)
            classConstraints[0][c] = cMin
            classConstraints[1][c] = cMax
        
        #Constraint limiting the number of classes a student should be assigned to
        maxAssignmentConstraint = [LpConstraint() for s in range(self.S)]
        for s in range(self.S):
            sumOfAssignments = LpAffineExpression()
            for c in range(self.C):
                sumOfAssignments += self.studentAssignments[s][c]
            maxAssignmentConstraint[s] = sumOfAssignments <= 1
    
        #Add constraints to model
        for s in range(self.S):
            self.model += maxAssignmentConstraint[s]
        for c in range(self.C):
            #Dr. Miller's Constraints
            # for s in range(self.S):
            #     self.model += runConstraints[c][s]
            # self.model += sizeConstraints[c]
            for i in range(len(classConstraints)):
                self.model += classConstraints[i][c]



