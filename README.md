### What does a matcher do?
1. Takes 2 files for input describing students and courses
    - Data needed to solve the problem
        - Student First Choices
        - Student Second Choices
        - Student Third Choices
        - Course Mins
        - Course Maxs
    - For outputting formatted results, we need the above plus:
        - Student first name
        - Student last name
        - Course name
2. Uses the input data to come up with a solution
3. Outputs stats, course data w/ sizes, student data w/ assignments, student unassignments

The [matcher.py](https://github.com/shubham170102/TAS_Schedule/blob/main/matcher.py) module allows us to instantiate a matcher, tell it which excel files to use, which sheet number in the excel file to use, and which columns to read data from before solving and outputing results. Currently, the column names should be passed as python dictionaries. The matcher should be called in this order:
1. matcher = HardConstraintMatcher(students_fileLocaion, students_SheetName, students_Columns, courses_FileLocation, courses_SheetName, courses_Columns)
2. matcher.solve()
3. matcher.outputResults()

### Example usage is as follows:
```
from matcher import HardConstraintMatcher

matcher = HardConstraintMatcher(
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
matcher.solve() #this is the step that takes a long time
matcher.outputResults()
```

### To specify these with a JSON file:
```
from matcher import HardConstraintMatcher
from json import load

with open('config.json') as config_file:
    config = load(config_file)
    
    matcher = HardConstraintMatcher(
        students_FileLocation = config["students_FileLocation"],
        students_SheetName = config["students_SheetName"],
        students_Columns = config["students_Columns"],
        courses_FileLocation = config["courses_FileLocation"],
        courses_SheetName = config["courses_SheetName"],
        courses_Columns = config["courses_Columns"]
    )
matcher.solve() #this is the step that takes a long time
matcher.outputResults()
```
config.json:
```
{
   "students_FileLocation":"data\\MOCK_Students.xlsx",
   "students_SheetName":0,
   "students_Columns":{
      "First_Name":"first_name",
      "Last_Name":"last_name",
      "P1":"Preference_1",
      "P2":"Preference_2",
      "P3":"Preference_3"
   },
   "courses_FileLocation":"data\\MOCK_Courses.xlsx",
   "courses_SheetName":1,
   "courses_Columns":{
      "Name":"Course Name",
      "Min":"Test Min",
      "Max":"Test Max"
   }
}
```

### Requires:
- Pulp
- Pandas
>Install these by running:
>`pip install pulp`
>`pip install pandas`

For the old hard constraints matcher script (not PuLP):
>Install Google OR-Tools by running:
>`python -m pip install --upgrade --user ortools`
