The matcher.py module can be used like this:
```
import matcher

m = matcher.HardConstraintMatcher(
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
)
m.solve() #this is the step that takes a long time
m.outputResults()
```

Requires:
- Pulp
- Pandas
>Install these by running:
>`pip install pulp`
>`pip install pandas`

For the regular hard constraints matcher script (not PuLP):
>Install OR-Tools by running:
>`python -m pip install --upgrade --user ortools`
