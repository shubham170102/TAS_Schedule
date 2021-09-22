The matcher.py module can be used like this:
```
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
