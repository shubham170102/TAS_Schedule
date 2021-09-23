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