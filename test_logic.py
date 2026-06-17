import pandas as pd
from ortools.sat.python import cp_model
from app import generate_timetable

df = pd.read_excel("dummy_data.xlsx")
timeslot_cols = ["Mon1", "Mon2", "Mon3"]

print("DataFrame:")
print(df)

success, df_class, df_teacher = generate_timetable(df, "Teacher", "Class", "Hours", timeslot_cols)
if success:
    print("Success")
else:
    print("Infeasible")
