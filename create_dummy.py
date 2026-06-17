import pandas as pd

# Teacher | Class | Hours | Mon1 | Mon2 | Mon3
data = [
    ["TeacherA", "1-1", 1, "1", "1", "1"],
    ["TeacherB", "1-1", 1, "", "1", "1"],
    ["TeacherA", "1-2", 1, "1", "1", "1"],
    ["TeacherC", "1-1", 0, "生活", "", ""],
]

df = pd.DataFrame(data, columns=["Teacher", "Class", "Hours", "Mon1", "Mon2", "Mon3"])
df.to_excel("dummy_data.xlsx", index=False)
