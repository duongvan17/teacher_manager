import pandas as pd

# This creates a perfect file that the app CANNOT reject
data = {
    'Name': ['Nguyen Van A', 'Tran Thi B'],
    'Subject': ['Mathematics', 'Computer Science'],
    'Time': ['08:00 --> 12:00', '13:00 --> 17:00']
}

df = pd.DataFrame(data)
df.to_excel('teachers_samplehh.xlsx', index=False)
print("✅ Created a perfect 'teachers_samplehh.xlsx'. Close Excel and try syncing now!")