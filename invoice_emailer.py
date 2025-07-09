import pandas as pd # For reading and manipulating excel files

# Create a path to the files in OneDrive. Use your own ID.
BASWARE_USERS_PATH = r"C:\Users\uac7200\OneDrive - enercity\basware Rechnungsbearbeitung - Technische Umsetzung - Dokumente\Technische Umsetzung\Rohdaten\basware\User-Stammdaten\User Rights Report.xlsx"
BASWARE_CURRENT_REPORT = r"C:\Users\uac7200\OneDrive - enercity\basware Rechnungsbearbeitung - Technische Umsetzung - Dokumente\Technische Umsetzung\Rohdaten\basware\Rechnungen (basware_080725).xlsx"
WORKDAY_USERS_PATH = r"C:\Users\uac7200\OneDrive - enercity\basware Rechnungsbearbeitung - Technische Umsetzung - Dokumente\Technische Umsetzung\Rohdaten\workday\Personalsuchliste_für_F_(RW)_monatlich.xlsx"

# Read the excel files
df_basware_users = pd.read_excel(BASWARE_USERS_PATH, sheet_name="User Rights")
df_basware_current_report = pd.read_excel(BASWARE_CURRENT_REPORT, sheet_name="Rechnungen")
df_workday_users = pd.read_excel(WORKDAY_USERS_PATH, sheet_name="Sheet1", header=1)

# Print the first entry of the files
print(df_basware_users.head())
print(df_basware_current_report.head())
print(df_workday_users.head())

# First step: add a Username column in df_basware_current_report from df_basware_users, by matching Empänger in df_current_report 
# with User in df_basware_users.
merged_step1 = df_basware_current_report.merge(
    df_basware_users[['User', 'Username']],
    left_on = 'Empfänger',
    right_on = 'User',
    how = 'left'
)

# Second step: add a OE column in the previously df_basware_current_report from df_workday_users, by matching username in merged_step1
# with UserID in df_workday_users.
df_final = merged_step2 = merged_step1.merge(
    df_workday_users[['UserID', 'OE']],
    left_on = 'Username',
    right_on = 'UserID',
    how = 'left'
)

# Show the result
print(df_final.head())
df_final.to_excel("Verknüpfung_basware_workday.xlsx")





