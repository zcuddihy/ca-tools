#%%
from src.outlook import OutlookConnection
from src.project import Project
from src.rfi_log import RFI_Log

# Variables for setting up the project
project_name = "OHMC"
excel_file = "OHMC NET - CA Log.xlsx"
main_file_path = "I:\OHMC-Exp\Construction Admin"

# Set up base project information
project_obj = Project(project_name, excel_file, main_file_path)

# Connect to outlook and check for RFI's and submittals
Inbox = OutlookConnection()
Inbox.check_email()

# Log and save all new RFIs
new_RFIs = RFI_Log(Inbox.new_RFI, project_obj)
new_RFIs.save()

# %%
