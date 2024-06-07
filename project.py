import pandas as pd
from shareplum import Site, Office365
from shareplum.site import Version
 
# Authenticate and connect to the SharePoint site
authcookie = Office365('https://ssigroups.sharepoint.com', username='udomw@ssi-steel.com', password='Ab123456').GetCookies()
site = Site('https://ssigroups.sharepoint.com/sites/SSI-ITScouting-KnowledgeSharing', version=Version.v365, authcookie=authcookie)
 
# Access the SharePoint list
sp_list = site.List('CL1_S1007_Excel')
 
# Retrieve all items from the list
try:
    data = sp_list.GetListItems('All Items')
    # Create a pandas DataFrame
    data_df = pd.DataFrame(data)
    # Save DataFrame to an Excel file
    data_df.to_excel("data.xlsx", index=False)
    print("Data has been successfully exported to data.xlsx")
except Exception as e:
    print(f"An error occurred: {e}")



