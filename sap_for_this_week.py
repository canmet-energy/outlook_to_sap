# create the TableMaker object.
from lib.OutlookToPandas import *
outlook = OutlookToPandas()
now = datetime.datetime.now()
weeknumber = now.isocalendar()[1]
print(f"Getting hours for week #{weeknumber} from Outlook.")
outlook = OutlookToPandas()
print(outlook.create_week_sap_report(now.year, weeknumber))
print('Week is now copied to your clipboard. Paste it in SAP. ')
