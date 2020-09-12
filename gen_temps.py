from datetime import datetime,timedelta
from pandas import ExcelWriter
import pandas as pd
from datetime import time as column_time
column_names = ["A","B","C","D","E","F","G"]

time_slot = datetime(datetime.now().year, datetime.now().month, datetime.now().day, 15, 30)
for i in range(20):
    column_names.append(column_time(time_slot.hour, time_slot.minute))
    time_slot+=timedelta(0,300)
templateautocopynifty200 = pd.DataFrame(columns=column_names)

with ExcelWriter("/home/alexis/projects/misc/excel_copy_paste/AutoCopyNifty200.xls") as writer:
    templateautocopynifty200.to_excel(writer, index=False, sheet_name="Total Qty")
    templateautocopynifty200.to_excel(writer, index=False, sheet_name="Buy Qty")
    templateautocopynifty200.to_excel(writer, index=False, sheet_name="Sell Qty")
    writer.save()

with ExcelWriter("/home/alexis/projects/misc/excel_copy_paste/NSE Cash.xls") as writer:
    templateautocopynifty200.to_excel(writer, index=False, sheet_name="Total Qty")
    templateautocopynifty200.to_excel(writer, index=False, sheet_name="Buy Qty")
    templateautocopynifty200.to_excel(writer, index=False, sheet_name="Sell Qty")
    writer.save()
    
    
column_names = ["A"]

time_slot = datetime(datetime.now().year, datetime.now().month, datetime.now().day, 9, 30)
for i in range(13):
    column_names.append(column_time(time_slot.hour, time_slot.minute))
    time_slot+=timedelta(0,1800)
copy_trigger_formula_nifty = pd.DataFrame(columns=column_names)

copy_trigger_formula_nifty.to_excel("Nifty BankNifty Formula.xls", index=False, sheet_name="Sheet1")
