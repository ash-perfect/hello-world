from datetime import time as column_time
import pandas as pd
from pandas import ExcelWriter

from datetime import datetime, timedelta
import math
import time

startH, startM = 9,30
endH, endM = 15,30
minutes_between_copying = 30
time_between_copying = minutes_between_copying * 60

#File paths
trigger_formula_original_path = "/home/alexis/projects/misc/excel_copy_paste/Trigger formula Nifty BankNifty.xls"
trigger_formula_copy_path = "/home/alexis/projects/misc/excel_copy_paste/Nifty BankNifty Formula.xls"

nifty200_original_path = "/home/alexis/projects/misc/excel_copy_paste/Nifty200.xls"
autocopynifty200_copy_path = "/home/alexis/projects/misc/excel_copy_paste/AutoCopyNifty200.xls"

nse_cash_original_path = "/home/alexis/projects/misc/excel_copy_paste/NSE Cash Others.xls"
nse_cash_copy_path = "/home/alexis/projects/misc/excel_copy_paste/NSE Cash.xls"

def trigger_formula_nifty_copy(datetime_ob):
    trigger_formula_nifty = pd.read_excel(trigger_formula_original_path, sheet_name="Formula", header=67)
    
    copy_trigger_formula_nifty_path = trigger_formula_copy_path
    copy_trigger_formula_nifty = pd.read_excel(copy_trigger_formula_nifty_path, sheet_name="Sheet1", header=0)
    copy_trigger_formula_nifty[str(column_time(datetime_ob.hour,datetime_ob.minute))] = trigger_formula_nifty["Unnamed: 1"]
    copy_trigger_formula_nifty.to_excel(copy_trigger_formula_nifty_path, index=False, sheet_name="Sheet1")

def nifty200_copy(datetime_ob):
    nifty200 = pd.read_excel(nifty200_original_path, sheet_name="Sheet1", header=7)

    autocopynifty200_path = autocopynifty200_copy_path
    total_autocopynifty200 = pd.read_excel(autocopynifty200_path, sheet_name="Total Qty", header=0)
    buy_autocopynifty200 = pd.read_excel(autocopynifty200_path, sheet_name="Buy Qty", header=0)
    sell_autocopynifty200 = pd.read_excel(autocopynifty200_path, sheet_name="Sell Qty", header=0)
    
    with ExcelWriter(autocopynifty200_path) as writer:
        total_autocopynifty200[str(column_time(datetime_ob.hour,datetime_ob.minute))] = nifty200.iloc[:,19]
        buy_autocopynifty200[str(column_time(datetime_ob.hour,datetime_ob.minute))] = nifty200.iloc[:,20]
        sell_autocopynifty200[str(column_time(datetime_ob.hour,datetime_ob.minute))] = nifty200.iloc[:,21]

        total_autocopynifty200.to_excel(writer, index=False, sheet_name="Total Qty")
        buy_autocopynifty200.to_excel(writer, index=False, sheet_name="Buy Qty")
        sell_autocopynifty200.to_excel(writer, index=False, sheet_name="Sell Qty")
        writer.save()

def nse_cash_others_copy(datetime_ob):
    nse_cash = pd.read_excel(nse_cash_original_path, sheet_name="Sheet1", header=7)

    nse_cash_path = nse_cash_copy_path
    total_nse_cash = pd.read_excel(nse_cash_path, sheet_name="Total Qty", header=0)
    buy_nse_cash = pd.read_excel(nse_cash_path, sheet_name="Buy Qty", header=0)
    sell_nse_cash = pd.read_excel(nse_cash_path, sheet_name="Sell Qty", header=0)
    
    with ExcelWriter(nse_cash_path) as writer:
        total_nse_cash[str(column_time(datetime_ob.hour,datetime_ob.minute))] = nse_cash.iloc[:,19]
        buy_nse_cash[str(column_time(datetime_ob.hour,datetime_ob.minute))] = nse_cash.iloc[:,20]
        sell_nse_cash[str(column_time(datetime_ob.hour,datetime_ob.minute))] = nse_cash.iloc[:,21]

        total_nse_cash.to_excel(writer, index=False, sheet_name="Total Qty")
        buy_nse_cash.to_excel(writer, index=False, sheet_name="Buy Qty")
        sell_nse_cash.to_excel(writer, index=False, sheet_name="Sell Qty")
        writer.save()

def copy_stuff(datetime_ob):
    print("Copying data at "+str(column_time(datetime_ob.hour,datetime_ob.minute)))
    nifty200_copy(datetime_ob)
    nse_cash_others_copy(datetime_ob)
    trigger_formula_nifty_copy(datetime_ob)

print("Good Morning, starting the script... ")
current_time = datetime.now()
print("Current time is "+str(current_time)+" \n")
starting_time = datetime(datetime.now().year, datetime.now().month, datetime.now().day, startH, startM)
end_time = datetime(datetime.now().year, datetime.now().month, datetime.now().day, endH, endM)

while end_time>current_time:
    current_time = datetime.now()
    if datetime.now() < starting_time:
        time_to_start = starting_time - datetime.now()
        hours_to_go = time_to_start.seconds//3600
        minutes_to_go = (time_to_start.seconds%3600)//60
        secs_to_go = time_to_start.seconds%60
        print("Next copy is in "+str(hours_to_go)+"h "+str(minutes_to_go)+"m "+str(secs_to_go)+"s ")
        print()
        time.sleep(time_to_start.total_seconds())
    else:
        time_to_start = datetime.now() - starting_time
        hours_passed = time_to_start.seconds//3600
        minutes_passed = (time_to_start.seconds%3600)//60
        secs_passed = time_to_start.seconds%60
        if hours_passed==0 and minutes_passed==0:
            copy_stuff(starting_time)
            starting_time+=timedelta(0,time_between_copying)
            continue
        while starting_time < datetime.now():
            starting_time+=timedelta(0,time_between_copying)
            time_to_start = datetime.now() - starting_time
            hours_passed = time_to_start.seconds//3600
            minutes_passed = (time_to_start.seconds%3600)//60
            secs_passed = time_to_start.seconds%60
            if hours_passed==0 and minutes_passed==0:
                copy_stuff(starting_time)
                starting_time+=timedelta(0,time_between_copying)
                continue
