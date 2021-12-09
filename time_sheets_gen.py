import random
import datetime
import pandas as pd
import locale
import calendar
import os


BASE_HOUR = 9           # 7 to 9


# names of months in spanish
locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')

# open source files
parameters = open("parameters.txt", "r")
template = open("template.svg", "r")
template_str = template.read()

# read parameters
text = parameters.readline()
marker = text.find("NAME: ")
if(marker != -1):
    tname = (text.split("NAME: "))[1].split("\n")[0]
else:
    print("NAME error, check parameters.txt")

text = parameters.readline()
marker = text.find("DNI: ")
if(marker != -1):
    tdni = (text.split("DNI: "))[1].split("\n")[0]
else:
    print("DNI error, check parameters.txt")

text = parameters.readline()
marker = text.find("SHEETS: ")
if(marker != -1):
    months_range = (text.split("SHEETS: "))[1].split("\n")[0]
    first_month =  datetime.datetime.strptime(months_range.split("-")[0], "%m/%Y")
    last_month = datetime.datetime.strptime(months_range.split("-")[1], "%m/%Y")
    months_list = pd.date_range(first_month, last_month+pd.offsets.MonthEnd(), freq='M').strftime("%B/%Y").tolist()
else:
    print("SHEET error, check parameters.txt")

text = parameters.readline()
marker = text.find("HOLIDAYS: ")
if(marker != -1):
    holidays_str = (text.split("HOLIDAYS: "))[1].split("\n")[0]
    holidays_str = holidays_str.split(",")
    holidays = []
    for h_date in holidays_str:
        if h_date.find("-") != -1:
            first_day = datetime.datetime.strptime(h_date.split("-")[0], "%d/%m/%Y")
            last_day = datetime.datetime.strptime(h_date.split("-")[1], "%d/%m/%Y")
            holidays += ([first_day + datetime.timedelta(days=x) for x in range((last_day-first_day).days + 1)])
        else:
            single_day = datetime.datetime.strptime(h_date, "%d/%m/%Y")
            holidays.append(single_day)
else:
    print("HOLIDAYS error, check parameters.txt")

text = parameters.readline()
marker = text.find("PUBLIC HOLIDAYS: ")
if(marker != -1):
    public_holidays_str = (text.split("PUBLIC HOLIDAYS: "))[1].split("\n")[0]
    public_holidays_str = public_holidays_str.split(",")
    public_holidays = []
    for h_date in public_holidays_str:
        if h_date.find("-") != -1:
            first_day = datetime.datetime.strptime(h_date.split("-")[0], "%d/%m/%Y")
            last_day = datetime.datetime.strptime(h_date.split("-")[1], "%d/%m/%Y")
            public_holidays += ([first_day + datetime.timedelta(days=x) for x in range((last_day-first_day).days + 1)])
        else:
            single_day = datetime.datetime.strptime(h_date, "%d/%m/%Y")
            public_holidays.append(single_day)
else:
    print("PUBLIC HOLIDAYS error, check parameters.txt")

text = parameters.readline()
marker = text.find("OTHER DAYS OFF: ")
if(marker != -1):
    other_days_off_str = (text.split("OTHER DAYS OFF: "))[1].split("\n")[0]
    other_days_off_str = other_days_off_str.split(",")
    other_days_off = []
    for h_date in other_days_off_str:
        if h_date.find("-") != -1:
            first_day = datetime.datetime.strptime(h_date.split("-")[0], "%d/%m/%Y")
            last_day = datetime.datetime.strptime(h_date.split("-")[1], "%d/%m/%Y")
            other_days_off += ([first_day + datetime.timedelta(days=x) for x in range((last_day-first_day).days + 1)])
        else:
            single_day = datetime.datetime.strptime(h_date, "%d/%m/%Y")
            other_days_off.append(single_day)
else:
    print("OTHER error, check parameters.txt")


for month in months_list:
    month_str = month.split("/")[0].encode('ascii','ignore')
    year_str = month.split("/")[1].encode('ascii','ignore')
    month_number = list(calendar.month_name).index(month_str)
    year = int(year_str)
    n_days = calendar.monthrange(int(year_str), month_number)[1]

    print "Generating " + month_str + " " + year_str + " sheet..."

    sheet = template_str

    # Dates and personal data
    sheet = sheet.replace("#TNAME#", tname)
    sheet = sheet.replace("#TDNI#", tdni)    
    sheet = sheet.replace("#TMONTH#", month_str.upper())
    sheet = sheet.replace("#TYEAR#", year_str)
    sheet = sheet.replace("#FDAY#", str(n_days))
    sheet = sheet.replace("#FMONTH#", month_str)
    sheet = sheet.replace("#FYEAR#", year_str)

    # Hour register
    htotal = 0
    for i in range(n_days):
        date = datetime.datetime(int(year_str),month_number, i+1)
        if date.weekday() > 4: # date is part of weekend
            marker = sheet.find("A:"+str(i+1).zfill(2))
            sheet = sheet.replace(sheet[marker: marker+59], " "*59)
        elif date in public_holidays: # date is part of public holidays
            marker = sheet.find("A:"+str(i+1).zfill(2))
            sheet = sheet.replace(sheet[marker: marker+59], "FESTIVO   " + " "*49)
        elif date in other_days_off: # date is part of other days off
            marker = sheet.find("A:"+str(i+1).zfill(2))
            sheet = sheet.replace(sheet[marker: marker+59], "OTROS     " + " "*49)
        elif date in holidays: # date is part of holidays
            marker = sheet.find("A:"+str(i+1).zfill(2))
            sheet = sheet.replace(sheet[marker: marker+59], "VACACIONES" + " "*49)
        else: # standard working day
            minutes = random.randint(0,30)
            if date.weekday() == 4 or (datetime.datetime(year, 6, 1) <= date <= (datetime.datetime(year, 9, 14))): # friday
                daily_hours = 7
                daily_hours_str = " 7  "
            else:
                daily_hours = 8.5
                daily_hours_str = str(daily_hours)
            
            sheet = sheet.replace("A:"+str(i+1).zfill(2), str(BASE_HOUR)+":"+str(minutes).zfill(2))
            init_time = datetime.datetime(1,1,1,BASE_HOUR,minutes)
            finish_time = init_time + datetime.timedelta(hours=daily_hours+1)
            finish_time_str = str(finish_time.hour) + ":" + str(finish_time.minute).zfill(2)
            sheet = sheet.replace("BB:"+str(i+1).zfill(2), finish_time_str)
            sheet = sheet.replace("C"+str(i+1).zfill(2), daily_hours_str)
            htotal += daily_hours

    for i in range(n_days, 31): # non existent days
        marker = sheet.find("A:"+str(i+1).zfill(2))
        sheet = sheet.replace(sheet[marker: marker+59], " "*59)
    
    sheet = sheet.replace("#HTOTAL#", str(htotal))


    filename = "output/" + month_str.upper() + "_" + year_str
    new_file = open(filename + ".svg", "w")
    new_file.write(sheet)
    new_file.close()
    os.system("inkscape " + filename + ".svg --export-pdf=" + filename + ".pdf")
    os.system("rm " + filename + ".svg")

exit()

