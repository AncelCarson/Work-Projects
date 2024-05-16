# Author: Ancel Carson
# Orginization: Napps Technology Comporation
# Creation Date: 13/2/2024
# Update Date: 16/2/2024
# WeekendReport.py

"""This Program compiles a weeks worth of logs and generates a report.

Leave one blank line.  The rest of this docstring should contain an
overall description of the module or program.  Optionally, it may also
contain a brief description of exported classes and functions and/or usage
examples.

Functions:
   main: Driver of the program
"""
#Libraries
import os
import sys
import glob
import pandas as pd
from datetime import datetime, date
from dateutil.relativedelta import relativedelta, MO, SU

#custom Modules
sys.path.insert(0,r'S:\Programs\Add_ins')

#Variables
# day_set = 1 ## Used when Running Past Weeks
input_folder = r"U:\Daily Log\*"
output_folder = r"U:\Weekly Log"
today = date.today()
# today = date.today() + relativedelta(weekday=SU(day_set * -1)) ## Used when Running Past Weeks
file_day = today.strftime('%y%m%d')
last_monday = today + relativedelta(weekday=MO(-1))
week_days = (today - last_monday).days

#Functions
" Main Finction "
def main():
   count = 0
   week_files = []
   week_tasks = []
   week_stats = [[],[],[],[]]
   lunch = []
   wfh = []
   list_of_files = glob.glob(input_folder) # * means all if need specific format then *.txt
   list_of_files.sort(key=os.path.getctime, reverse=True)
   for file in list_of_files:
      if count > 7:
         break
      created_day = datetime.fromtimestamp(os.path.getctime(file)).date()
      if created_day <= today and created_day >= last_monday:
         week_files.append(file)
         count += 1
   for file in week_files:
      day_list = get_day_list(file)
      day_review = review_day(day_list) # tasks, start, end, length, switch
      week_stats[0].append(day_review[1])
      week_stats[1].append(day_review[2])
      week_stats[2].append(day_review[3])
      week_stats[3].append(day_review[4])
      lunch.append(day_review[5])
      wfh.append(day_review[6])
      for task in day_review[0]:
         week_tasks.append(task)
   day_count = len(week_stats[0])
   days_home = sum(wfh)
   week_length = sum(week_stats[2])
   start_min = min(week_stats[0])
   end_max = max(week_stats[1])
   lunch_average = sum(lunch)/len(lunch)
   week_review = review_week(week_tasks, week_stats)

   file = output_folder + "\\Week of " + file_day + " Log.csv"
   week_review[0].to_csv(file, sep=',', mode='a')
   with open(file,"a") as f:
      f.write("\nLength of Week: {0} Days\n".format(day_count))
      f.write("Days worked from Home: {0}\n".format(days_home))
      f.write("Total Hours: {:0.2f}\n".format(week_length))
      f.write("Average Start Time: {0}\n".format(week_review[1]))
      f.write("Average End Time: {0}\n".format(week_review[2]))
      f.write("Earliest Morning: {0}\n".format(start_min))
      f.write("Latest Night: {0}\n".format(end_max))
      f.write("Average Lunch Break: {:0.2f} Hours\n".format(lunch_average))
      f.write("Average Length of Day: {:0.2f} Hours\n".format(week_review[3]))
      f.write("Average Daily Tasks Changes: {:0.0f}\n".format(week_review[4]))

def get_day_list(file):
   lines = None
   with open(file) as f:
      lines = f.readlines()
   del lines[0]
   for num in range(len(lines)):
      lines[num] = lines[num].split(": ")
   return lines

def review_day(day_list):
   day_list[0][1] = "Misc\n"
   start_time = day_list[0][0]
   end_time = None
   item_list = []
   lunch = 0
   ooo = 0
   wfh = 0
   for num in range(len(day_list)):
      if day_list[num][1] == "End Day" or day_list[num][1] == "End Week":
         end_time = day_list[num][0]
         break
      day_list[num][1] = day_list[num][1][:-1]
      current_time = datetime.strptime(day_list[num][0], '%H:%M')
      next_time = datetime.strptime(day_list[num+1][0], '%H:%M')
      day_list[num][0] = (next_time - current_time).total_seconds()/3600
      if "/" in day_list[num][1]:
         split_task = day_list[num][1].split("/")
         for task in split_task:
            if task == "Lunch" or task == "Lunch From Home":
               lunch += day_list[num][0]/len(split_task)
               continue
            if task == "Out of Office" or task == "OOO":
               ooo += day_list[num][0]/len(split_task)
            item_list.append([day_list[num][0]/len(split_task),task])
      else:
         if day_list[num][1] == "Lunch" or day_list[num][1] == "Lunch From Home":
            lunch += day_list[num][0]
            continue
         if day_list[num][1] == "Out of Office" or day_list[num][1] == "OOO":
            ooo += day_list[num][0]
            continue
         if day_list[num][1] == "Work From Home" or day_list[num][1] == "WFH":
            wfh += 1
            day_list[num][1] = "Misc"
            continue
         item_list.append(day_list[num])
   task_change = len(item_list)

   day_start = datetime.strptime(start_time, '%H:%M')
   day_end = datetime.strptime(end_time, '%H:%M')
   day_length = (day_end - day_start).total_seconds()/3600 - lunch - ooo
   return [item_list, start_time, end_time, day_length, task_change, lunch, wfh]

def review_week(week_tasks, week_stats):
   task_df = pd.DataFrame(week_tasks, columns = ["Hours", "Task"])
   task_df = task_df.groupby(by=["Task"]).sum()
   task_df.sort_values(by=['Hours'],inplace=True)
   task_df = task_df.reindex(index=task_df.index[::-1])

   times_lists = week_stats[:2]
   counts_list = week_stats[2:]

   num = 0
   for times in times_lists:
      times = times[:5]
      average_time=pd.to_datetime(pd.Series(times), format="%H:%M").mean().strftime('%H:%M')
      week_stats[num] = average_time
      num += 1

   for counts in counts_list:
      average_count = (sum(counts)/len(counts))
      week_stats[num] = average_count
      num += 1

   task_df["Percentage"] = task_df.apply(lambda x: (x["Hours"]*100)/sum(counts_list[0]), axis = 1)
   task_df["Percentage"] = task_df["Percentage"].round(2)
   task_df["Hours"] = task_df["Hours"].round(2)
   return [task_df, week_stats[0], week_stats[1], week_stats[2], week_stats[3]]


" Checks if this program is being called "
if __name__ == "__main__":
   main()