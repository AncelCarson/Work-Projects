import glob
import os


list_of_files = glob.glob(r'S:\GEMBA Board\*.xlsx') # * means all if need specific format then *.csv
latest_file = max(list_of_files, key=os.path.getctime)
print(latest_file)
#os.startfile(latest_file,'read')
os.startfile(latest_file,'print')
