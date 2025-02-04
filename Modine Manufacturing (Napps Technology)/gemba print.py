#pylint: disable = invalid-name,bad-indentation,non-ascii-name
#-*- coding: utf-8 -*-

import os
import glob
from dotenv import load_dotenv

#Secret Variables
load_dotenv()
Shared_Drive = os.getenv('Shared_Drive')

list_of_files = glob.glob(fr'\\{Shared_Drive}\GEMBA Board\*.xlsx') # * means all if need specific format then *.csv
latest_file = max(list_of_files, key=os.path.getctime)
print(latest_file)
#os.startfile(latest_file,'read')
os.startfile(latest_file,'print')
