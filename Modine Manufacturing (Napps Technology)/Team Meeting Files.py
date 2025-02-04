#pylint: disable = invalid-name,bad-indentation,non-ascii-name
#-*- coding: utf-8 -*-

# Author: Ancel Carson
# Orginization: Napps Technology Comporation
# Creation Date: 25/9/2023
# Update Date: 4/2/2025
# Team Meeting Files.py: Opens files needed for the Monday Team Meetings

#Libraries
import os
import glob
from dotenv import load_dotenv

#Secret Variables
load_dotenv()
Shared_Drive = os.getenv('Shared_Drive')

#Variables

#Functions
" Main Finction "
def main():
   openProjects()
   openGemba()
   openScoreCard()

def openScoreCard():
   list_of_files = glob.glob(fr'\\{Shared_Drive}\Scorecard\Scorecard *.xlsx') # * means all if need specific format then *.csv
   latest_file = max(list_of_files, key=os.path.getctime)
   os.startfile(latest_file)

def openGemba():
   list_of_files = glob.glob(fr'\\{Shared_Drive}\GEMBA Board\*.xlsx') # * means all if need specific format then *.csv
   latest_file = max(list_of_files, key=os.path.getctime)
   os.startfile(latest_file)

def openProjects():
   list_of_files = glob.glob(fr'\\{Shared_Drive}\WOPO Board\Work Order Project List *.xlsx') # * means all if need specific format then *.csv
   latest_file = max(list_of_files, key=os.path.getctime)
   os.startfile(latest_file)

" Checks if this program is beiong called "
if __name__ == "__main__":
   main()