# Author: Ancel Carson
# Orginization: Napps Technology Comporation
# Creation Date: 12/10/2021
# Update Date: 12/10/2021
# JetsonChime.py

"""A one line summary of the module or program, terminated by a period.

Leave one blank line.  The rest of this docstring should contain an
overall description of the module or program.  Optionally, it may also
contain a brief description of exported classes and functions and/or usage
examples.

Functions:
   main: Driver of the program
"""
#Libraries
from playsound import playsound

#Variables
wavChime = r'S:\\Ancel\\Python_Modules\\AdIns\\Resources\\jetson.wav'

#Functions
" Main Finction "
def chime():
   playsound(wavChime)

" Checks if this program is beiong called "
if __name__ == "__main__":
   chime()