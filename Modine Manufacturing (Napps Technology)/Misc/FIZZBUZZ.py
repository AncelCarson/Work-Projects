# Author: Ancel Carson
# Orginization: Napps Technology Comporation
# Creation Date: 7/1/2021
# Update Date: 7/1/2021
# FIZZBUZZ.py

"""A one line summary of the module or program, terminated by a period.

Leave one blank line.  The rest of this docstring should contain an
overall description of the module or program.  Optionally, it may also
contain a brief description of exported classes and functions and/or usage
examples.

Functions:
   main: Driver of the program
"""
#Libraries

#Variables

#Functions
" Main Finction "
def main():
   run = True
   count = 0
   while(run):
      count += 1
      out = ""
      blank = True
      if count % 3 == 0:
         out = out + "Fizz"
         blank = False
      if count % 5 == 0:
         out = out + "Buzz"
         blank = False
      if blank == True:
         out = out + str(count)
      print(out)
      if count == 100:
         run = False

" Checks if this program is beiong called "
if __name__ == "__main__":
   main()