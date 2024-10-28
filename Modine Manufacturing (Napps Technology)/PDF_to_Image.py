# Author: Ancel Carson
# Orginization: Napps Technology Comporation
# Creation Date: 3/2/2024
# Update Date: 3/2/2024
# PDF_to_Image.py

"""A one line summary of the module or program, terminated by a period.

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
import fitz

#custom Modules
sys.path.insert(0,r'S:\Programs\Add_ins')

#Variables
source_path = r'U:\JESS Drawings\PDFs\*.pdf' # * means all if need specific format then *.PDF
dest_path = r'U:\JESS Drawings\PNGs'

#Functions
" Main Finction "
def main():
   files = glob.glob(source_path)
   for file in files:
      doc = fitz.open(file)
      name = file.split("\\")[3].split(".")[0]
      
      for _, page in enumerate(doc):
         pix = page.get_pixmap()  # render page to an image
         pix.save(dest_path + "\\" + name + ".png")
      # image.save()
      print(name + " Saved as .PNG")
   print("Program Complete")
   input("Press ENTER to close window...")

" Checks if this program is being called "
if __name__ == "__main__":
   main()