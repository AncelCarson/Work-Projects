# Author: Ancel Carson
# Orginization: Napps Technology Comporation
# Creation Date: 20/6/2022
# Update Date: 20/6/2022
# MenuMaker.py
# Rev 1

"""Variable Size Menu Creater.

Program receives a list of items consisting of a title and options. It will 
then create items that are sized to havwe a uniform menu.


Functions:
   main: Driver of the program
"""

#Functions
" Main Finction "
def menu(MenuItems):
   title = MenuItems[0]
   length = maxLength(MenuItems)
   if(len(title) == length):
      title = " " + title
      length = length + 1

   diff = length - len(title) + 1
   cSpace = " " * int(diff / int(2))
   title = cSpace + title + cSpace
   if (diff%2 != 0):
      title = title + " "

   spacer = "-" * (length + 1)
   MenuItems[0] = spacer
   MenuItems.append(spacer)
   wrapping = lambda item: "|" + item + "|"
   print("\n" + wrapping(title))
   for item in MenuItems:
      option = item + (" " * (length - len(item) + 1))
      print(wrapping(option))

def makeMenu(title, list):
   menuList = [title]
   count = 0
   for item in list:
      count += 1
      menuList.append(str(str(count) + ": " + item))
   menu(menuList)

def maxLength(lst):
   maxLength = max(len(x) for x in lst)
   return maxLength


" Checks if this program is beiong called "
if __name__ == "__main__":
   menu(["Fruit","A: Apple","B: Orange","C: Pear","D: Grape","E: Pomegranite"])