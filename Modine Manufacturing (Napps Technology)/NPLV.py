""" Ancel Carson
   Napps Technology Comporation
   21/12/2020
   IPLV.py: Calculates the IPLV of a unit using prompted user input
   Proof of concept for JESS
"""
#Libraries
import pandas as pd

#Variables
tonnage = None
frame = None
numFans = None
voltage = None
compType = None
compKW = [0,0,0,0]
fanKW = [0,0,0,0]
dbTemp = [95,80,65,55]
dbCorrect = 0
compQs = [0,0]
compCalc = [0,0,0]
waterTemp = [54,51.5,49,46.5]
IPLVTable = [[0,0,0,0,0],[0,0,0,0,0],[0,0,0,0,0],[0,0,0,0,0]]
percentLoad = [0,0,0,0]
calcTable = [[1,0,0],[.75,0,0],[.5,0,0],[.25,0,0],['Part Loads',0,0]]

#Functions
" Main Finction "
def main():
   unitInfo = startingQuestions()
   fansDf = fanTable()
   setFans(unitInfo, fansDf, fanKW)
   runCalc(unitInfo)
   '''IPLVTable = [[1, 95, 44.92576833333333, 52.365808909730376, 10.295061438453693], 
                [0.817340593803979, 84.04049235375719, 36.719654166666665, 35.80917350527552, 12.305110865937303], 
                [0.6023463297474304, 71.12045578978413, 27.060871666666667, 18.377157092614304, 17.67033161677155], 
                [0.31311664452705895, 55, 14.067005833333335, 9.103522860492376, 18.54271940509743]]'''
   IPLVCalc(IPLVTable, unitInfo)
   printResults(calcTable, IPLVTable, unitInfo)

" Questions to set up slelction "
" These can be pulled from inside JESS "
def startingQuestions(): 
   tonnage = int(input("What is this Unit's Tonnage?\n"))
   frame = input('What size frame?\n')
   numFans = int(input('How many fans?\n'))
   voltage = int(input('What voltage is the unit?\n'))
   compType = input('Singles or tandems?\n')
   compSize = input('What size are the compressors?\n')
   return [tonnage, frame, numFans, voltage, compType, compSize]
   '''return [50, 'green', 2, 460, 'tandems', 'DSH 140']'''

" Lookup Table for Fan Power "
def fanTable():
   fanDf = pd.DataFrame({'ACCSize':['orange','orange','green','green','green','green','blue','blue','blue','blue','blue','blue'],
                             'Volt':[208,460,208,208,460,460,208,208,208,460,460,460],
                             'Fans':[1,1,1,2,1,2,2,3,4,2,3,4],
                             'Fan1':[1.68,2.53,1.68,1.76,2.53,2.78,1.68,1.76,1.76,2.53,2.78,2.78],
                             'Fan2':[0,0,0,0,0,0,0,1.76,1.76,0,2.78,2.78],
                             'Fan3':[0,0,0,1.76,0,2.78,1.68,1.68,1.76,2.53,2.53,2.78],
                             'Fan4':[0,0,0,0,0,0,0,0,1.76,0,0,2.78]},
                        columns = ['ACCSize','Volt','Fans','Fan1','Fan2','Fan3','Fan4'])
   return fanDf

" Sets the Fan Power to an array from the fanTable "
def setFans(unitInfo, fansDf, fanKW):
   queryDf = fansDf.query('ACCSize == "{}" and Volt == {} and Fans == {}'.format(unitInfo[1],unitInfo[3],unitInfo[2])).reset_index()
   fanKW[0] = queryDf.iloc[0]['Fan1']
   fanKW[1] = queryDf.iloc[0]['Fan2']
   fanKW[2] = queryDf.iloc[0]['Fan3']
   fanKW[3] = queryDf.iloc[0]['Fan4']

" Runs the calculations and describes which pass is being run "
def runCalc(unitInfo):
   pastCap = 0
   print(unitInfo)
   stepRange = [0,1,2,3]
   if unitInfo[4] == 'singles':
      stepRange = [0,2]
   for step in stepRange:
      if unitInfo[4] == "tandems":
         if step == 0:
            print("Set both circuits to {} tandem".format(unitInfo[5]))
         elif step == 1:
            print("Set circuit 1 to {} tandem and circuit 2 to {}".format(unitInfo[5], unitInfo[5]))
         elif step == 2:
            print("Set both circuits to {}".format(unitInfo[5]))
         else:
            print("Set circuit 1 to {} tandem and set circuit 2 to 0".format(unitInfo[5]))
      if unitInfo[4] == "singles":
         if step == 0:
            print("Set both circuits to {}".format(unitInfo[5]))
         else:
            print("Set circuit 1 to {} tandem and set circuit 2 to 0".format(unitInfo[5]))
      print('Set the EWT to {}°F'.format(waterTemp[step]))
      converge(dbTemp, step, pastCap)
      fillTable(step)
      if step == 0:
         pastCap = compCalc[0]

" Converting mBh to Tons "
def tonCalc(compQs, compCalc):
   compCalc[0] = compQs[0]/12000
   compCalc[1] = compQs[1]/12000
   compCalc[2] = (compQs[1]-compQs[0])/3412

" Itterates to determine correct db Temp "
def converge(dbTemp, step, pastCap):
   print('Run the selection with {0:.3f}°F db and {1:.3f}°F wb'.format(dbTemp[step], dbTemp[step]-8))
   compQs[0] = float(input("What is the Qactual of Circuit 1?\n"))
   compQs[1] = float(input("What is the QCond of Circuit 1?\n"))
   compQs[0] += float(input("What is the Qactual of Circuit 2?\n'0' if no circuit 2\n"))
   compQs[1] += float(input("What is the QCond of Circuit 2?\n'0' if no circuit 2\n"))
   tonCalc(compQs, compCalc)
   if step == 0:
      percentLoad[step] = 1
   elif step == 3:
      percentLoad[step] = compCalc[0]/pastCap
   else:
      percentLoad[step] = compCalc[0]/pastCap
      dbCorrect = (percentLoad[step]*100)*0.6+35
      if abs(dbTemp[step] - dbCorrect) <= .1:
         print('Convergence Complete')
      else:
         print('Convergence Incomplete. Itterating...')
         dbTemp[step] += (dbCorrect-dbTemp[step])*.75
         converge(dbTemp, step, pastCap)

" Fills the IPLV Table with information from the Converge "
def fillTable(step):
   IPLVTable[step][0] = percentLoad[step]
   IPLVTable[step][1] = dbTemp[step]
   IPLVTable[step][2] = compCalc[0]
   IPLVTable[step][3] = compCalc[2]+.5
   for num in range(4-step):
      IPLVTable[step][3] += fanKW[num]
   IPLVTable[step][4] = 12/(IPLVTable[step][3]/IPLVTable[step][2])

" Calculates IPLV values from the raw data "
def IPLVCalc(IPLVTable, unitInfo):
   partLoadMults  =[.01,.42,.45,.12]
   calcTable[0][2] = IPLVTable[0][4]
   calcTable[0][1] = 12/calcTable[0][2]
   if unitInfo[4] == "tandems":
      for row in [1,2]:
         calcTable[row][2] = (((calcTable[row][0]-IPLVTable[row+1][0])/
                               (IPLVTable[row][0]-IPLVTable[row+1][0]))*
                               (IPLVTable[row][4]-IPLVTable[row+1][4])+
                               IPLVTable[row+1][4])
         calcTable[row][1] = 12/calcTable[row][2]
      Cd = (-.13*(calcTable[3][0]*IPLVTable[0][2])/IPLVTable[3][2])+1.13
      calcTable[3][2] = IPLVTable[3][4]/Cd
      calcTable[3][1] = 12/calcTable[3][2]
   elif unitInfo[4] == "singles":
      Cd = (-.13*(calcTable[2][0]*IPLVTable[0][2])/IPLVTable[2][2])+1.13
      calcTable[2][2] = IPLVTable[2][4]/Cd
      Cd = (-.13*(calcTable[3][0]*IPLVTable[0][2])/IPLVTable[2][2])+1.13
      calcTable[3][2] = IPLVTable[2][4]/Cd
      calcTable[1][2] = (((calcTable[1][0]-IPLVTable[2][0])/
                            (IPLVTable[0][0]-IPLVTable[2][0]))*
                            (IPLVTable[0][4]-IPLVTable[2][4])+
                            IPLVTable[2][4])
      for num in range(1,4):
         calcTable[num][1] = 12/calcTable[num][2]
   for num in range(4):
      calcTable[4][1] += partLoadMults[num]*calcTable[num][2]
      calcTable[4][2] += partLoadMults[num]/calcTable[num][1]
   calcTable[4][2] = 1/calcTable[4][2]

" FOrmats output into a readable form "
def printResults(calcTable, IPLVTable, unitInfo):
   print("Results for {} Ton {} Volt Unit\n".format(unitInfo[0],unitInfo[3]))
   print('Predicted Performance:\n\t% Load\tEDB\tCap.\tPower\tEfficiency')
   for row in IPLVTable:
      print("\t{0:.1f}%\t{1:.1f}\t{2:.2f}\t{3:.2f}\t{4:.2f}".format(row[0]*100,row[1],row[2],row[3],row[4]))
   print('\nIPLV:\n\tLoad Point\tKW/Tons\t\tEER (Btu/W*h)')
   for row in calcTable:
      if type(row[0]) == str:
         break
      print('\t{0}\t\t{1:.3f}\t\t{2:.2f}'.format(row[0]*100,row[1],row[2]))
   print('\nIntegrated Part Load Value (EER): {0:.2f}'.format(calcTable[4][1]))
   print('Integrated Part Load Value (KW/Ton): {0:.2f}\n'.format(calcTable[4][2]))
   input("Press 'Enter' to close the program...")

" Checks if this program is beiong called "
if __name__ == "__main__":
   main()