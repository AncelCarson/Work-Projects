# Author: Ancel Carson
# Orginization: Napps Technology Comporation
# Creation Date: 21/12/2020
# Update Date: 8/7/2021
# IPLVCalculator

"""A calculator for the IPLV of any given ACCX unit.

Calculates the IPLV of a unit using prompted user input. Given set inputs,
will itterate values to calculate effieicnse at 4 steps of capacity.
Proof of concept for JESS

Functions:
   main: Driver of the program
   startingQuestions: Collects Unit information
   fanTable: Creates dataframe for fan power
   setFans: Selects fan power based off unit information
   runCalc: Runs the calculation and drives the convergence
   tonCalc: Converts BTUs to Tons for Qactual and Qcond
   converge: Itterates to to find the condenser relief point
   fillTable: Fills the IPLV Table with converged data
   IPLVCalc: Calculated the IPLVs for each step of capacity
   printResults: Outputs the results in a readable format
"""

# Libraries
import pandas as pd

# Functions
def main():
   """Driver of the program.

   Called from when the program is run standalone. Calls all
   other methods.

   Variables:
      waterTemp: list[int]: Water Temperature set points at each step of capacity.
      dbTemp: list[int]: Starting db temperatures .
      unitInfo: list[int, str]: List of information pertaining to the unit being run.
      fansDf: DataFrame of fan power requirements.
      fanKW: list[int]: Fan power requirements for the unit described by unitInfo.
      IPLVTable: list[int]: List of values collected from convergence().
      calcTable: list[int]: List of IPLV values calculated from IPLVTable.
   """

   waterTemp = [54,51.5,49,46.5]
   dbTemp = [95,80,65,55]
   unitInfo = startingQuestions()
   fansDf = fanTable()
   fanKW = setFans(unitInfo, fansDf)
   IPLVTable = runCalc(unitInfo, fanKW, waterTemp, dbTemp)
   calcTable = IPLVCalc(IPLVTable, unitInfo)
   printResults(calcTable, IPLVTable, unitInfo)

def startingQuestions() -> list: 
   """Collects data to set up slelction.

   Asks a series of questions to gather information on the unit being run.
   This data can be pulled from a selection in JESS.

   Returns:
      A list of integers and strings that describes the unit for which the 
      selection is being run.

      [int, str, int, int, str, str]

      The values follow the order of the questions being asked.
   """

   tonnage = tonMenu()
   frame = frameMenu()
   numFans = fanMenu()
   voltage = voltMenu()
   compType = typeMenu()
   compOne = compMenu()
   if tonnage == 70:
      print("\nPlease select compressors for circuit 2")
      compTwo = compMenu()
   else:
      compTwo = compOne
   return [tonnage, frame, numFans, voltage, compType, compOne, compTwo]

def tonMenu():
   print("\n| Unit Tonage |")
   print("|-------------|")
   print("| A: 10 Tons  |")
   print("| B: 15 Tons  |")
   print("| C: 20 Tons  |")
   print("| D: 25 Tons  |")
   print("| E: 30 Tons  |")
   print("| F: 40 Tons  |")
   print("| G: 50 Tons  |")
   print("| H: 60 Tons  |")
   print("| I: 70 Tons  |")
   print("| J: 80 Tons  |")
   print("|-------------|")
   selection = input("What is this Unit's Tonnage?\n")
   if selection == 'a' or selection == 'A':
      tonnage = 10
   elif selection == 'b' or selection == 'B':
      tonnage = 15
   elif selection == 'c' or selection == 'C':
      tonnage = 20
   elif selection == 'd' or selection == 'D':
      tonnage = 25
   elif selection == 'e' or selection == 'E':
      tonnage = 30
   elif selection == 'f' or selection == 'F':
      tonnage = 40
   elif selection == 'g' or selection == 'G':
      tonnage = 50
   elif selection == 'h' or selection == 'H':
      tonnage = 60
   elif selection == 'i' or selection == 'I':
      tonnage = 70
   elif selection == 'j' or selection == 'J':
      tonnage = 80
   else:
      print("The given answer was not one of the Options")
      print("Please try again")
      tonnage = tonMenu()
   return tonnage

def frameMenu():
   print("\n| Frame Size |")
   print("|------------|")
   print("| O: Orange  |")
   print("| G: Green   |")
   print("| B: Blue    |")
   print("|------------|")
   selection = input('What size frame?\n')
   if selection == 'o' or selection == 'O':
      frame = "orange"
   elif selection == 'g' or selection == 'G':
      frame = "green"
   elif selection == 'b' or selection == 'B':
      frame = "blue"
   else:
      print("The given answer was not one of the Options")
      print("Please try again")
      frame = frameMenu()
   return frame

def fanMenu():
   print("\n| Fan Count |")
   print("|-----------|")
   print("| A: 1 Fan  |")
   print("| B: 2 Fans |")
   print("| C: 4 Fans |")
   print("|-----------|")
   selection = input('How many fans?\n')
   if selection == 'a' or selection == 'A':
      numFans = 1
   elif selection == 'b' or selection == 'B':
      numFans = 2
   elif selection == 'c' or selection == 'C':
      numFans = 4
   else:
      print("The given answer was not one of the Options")
      print("Please try again")
      numFans = fanMenu()
   return numFans

def voltMenu():
   print("\n| Unit Voltage |")
   print("|--------------|")
   print("| A: 208V      |")
   print("| B: 230V      |")
   print("| F: 460V      |")
   print("|--------------|")
   selection = input('What voltage is the unit?\n')
   if selection == 'a' or selection == 'A':
      voltage = 208
   elif selection == 'b' or selection == 'B':
      voltage = 208
   elif selection == 'f' or selection == 'F':
      voltage = 460
   else:
      print("The given answer was not one of the Options")
      print("Please try again")
      voltage = voltMenu()
   return voltage

def typeMenu():
   print("\n| Compressor Type |")
   print("|-----------------|")
   print("| S: Singles      |")
   print("| T: Tandems      |")
   print("|-----------------|")
   selection = input('What type of compressors are on the unit?\n')
   if selection == 's' or selection == 'S':
      compType = "singles"
   elif selection == 't' or selection == 'T':
      compType = "tandems"
   else:
      print("The given answer was not one of the Options")
      print("Please try again")
      compType = typeMenu()
   return compType

def compMenu():
   print("\n| Compressor Size |")
   print("|-----------------|")
   print("| 1: DSH-090      |")
   print("| 2: DSH-120      |")
   print("| 3: DSH-140      |")
   print("| 4: DSH-184      |")
   print("| 5: DSH-240      |")
   print("| 6: DSH-295      |")
   print("| 7: DSH-381      |")
   print("| 8: DSH-485      |")
   print("| 9: HLH-061      |")
   print("|-----------------|")
   selection = int(input('What size are the compressors?\n'))
   if selection == 1:
      compSize = "DSH-090"
   elif selection == 2:
      compSize = "DSH-120"
   elif selection == 3:
      compSize = "DSH-140"
   elif selection == 4:
      compSize = "DSH-184"
   elif selection == 5:
      compSize = "DSH-240"
   elif selection == 6:
      compSize = "DSH-295"
   elif selection == 7:
      compSize = "DSH-381"
   elif selection == 8:
      compSize = "DSH-485"
   elif selection == 9:
      compSize = "HLH-061"
   else:
      print("The given answer was not one of the Options")
      print("Please try again")
      compSize = compMenu()
   return compSize

def fanTable():
   """Lookup Table for Fan Power.

   Creates a Pandas dataframe containing different fan voltages.
   Can be queried to get the row of values associated with a 
   specific unit. Dataframe contains values for all possible 
   combinations of frame size, voltage, and # of fans.

   Returns:
      fanDf: Dataframe containing all fan power options.
   """
   
   fanDf = pd.DataFrame({'ACCSize':['orange','orange','green','green','green','green','blue','blue','blue','blue','blue','blue'],
                             'Volt':[208,460,208,208,460,460,208,208,208,460,460,460],
                             'Fans':[1,1,1,2,1,2,2,3,4,2,3,4],
                             'Fan1':[1.68,2.53,1.68,1.76,2.53,2.78,1.68,1.76,1.76,2.53,2.78,2.78],
                             'Fan2':[0,0,0,0,0,0,0,1.76,1.76,0,2.78,2.78],
                             'Fan3':[0,0,0,1.76,0,2.78,1.68,1.68,1.76,2.53,2.53,2.78],
                             'Fan4':[0,0,0,0,0,0,0,0,1.76,0,0,2.78]},
                        columns = ['ACCSize','Volt','Fans','Fan1','Fan2','Fan3','Fan4'])
   return fanDf

def setFans(unitInfo:list, fansDf:dict) -> list:
   """Sets the Fan Power to an array from the fanTable.

   Selects the fans powers pased of the unit's description. Using
   frame size, voltage, and # of fans.

   Args:
      unitInfo: Used in query to select fan powers.
      fansDf: Dataframe from which the fan powers are pulled.

   Returns:
      fanKW: A list of values coresponding the power of each fan on the unit.
   """
   
   fanKW = [0,0,0,0]
   queryDf = fansDf.query('ACCSize == "{}" and Volt == {} and Fans == {}'.format(unitInfo[1],unitInfo[3],unitInfo[2])).reset_index()
   fanKW[0] = queryDf.iloc[0]['Fan1']
   fanKW[1] = queryDf.iloc[0]['Fan2']
   fanKW[2] = queryDf.iloc[0]['Fan3']
   fanKW[3] = queryDf.iloc[0]['Fan4']
   return fanKW

def runCalc(unitInfo:list, fanKW:list, waterTemp:list, dbTemp:int) -> list:
   """Runs the calculations and describes which pass is being run.

   Controls the convergence for the unit and describes which step
   is being run. Thje user is given the description of the water temps
   and compressors to use for whichever step is being run. 

   Args:
      unitInfo: Used to give prompts on compressor types and number.
      fanKW: Used in calculating total power of the unit.
      waterTemp: Water temps used to prompt user.
      dbTemp: Starting db temps to run convergence off of.
              These will be updated as the convergence runs.

   Returns:
      List of all values collected from running the convergence at each
      stage of capacity. Set up as a 4x5 array.

      [[% Load, EDB, Capacity, Power, Efficiency],...]

      Each row contains the same values for the different steps.
   """
   
   tempCap = 0 # The temporary capacity after the convergence has been run
   fullCap = 0 # The capacity after the first convergence used in the following runs
   percentLoad = [0,0,0,0]
   IPLVTable = [[0,0,0,0,0],[0,0,0,0,0],[0,0,0,0,0],[0,0,0,0,0]]
   print(unitInfo)
   print("Set the LWT to 44°F")
   stepRange = [0,1,2,3]
   if unitInfo[4] == 'singles':
      stepRange = [0,2] # Singles only need convergences on step 1 and 3
   
   # Each step requires a different number of compressors.
   # The following handles the output to describe which
   # compressors to use. It then runs the convergence. 
   for step in stepRange:
      if unitInfo[4] == "tandems":
         if step == 0:
            print("\nSet circuit 1 to {0} tandem and circuit 2 to {1} tandem".format(unitInfo[5], unitInfo[6]))
         elif step == 1:
            print("\nSet circuit 1 to {0} tandem and circuit 2 to {1}".format(unitInfo[5], unitInfo[6]))
         elif step == 2:
            print("\nSet circuit 1 to {0} and circuit 2 to {1}".format(unitInfo[5], unitInfo[6]))
         else:
            print("\nSet circuit 1 to {0} and set circuit 2 to 0".format(unitInfo[5]))
      if unitInfo[4] == "singles":
         if step == 0:
            print("\nSet circuit 1 to {0} and circuit 2 to {1}".format(unitInfo[5], unitInfo[6]))
         else:
            print("\nSet circuit 1 to {0} and set circuit 2 to 0".format(unitInfo[5]))
      print('Set the EWT to {}°F'.format(waterTemp[step]))
      tempCap = converge(dbTemp, step, fullCap, fanKW, IPLVTable, percentLoad, unitInfo[4])
      
      # Sets the capacity of the first run to be used later. 
      if step == 0:
         fullCap = tempCap
   return IPLVTable

def tonCalc(compQs:list) -> list:
   """Converts BTUs to Tons.

   Takes in the values input from the user, BTUs from JESS, and converts
   them into tons to be used in the rest of the calculations.

   Args:
      compQs: [Qactual, Qcond] From Jess in BTUs.

   Returns:
      Qactual, Qcond, and Win in Tons to be used in the rest of the program.
      The values are kept in a 1x3 list.
   """
   
   compCalc = [0,0,0]
   compCalc[0] = compQs[0]/12000.0
   compCalc[1] = compQs[1]/12000.0
   compCalc[2] = (compQs[1]-compQs[0])/3412.0
   return compCalc

def converge(dbTemp:int, step:int, fullCap:int, fanKW:list, IPLVTable:list, percentLoad:list, comp:str) -> float:
   """Itterates to determine correct db temp.

   Itterates, changing the EDB and EWB temperatures to reach the condenser
   relief point. Will loop until the 'dbTemp' for the step and 'dbCorrect'
   are within 'targetDelta' of each other. Each loop will adjust the
   'dbTemp[step]' by the 'actualDelta' {bdCorrect - dbTemp[step]} multiplied
   by the 'deltaAdjustment'

   Args:
      dbTemp: Starting db temps to run convergence off of.
              These will be updated as the convergence runs.
      step: The current step the program is running. 
      fullCap: The Capacity of the first convergence
      fanKW: Used in calculating total power of the unit. 
      IPLVTable: Table of values collected through the convergences.
      percentLoad: List of percent difference between fullCap and and
                   the current capacity.

   Returns:
      The capacity calculted in the convergence. Returned a a float.
      Only the first time it is returned is kept to be used later. 
   """
   
   actualDelta = 0 # The difference between the actual and corrected db
   targetDelta = .1 # How close the actual and corrected db need to be
   # The adjustment value each time itteraing. This value is used because
   # the corrected db overcorrects going past the final temperature.
   deltaAdjustment = .75 
   dbCorrect = 0
   compQs = [0,0]
   print('Run the selection with {0:.3f}°F db and {1:.3f}°F wb'.format(dbTemp[step], dbTemp[step]-8))
   compQs[0] = float(input("What is the Qactual of Circuit 1?\n"))
   compQs[1] = float(input("What is the QCond of Circuit 1?\n"))
   compQs[0] += float(input("What is the Qactual of Circuit 2?\n'0' if no circuit 2\n"))
   compQs[1] += float(input("What is the QCond of Circuit 2?\n'0' if no circuit 2\n"))
   compCalc = tonCalc(compQs)
   if step == 0:
      percentLoad[step] = 1
   elif step == 3 or comp == "singles":
      percentLoad[step] = compCalc[0]/fullCap
   else:
      percentLoad[step] = compCalc[0]/fullCap
      dbCorrect = (percentLoad[step]*100)*0.6+35 # From AHRI Standard 550/590 5.4.1.2
      actualDelta = dbCorrect - dbTemp[step]
      if abs(actualDelta) <= targetDelta:
         print('\nConvergence Complete')
      else:
         print('\nConvergence Incomplete. Itterating...')
         dbTemp[step] += actualDelta*deltaAdjustment
         converge(dbTemp, step, fullCap, fanKW, IPLVTable, percentLoad, comp)
         return -1
   IPLVTable[step] = fillTable(step, fanKW, compCalc, dbTemp, percentLoad)
   return compCalc[0]

def fillTable(step:int, fanKW:list, compCalc:list, dbTemp:int, percentLoad:list) -> list:
   """Fills the IPLV Table with information from the Converge.

   The different Lists are pulled together into one organized list.
   This list will be output 

   Args:
      step: The current step the program is on
      fanKW: List of values to be added to the total power
      compCalc: Values for capacity in Tons
      dbTemp: List of temperatures after the convergence has run
      percentLoad: Percent difference of current and full capacity

   Returns:
      List of values collected from a convergence for a single step 
      of capacity

      [% Load, EDB, Capacity, Power, Efficiency]
   """
   
   IPLVRow = [0,0,0,0,0]
   IPLVRow[0] = percentLoad[step]
   IPLVRow[1] = dbTemp[step]
   IPLVRow[2] = compCalc[0]
   IPLVRow[3] = compCalc[2]+.5

   # Adding the correct fan power to the step. Each time
   # the table is filled one less fan power is added to 
   # account for the part load. The current assumption is
   # half the fans running at half capacity rounded up
   for num in range(4-step):
      IPLVRow[3] += fanKW[num] 
   IPLVRow[4] = 12/(IPLVRow[3]/IPLVRow[2])
   return IPLVRow

def IPLVCalc(IPLVTable:list, unitInfo:list) -> list:
   """Calculates IPLV values from the raw data.

   Uses the data from the IPLV list to calculate the calcTableList.
   Calculations come from From AHRI Standard 550/590 for Integrated
   Part Load Values.

   KW/Ton = 12/EER

   For Values <= Final Calculated Capacity:
      EER = Efficiency/Cd
      Cd = (-.13*LF) + 13
      LF = (%Load*FullLoad)/StepLoad

   For Values > Final Calculated Capacity;
      EER = ((Rating%-CurrentRating%)/(PastRating%-CurrentRating%))*
            (PastEER-CurrentEER)+CurrentEER

   Args:
      IPLVTable: list of values to be calculated to fine IPLV
      UnitInfo: List containing Unit Information

   Returns:
      list of calculated IPLVs to be displayed to the user. The final
      portion of the list is dedicated to the total EER and KW/Ton
      values of the unit.

      [[LoadPoint%, KW/Ton, EER],...,
       ['Part Loads', EER, KW/Ton]]
   """
   
   calcTable = [[1,0,0],[.75,0,0],[.5,0,0],[.25,0,0],['Part Loads',0,0]]
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
   # Calculating the Integrated Part Load Values.
   # From AHRI Standard 550/590 5.4.1.3.
   for num in range(4):
      calcTable[4][1] += partLoadMults[num]*calcTable[num][2]
      calcTable[4][2] += partLoadMults[num]/calcTable[num][1]
   calcTable[4][2] = 1/calcTable[4][2]
   return calcTable

def printResults(calcTable:list, IPLVTable:list, unitInfo:list):
   """Formats output into a readable form.

   Takes in the two large lists and outputs them in a table format. 
   For loops are uses to print each line of the table

   Args:
      calcTable: The Calculated IPLV values 
      IPLVTable: The Raw data pulled from the Convergence and fan DataFrame
      unitInfo: Description of the unit
   """
   
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

if __name__ == "__main__":
   main()