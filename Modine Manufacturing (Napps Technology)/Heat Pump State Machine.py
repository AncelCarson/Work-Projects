### Ancel Carson
### Created: 14/9/2023
### UPdated: 14/9/2023
### Windows 11
### Python command line, Notepad, IDLE
### Heat Pump State machine.py

# Libraries

# Object Class
class Chiller():
   def __init__(self, number):
      self.number = number    # Unit Number in Array
      self.state = "Off"      # Current State of the Unit
      self.chilledLoad = 0    # Percentage of Total Cooling Capacity
      self.heatLoad = 0       # Percentage of Total Cooling Capacity
      self.heatMax = 100      # Percentage of Total Unit Heating Capacity Currently Available
      # Valve States
      self.valvesOpen = {"Chilled": False, "HR": False, "Hot": False}
   
   # Set Unit to Cooling Mode
   def startCool(self):
      self.state = "Cooling"
      self.setValves()
   
   # Set Unit to Heat Recovery Mode
   def startHR(self):
      self.state = "HeatRecovery"
      self.setValves()
   
   # Set Unit to Heating Mode
   def startHeat(self):
      self.state = "Heating"
      self.setValves()
   
   # Set Unit to Off Mode
   def turnOff(self):
      self.state = "Off"
      self.chilledLoad = 0
      self.heatLoad = 0
      self.setValves()
   
   # Returns the Current State of the Unit
   def getState(self):
      return self.state
   
   # Returns the Percetage of Unit Maximum Chilled Load Currently Running
   def getChilledLoad(self):
      return self.chilledLoad
   
   # Returns the Percetage of Unit Maximum Heating Load Currently Running
   def getHeatLoad(self):
      return self.heatLoad
   
   # Sets the Percetage of Unit Maximum Chilled Load Currently Running
   def setChilledLoad(self, value):
      self.chilledLoad = value
      self.heatMax = self.chilledLoad
   
   # Sets the Percetage of Unit Maximum Heating Load Currently Running
   def setHeatLoad(self, value):
      self.heatLoad = value

   # Sets the valve states based off of the Current Unit State
   def setValves(self):
      if self.state == "Cooling":
         self.valvesOpen = {"Chilled": True, "HR": False, "Hot": False}
      if self.state == "HeatRecovery":
         self.valvesOpen = {"Chilled": True, "HR": True, "Hot": False}
      if self.state == "Heating":
         self.valvesOpen = {"Chilled": False, "HR": False, "Hot": True}
      if self.state == "Off":
         self.valvesOpen = {"Chilled": False, "HR": False, "Hot": False}

   # Displays in thest the current state of the Unit
   def showStatus(self):
      print("|  Unit #{} Current Operation  |".format(self.number))
      print("--------------------------------")
      print("Current State: {}".format(self.state))
      print("--------------------------------")
      print("Percent Cooling Capacity: {}%".format(self.chilledLoad))
      print("Percent Heating Capacity: {}%".format(self.heatLoad))
      print("--------------------------------")
      print("Chilled Valve Open: {}".format(self.valvesOpen["Chilled"]))
      print("HeatRec Valve Open: {}".format(self.valvesOpen["HR"]))
      print("Heating Valve Open: {}".format(self.valvesOpen["Hot"]))
      print("--------------------------------\n")

# Main Function
def main():
   unitCount = int(input("How many units are in the Array?\n"))   # Gets Number of Units Needed
   units = [Chiller(count + 1) for count in range(unitCount)]     # Creates and Array of Chillers
   arrayStatus = lambda: [unit.showStatus() for unit in units]    # Creates Array Display Variable
   arrayStatus()
   
   running = True
   while(running):
      running = False


# Self Program Call
if __name__ == '__main__':
   main()
