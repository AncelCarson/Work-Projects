# pylint: disable=invalid-name,bad-indentation
# -*- coding: utf-8 -*-

# Author: Ancel Carson
# Orginization: Napps Technology Comporation
# Creation Date: 8/7/2025
# Update Date: 9/7/2025
# Get_Job_Pipe_Files.py

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
from collections import defaultdict

import win32com.client
from dotenv import load_dotenv

#Secret Variables
load_dotenv()
Shared_Drive = os.getenv('Shared_Drive')

#custom Modules
#pylint: disable=wrong-import-position
sys.path.insert(0,fr'\\{Shared_Drive}\Programs\Add_ins')
from MenuMaker import makeMenu
from Loader import Loader
#pylint: enable=wrong-import-position

#Variables
assembly_file_path = [r"C:\NTC Vault\products\acca\acc-ll-liquid circuits\ACCH-LLDH-0800.SLDASM",
                      r"C:\NTC Vault\products\acca\acc-ll-liquid circuits\ACCH-LLDH-0801.SLDASM"]

#Classes
class Get_Job_Pipe_Files:
    def __init__(self):
        loader = Loader("Opening Solidworks...", "Solidworks Opened\n", .1).start()
        self.sw_app = win32com.client.Dispatch("SldWorks.Application")
        self.sw_app.Visible = 0  # Optional: Make SolidWorks visible
        loader.stop()

        loader = Loader("Connecting to the Vault...", "Vault Connected\n", .1).start()
        self.vaultPath = os.getenv('Vault_Path')
        self.vault = win32com.client.Dispatch("ConisioLib.EdmVault")
        self.vault.LoginAuto("NTC Vault", 0)
        loader.stop()

        self.files = None
        self.filePaths = []
        self.parts = defaultdict(int)

    def __call__(self):
        pass

    def __enter__(self):
        pass

    def __exit__(self, *exc):
        pass

    def getPipes(self):
        """Gets the list of pipes and their quantities from a list of PDFs"""
        drawingNames = []
        for file in self.files:
            fileName = os.path.basename(file).split(".")[0]
            if '_' in fileName:
                drawingNames.append(fileName.split("_")[0])
            else:
                drawingNames.append(fileName)
        self.getPaths(drawingNames)
        for file in self.filePaths:
            loader = Loader("Loading Solidworks File...", "File Loaded\n", .1).start()
            model = self.sw_app.OpenDoc(file, 2)
            loader.stop()
            config = self.getConfig(model)
            loader = Loader(f"Collecting Parts from {model.GetTitle}...",
                            "Parts Collected\n", .05).start()
            partList = self.getPartList(config)
            self._traverse_components(partList)
            self.sw_app.CloseDoc(model.GetTitle)
            loader.stop()
        for part, qty in self.parts.items():
            print(f"{part}: {qty}")
        return self.parts

    def setFiles(self, files: list[str]):
        """Sets the list of files to be processed"""
        self.files = files

    def getPaths(self, files: list[str]):
        """Getsd the File Paths for a list of files"""
        for file in files:
            self.filePaths.append(self.getPath(file))

    def getPath(self, file: str) -> str:
        """Searches PDM For the path of a file name"""
        search = self.vault.CreateSearch()
        search.FindFiles = True
        search.FileName = file + ".SLDASM"
        result = search.GetFirstResult()
        return result.Path

    def getConfig(self, model: object) -> str:
        """Gets the selected configuration from a SolidWorks Assembly"""
        configs = model.GetConfigurationNames
        if len(configs) != 1:
            makeMenu(f"Configurations for {model.GetTitle}", configs)
            selection = int(input("Which Configuration is needed?\n"))
            config = model.GetConfigurationByName(configs[selection - 1])
            return config
        return configs
    
    def getPartList(self, config: object) -> list[object]:
        """Gets a list of part objects from the assembly configuration"""
        root_components = config.GetRootComponent3(False)
        return root_components.GetChildren
    
    def _traverse_components(self, component_array):
        """Runs through the list of parts and adds select ones to a dictionary"""
        for comp in component_array:
            if comp.IsSuppressed:
                continue

            path = comp.GetPathName
            name = os.path.basename(path)
            if "TBGA" not in name:
                continue

            model_doc = comp.GetModelDoc2
            if model_doc and model_doc.GetType == 1:
                self.parts[name.split(".")[0]] += 1

            children = comp.GetChildren
            if children:
                self._traverse_components(children)

#Functions

if __name__ == "__main__":
    # extract_bom_from_assembly(assembly_file_path)
    # chat_example()
    instance = Get_Job_Pipe_Files()
    instance.setFiles(assembly_file_path)
    instance.getPipes()
    input("Program completed. Press ENTER to close...")
