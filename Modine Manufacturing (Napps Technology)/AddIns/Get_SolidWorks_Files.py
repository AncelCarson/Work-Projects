# pylint: disable=invalid-name,bad-indentation,no-member
# -*- coding: utf-8 -*-

# Author: Ancel Carson
# Orginization: Napps Technology Comporation
# Creation Date: 30/3/26
# Update Date: 30/3/26
# Get_SolidWorks_Files.py

"""This program Collects the varying copper parts for a job.

Given a folder this program will read for PDFs and find the matching
assembly in the PDM Vault. On finding the assembly, it will search the BOM
for bent pipe files and add them to the original folder.

Classes:
    Get_Job_Copper_Files: Handles the file collection and holds the SolidWorks instance
"""
#Libraries
import os
import sys
import pythoncom
from collections import defaultdict

import win32com.client
from dotenv import load_dotenv
from win32com.client import VARIANT

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
assembly_file_path = [r"G:\_A NTC GENERAL FILES\_JOB FILES\Job 9272A-01 Rod Hollenberg (2)ACCU030 HP 460V 8 17 25\Drawings - Mechanical\ACCU-STFN-0300_0.pdf"]

#Classes
class Get_SolidWorks_Files:
    """Handles the creation of new users and writing to the user file.

    Attributes:
        sw_app (Solidworks Instance): Instance of solidworks running hte processes
        vault (PDM Instance): PDM Vault used for file searching
        files (list(str)): The list of PDF files to search
        filePaths (list(str)): File Locations in the PDM Vault
        parts (dict(int)): A disctionary of pipe files and quantities 

    Functions:
        getParts: Driver of the process
        setFiles: Sets the file location to start the search in
        getPaths: Cyccles through the list of files to get their paths 
        getPath: Retrieves the file path from the Vault
        getConfig: Lists the configs of a model for user selection
        getPartList: Gets the list of pipe parts from an assembly
    """
    def __init__(self):
        loader = Loader("Opening Solidworks...", "Solidworks Opened\n", .1).start()
        self.sw_app = win32com.client.Dispatch("SldWorks.Application")
        self.sw_app.Visible = 0  # Optional: Make SolidWorks visible
        loader.stop()

        loader = Loader("Connecting to the Vault...", "Vault Connected\n", .1).start()
        self.vault = win32com.client.Dispatch("ConisioLib.EdmVault")
        self.vault.LoginAuto("NTC Vault", 0)
        loader.stop()

        self.files = None
        self.filePaths = []
        self.pipes = defaultdict(int)
        self.parts = defaultdict(int)

    def __call__(self):
        pass

    def __enter__(self):
        pass

    def __exit__(self, *exc):
        pass

    def getParts(self):
        """Gets the list of pipes and their quantities from a list of PDFs"""
        drawingNames = []
        self.parts = defaultdict(int)
        for file in self.files:
            fileName = os.path.basename(file).split(".")[0]
            if '_' in fileName:
                drawingNames.append(fileName.split("_")[0])
            else:
                drawingNames.append(fileName)
        self.getPaths(drawingNames)
        # self.filePaths = self.getLatestVersion(self.filePaths)
        for file in self.filePaths:
            self.checkVersion(file)
            loader = Loader("Loading Solidworks File...", "File Loaded\n", .1).start()
            model = self.sw_app.OpenDoc(file, 2)
            loader.stop()
            config = self.getConfig(model)
            # loader = Loader(f"Collecting Parts from {model.GetTitle}...",
            #                 f"Parts Collected from {model.GetTitle}\n", .05).start()
            partList = self.getPartList(config)
            self._traverse_fittings(partList)
            self._traverse_pipes(partList)
            self.sw_app.CloseDoc(model.GetTitle)
            # loader.stop()
        for part, qty in self.parts.items():
            print(f"{part}: {qty}")
        for pipe, qty in self.pipes.items():
            print(f"{pipe}: {qty}")
        return [self.parts,self.pipes]

    def setFiles(self, files: list[str]):
        """Sets the list of files to be processed"""
        self.files = None
        self.files = files

    def getPaths(self, files: list[str]):
        """Getsd the File Paths for a list of files"""
        self.filePaths = []
        for file in files:
            filePath = self.getPath(file)
            if filePath is None:
                print(f"{file} is not in the vault and has been skipped")
                continue
            self.filePaths.append(filePath)

    def getPath(self, file: str) -> str:
        """Searches PDM For the path of a file name"""
        search = self.vault.CreateSearch()
        search.FindFiles = True
        search.FileName = file + ".SLDASM"
        result = search.GetFirstResult()
        try:
            filePath = result.Path
        except AttributeError:
            filePath = None
        return filePath

    def checkVersion(self, file: str):
        """Checks if the user has the most recent file version and prompts if they do not"""
        edmFile, folder = self.vault.GetFileFromPath(file, None)
        if edmFile:
            localVer = edmFile.GetLocalVersionNo(folder.ID)
            latestVer = edmFile.CurrentVersion
            if localVer < latestVer:
                fileName = os.path.basename(file)
                print(f"{fileName} is cached as an old version and will be updated.")

    def getConfig(self, model: object) -> str:
        """Gets the selected configuration from a SolidWorks Assembly"""
        configs = model.GetConfigurationNames
        if len(configs) != 1:
            makeMenu(f"Configurations for {model.GetTitle}", configs)
            selection = int(input("Which Configuration is needed?\n"))
            model.ShowConfiguration2(configs[selection - 1])
            config = model.GetConfigurationByName(configs[selection - 1])
            return config
        return model.GetConfigurationByName(configs[0])

    def getPartList(self, config: object) -> list[object]:
        """Gets a list of part objects from the assembly configuration"""
        root_components = config.GetRootComponent3(False)
        return root_components.GetChildren

    def _traverse_fittings(self, component_array):
        """Runs through the list of parts and adds select ones to a dictionary"""
        partList = ["BRKT","CPPR-FTGW","HDWR","BRSS-FTTG"]
        skippedBrass = ["BRSS-FTTG-0001","BRSS-FTTG-0005","BRSS-FTTG-0010","BRSS-FTTG-0040"]
        for comp in component_array:
            if comp.IsSuppressed:
                continue

            path = comp.GetPathName
            name = os.path.basename(path)
            if len(name) == 21:
                name = name.split(".")[0]
            else:
                name = self._get_PartNumber(comp)

            if name in ("",None):
                continue

            if not any(part.lower() in name.lower() for part in partList):
                continue

            if name.upper() in skippedBrass:
                continue

            model_doc = comp.GetModelDoc2
            if model_doc and model_doc.GetType == 1:
                self.parts[name.split(".")[0]] += 1

            children = comp.GetChildren
            if children:
                self._traverse_fittings(children)

    def _traverse_pipes(self, component_array):
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
                self.pipes[name.split(".")[0]] += 1

            children = comp.GetChildren
            if children:
                self._traverse_pipes(children)

    def _get_PartNumber(self, comp) -> str:
        compDoc = comp.GetModelDoc2
        compConfig = comp.ReferencedConfiguration
        prop = compDoc.Extension.CustomPropertyManager(compConfig)
        raw = VARIANT(pythoncom.VT_BSTR | pythoncom.VT_BYREF, "")
        resolved = VARIANT(pythoncom.VT_BSTR | pythoncom.VT_BYREF, "")
        wasResolved = VARIANT(pythoncom.VT_BOOL | pythoncom.VT_BYREF, False)
        prop.Get5("Part Number",False,raw,resolved,wasResolved)
        if wasResolved.value:
            return resolved.value
        return None

#Functions

if __name__ == "__main__":
    # extract_bom_from_assembly(assembly_file_path)
    # chat_example()
    instance = Get_SolidWorks_Files()
    instance.setFiles(assembly_file_path)
    instance.getParts()
    input("Program completed. Press ENTER to close...")
