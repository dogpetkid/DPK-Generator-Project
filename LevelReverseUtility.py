"""
"""

import argparse
import io
import json
import re
import shutil
import typing

import numpy
import pandas
import xlrd

import DatablockIO
import EnumConverter
import XlsxInterfacer

# argparse: used to get arguments in CLI (to decide which files to turn into levels encoding/decoding and which file)
# io:       used to read from and write to files
# json:     used to export the data to a json
# re:       used to preform regex searches
# shutil:   used to copy the template
# typing:   used to give types to function parameters
# numpy:    used to manipulate the inconsistant numpy data read by pandas
# pandas:   used to read the excel data
# xlrd:     used to catch and throw excel errors when initially reading the sheets

# a regex to capture the newlines the devs put into their json
devnewlnregex = "(\\\\n|\\\\r){1,2}"

# Settings
#####
# Version number meaning:
# R.G.S
# R: Rundown
# G: Generator
# S: Sheet (minor changes to the sheet are insignificant to the generator)
Version = "4e.1"
# relative path to location for datablocks, defaultly its folder should be on the same layer as this project's folder
blockpath = "../Datablocks/"
# default paths to xlsx files when running the program
defaultpaths = ["in.xlsx"]
# persistentID of the default rundown to insert levels into
rundowndefault = 25 # R4
#####

# load all datablock files
if True:
    # DATABLOCK_Rundown = DatablockIO.datablock(open(blockpath+"RundownDataBlock.json", 'r', encoding="utf8"))
    # DATABLOCK_LevelLayout = DatablockIO.datablock(open(blockpath+"LevelLayoutDataBlock.json", 'r', encoding="utf8"))
    # DATABLOCK_WardenObjective = DatablockIO.datablock(open(blockpath+"WardenObjectiveDataBlock.json", 'r', encoding="utf8"))
    DATABLOCK_ComplexResourceSet = DatablockIO.datablock(open(blockpath+"ComplexResourceSetDataBlock.json", 'r', encoding="utf8"))
    DATABLOCK_LightSettings = DatablockIO.datablock(open(blockpath+"LightSettingsDataBlock.json", 'r', encoding="utf8"))
    DATABLOCK_FogSettings = DatablockIO.datablock(open(blockpath+"FogSettingsDataBlock.json", 'r', encoding="utf8"))
    DATABLOCK_EnemyPopulation = DatablockIO.datablock(open(blockpath+"EnemyPopulationDataBlock.json", 'r', encoding="utf8"))
    DATABLOCK_ExpeditionBalance = DatablockIO.datablock(open(blockpath+"ExpeditionBalanceDataBlock.json", 'r', encoding="utf8"))
    DATABLOCK_SurvivalWaveSettings = DatablockIO.datablock(open(blockpath+"SurvivalWaveSettingsDataBlock.json", 'r', encoding="utf8"))
    DATABLOCK_SurvivalWavePopulation = DatablockIO.datablock(open(blockpath+"SurvivalWavePopulationDataBlock.json", 'r', encoding="utf8"))
    DATABLOCK_ChainedPuzzle = DatablockIO.datablock(open(blockpath+"ChainedPuzzleDataBlock.json", 'r', encoding="utf8"))
    DATABLOCK_EnemyGroup = DatablockIO.datablock(open(blockpath+"EnemyGroupDataBlock.json", 'r', encoding="utf8"))
    DATABLOCK_ConsumableDistribution = DatablockIO.datablock(open(blockpath+"ConsumableDistributionDataBlock.json", 'r', encoding="utf8"))
    DATABLOCK_BigPickupDistribution = DatablockIO.datablock(open(blockpath+"BigPickupDistributionDataBlock.json", 'r', encoding="utf8"))
    DATABLOCK_StaticSpawn = DatablockIO.datablock(open(blockpath+"StaticSpawnDataBlock.json", 'r', encoding="utf8"))
    DATABLOCK_Item = DatablockIO.datablock(open(blockpath+"ItemDataBlock.json", 'r', encoding="utf8"))

# load all enum files
if True:
    ENUMFILE_eExpeditionAccessibility = open(blockpath+"TypeList/Enums/eExpeditionAccessibility.txt",'r')
    ENUMFILE_eLocalZoneIndex = open(blockpath+"TypeList/Enums/eLocalZoneIndex.txt",'r')
    ENUMFILE_eWardenObjectiveWinCondition = open(blockpath+"TypeList/Enums/eWardenObjectiveWinCondition.txt",'r')
    ENUMFILE_LG_LayerType = open(blockpath+"TypeList/Enums/LG_LayerType.txt",'r')
    ENUMFILE_SubComplex = open(blockpath+"TypeList/Enums/SubComplex.txt",'r')
    ENUMFILE_eZoneBuildFromType = open(blockpath+"TypeList/Enums/eZoneBuildFromType.txt",'r')
    ENUMFILE_eZoneBuildFromExpansionType = open(blockpath+"TypeList/Enums/eZoneBuildFromExpansionType.txt",'r')
    ENUMFILE_eZoneExpansionType = open(blockpath+"TypeList/Enums/eZoneExpansionType.txt",'r')
    ENUMFILE_eWantedZoneHeighs = open(blockpath+"TypeList/Enums/eWantedZoneHeighs.txt",'r')
    ENUMFILE_eProgressionPuzzleType = open(blockpath+"TypeList/Enums/eProgressionPuzzleType.txt",'r')
    ENUMFILE_eSecurityGateType = open(blockpath+"TypeList/Enums/eSecurityGateType.txt",'r')
    ENUMFILE_eEnemyGroupType = open(blockpath+"TypeList/Enums/eEnemyGroupType.txt",'r')
    ENUMFILE_eEnemyRoleDifficulty = open(blockpath+"TypeList/Enums/eEnemyRoleDifficulty.txt",'r')
    ENUMFILE_eEnemyZoneDistribution = open(blockpath+"TypeList/Enums/eEnemyZoneDistribution.txt",'r')
    ENUMFILE_eZoneDistributionAmount = open(blockpath+"TypeList/Enums/eZoneDistributionAmount.txt",'r')
    ENUMFILE_TERM_State = open(blockpath+"TypeList/Enums/TERM_State.txt",'r')
    ENUMFILE_LG_StaticDistributionWeightType = open(blockpath+"TypeList/Enums/LG_StaticDistributionWeightType.txt",'r')
    ENUMFILE_eWardenObjectiveType = open(blockpath+"TypeList/Enums/eWardenObjectiveType.txt",'r')
    ENUMFILE_eRetrieveExitWaveTrigger = open(blockpath+"TypeList/Enums/eRetrieveExitWaveTrigger.txt",'r')
    ENUMFILE_eWardenObjectiveEventTrigger = open(blockpath+"TypeList/Enums/eWardenObjectiveEventTrigger.txt",'r')
    ENUMFILE_eWardenObjectiveEventType = open(blockpath+"TypeList/Enums/eWardenObjectiveEventType.txt",'r')
    ENUMFILE_eReactorWaveSpawnType = open(blockpath+"TypeList/Enums/eReactorWaveSpawnType.txt",'r')

def writePublicNameFromDict(datablock:DatablockIO.datablock, interface:XlsxInterfacer.interface, x:int, y:int, dictionary:dict, key:str):
    """
    Takes a datablock and writes the publicName associated to the persisentID in the specified cell
    """
    DatablockIO.idInDict(datablock, dictionary, key)
    interface.writeFromDict(x, y, dictionary, key)
    # convert the name back into a persistentID since the value that was changed is still part of the dictionary
    DatablockIO.nameInDict(datablock, dictionary, key)

def writeEnumFromDict(enum:io.FileIO, interface:XlsxInterfacer.interface, x:int, y:int, dictionary:dict, key:str):
    """
    Takes an enum's index from a dictionary and convert it to a name to write in the specified cell
    """
    EnumConverter.indexInDict(enum, dictionary, key)
    interface.writeFromDict(x, y, dictionary, key)
    # convert the index back to an enum since the value that was changed is still a part of the dictionary
    EnumConverter.enumInDict(enum, dictionary, key)

def frameMeta(iMeta:XlsxInterfacer.interface, rundownID:int, tier:str, index:int):
    """
    edit the iMeta pandas dataFrame
    """
    iMeta.write(rundownID, 0, 2)
    iMeta.write(tier, 1, 2)
    iMeta.write(index, 2, 2)

def frameExpeditionInTier(iExpeditionInTier:XlsxInterfacer.interface, ExpeditionInTierData:dict):
    """
    edit the iExpeditionInTier pandas dataFrame
    """
    iExpeditionInTier.writeFromDict(0, 2, ExpeditionInTierData, "Enabled")
    writeEnumFromDict(ENUMFILE_eExpeditionAccessibility, iExpeditionInTier, 1, 2, ExpeditionInTierData, "Accessibility")

    iExpeditionInTier.writeFromDict(12, 12, ExpeditionInTierData["CustomProgressionLock"], "MainSectors")
    iExpeditionInTier.writeFromDict(12, 13, ExpeditionInTierData["CustomProgressionLock"], "SecondarySectors")
    iExpeditionInTier.writeFromDict(12, 14, ExpeditionInTierData["CustomProgressionLock"], "ThirdSectors")
    iExpeditionInTier.writeFromDict(12, 15, ExpeditionInTierData["CustomProgressionLock"], "AllClearedSectors")

    iExpeditionInTier.writeFromDict(12, 2, ExpeditionInTierData["Descriptive"], "Prefix")
    iExpeditionInTier.writeFromDict(12, 3, ExpeditionInTierData["Descriptive"], "PublicName")
    iExpeditionInTier.writeFromDict(12, 4, ExpeditionInTierData["Descriptive"], "IsExtraExpedition")
    iExpeditionInTier.writeFromDict(12, 5, ExpeditionInTierData["Descriptive"], "ExpeditionDepth")
    iExpeditionInTier.writeFromDict(12, 6, ExpeditionInTierData["Descriptive"], "EstimatedDuration")
    iExpeditionInTier.writeFromDict(12, 7, ExpeditionInTierData["Descriptive"], "ExpeditionDescription")
    iExpeditionInTier.writeFromDict(12, 8, ExpeditionInTierData["Descriptive"], "RoleplayedWardenIntel")
    iExpeditionInTier.writeFromDict(12, 9, ExpeditionInTierData["Descriptive"], "DevInfo")

    iExpeditionInTier.writeFromDict(0, 6, ExpeditionInTierData["Seeds"], "BuildSeed")
    iExpeditionInTier.writeFromDict(1, 6, ExpeditionInTierData["Seeds"], "FunctionMarkerOffset")
    iExpeditionInTier.writeFromDict(2, 6, ExpeditionInTierData["Seeds"], "StandardMarkerOffset")
    iExpeditionInTier.writeFromDict(3, 6, ExpeditionInTierData["Seeds"], "LightJobSeedOffset")

    writePublicNameFromDict(DATABLOCK_ComplexResourceSet, iExpeditionInTier, 0, 10, ExpeditionInTierData["Expedition"], "ComplexResourceData")
    writePublicNameFromDict(DATABLOCK_LightSettings, iExpeditionInTier, 1, 10, ExpeditionInTierData["Expedition"], "LightSettings")
    writePublicNameFromDict(DATABLOCK_FogSettings, iExpeditionInTier, 2, 10, ExpeditionInTierData["Expedition"], "FogSettings")
    writePublicNameFromDict(DATABLOCK_EnemyPopulation, iExpeditionInTier, 3, 10, ExpeditionInTierData["Expedition"], "EnemyPopulation")
    writePublicNameFromDict(DATABLOCK_ExpeditionBalance, iExpeditionInTier, 4, 10, ExpeditionInTierData["Expedition"], "ExpeditionBalance")
    writePublicNameFromDict(DATABLOCK_SurvivalWaveSettings, iExpeditionInTier, 5, 10, ExpeditionInTierData["Expedition"], "ScoutWaveSettings")
    writePublicNameFromDict(DATABLOCK_SurvivalWavePopulation, iExpeditionInTier, 6, 10, ExpeditionInTierData["Expedition"], "ScoutWavePopulation")

def getExpeditionInTierData(levelIdentifier:str, RundownDataBlock:DatablockIO.datablock):
    """
    Outputs the ExpeditionInTierData for a specified level
    The levelIdentifier can be either the name the level OR the RUNDOWN,TIER,INDEX of a level delimited by commas as shown
    e.g.
    "Cuernos"
    "Contact,Cuernos"
    "Contact,C,2"
    "Rundown 004 - EA,C,2"
    "Rundown 004 - EA,2,2"
    "25,2,2"
    (DO NOT CONFLATE DATA/TYPES BETWEEN EXAMPLES, IT MAY NOT WORK)
    """

    # deteremine what data has been given
    levelName = None    # str
    rundown = None      # int, str
    levelTier = None    # int, str
    levelIndex = None   # int

    if (levelIdentifier.find(",")==-1):
        # check for name only
        levelName = levelIdentifier
    else:
        # multiple items of information given, split them and parse them
        splitIdentifier = levelIdentifier.split(",")
        if (len(splitIdentifier)==2):
            # split length is two, therefore it is the rundown name and level name
            rundown, levelName = splitIdentifier
        else:
            # split length is greater than two, therefore it must be rundown,tier,index
            rundown, levelTier, levelIndex = splitIdentifier[0:3]

    def searchLevelInRundown(block, levelName):
        """
        search a rundown for a specific level by the level name
        returns rundown,tier,index
        """
        found = False
        for Tier in "ABCDE":
            # search through tiers A-E
            if found: break
            i = 0
            for ExpeditionInTierData in block["Tier"+Tier]:
                # search through expeditions in Tier
                if found: break
                try: # not all levels have PublicNames so a try except is required
                    if ExpeditionInTierData["Descriptive"]["PublicName"] == levelName:
                        # on finding the level, note it is found to skip further searches and note the tier and index
                        found = True
                        rundown = block["persistentID"]
                        levelTier = Tier
                        levelIndex = i
                except KeyError:
                    pass
                i+= 1
        if not(found): rundown, levelTier, levelIndex = -1,-1,-1
        return rundown, levelTier, levelIndex

    def rundownValueToIndex(RundownDataBlock, value):
        """
        takes the persistentID, PublicName, or title (or part of the title) of a rundown
        returns the index of the rundown (-1 if the rundown does not exist)
        """
        rundownIndex = RundownDataBlock.find(rundown)
        found = rundownIndex != -1
        if not(found):
            # if rundown index is -1, then the value of rundown must describe a portion of the title
            for block in RundownDataBlock.data["Blocks"]:
                if found: break
                rundownIndex+= 1
                # if the value of rundown describes some portion of the title, this is the rundown to return
                if block["StorytellingData"]["Title"].lower().find(rundown.lower()) != -1: found = True
        return [-1,rundownIndex][found]

    if (rundown==None):
        # if rundown is None, no information was given about it, therefore the level must be searched through all rundowns
        # the first level with the name will be the match
        #  then fill out rundown,tier,index
        rundown, levelTier, levelIndex = -1,-1,-1
        for block in RundownDataBlock.data["Blocks"]:
            # break if proper data is found
            if [rundown, levelTier, levelIndex] != [-1,-1,-1]: break
            rundown, levelTier, levelIndex = searchLevelInRundown(block, levelName)

    # convert numerical rundown persistentID to int
    try: rundown = int(rundown)
    except: pass

    if (levelTier==None or levelIndex==None):
        # if level index or tier is None, the rundown and levelName should be known
        # assume value of rundown describes either the rundown name or persistentID
        rundownIndex = rundownValueToIndex(RundownDataBlock,rundown)
        # with the rundown index, now the specific block with the level can be searched for
        if rundownIndex != -1:
            rundown, levelTier, levelIndex = searchLevelInRundown(RundownDataBlock.data["Blocks"][rundownIndex], levelName)

    # print(rundown,levelTier,levelIndex)

    if (-1 in [rundown, levelTier, levelIndex] or None in [rundown, levelTier, levelIndex]):
        return [[],-1,"",-1]

    # convert numerical tier to A-E
    try: levelTier = chr(65+int(levelTier))
    except: pass
    # make sure the tier letter is upper cased
    levelTier = levelTier.upper()

    # make sure the levelIndex is a number
    try:
        levelIndex = int(levelIndex)
    except ValueError:
        return [[],-1,"",-1]

    rundownIndex = rundownValueToIndex(RundownDataBlock,rundown)
    # if no such rundown exists, return a blank array
    if rundownIndex < 0:
        return [[],-1,"",-1]

    # get the persistentID of the rundown
    rundown = RundownDataBlock.data["Blocks"][rundownIndex]["persistentID"]

    try:
        return RundownDataBlock.data["Blocks"][rundownIndex]["Tier"+levelTier][levelIndex],rundown,"Tier"+levelTier,levelIndex
    except KeyError:
        return [[],-1,"",-1]
    except IndexError:
        return [[],-1,"",-1]

def UtilityJob(desiredReverse:str, RundownDataBlock, LevelLayoutBlock, WardenObjectiveDataBlock, silent:bool=False, debug:bool=False):
    """
    Have the utility start a job
    Take an identifier of which level to reverse (see below)
    Will output an xlsx file in the template format (same format to be fed into the LevelUtility)
    desiredReverse must follow one of the following example's format:
    "Cuernos"
    "Contact,Cuernos"
    "Contact,C,2"
    "Rundown 004 - EA,C,2"
    "Rundown 004 - EA,2,2"
    "25,2,2"
    (DO NOT CONFLATE DATA/TYPES BETWEEN EXAMPLES, IT MAY NOT WORK)
    """

    ExpeditionInTierData, rundown, levelTier, levelIndex = getExpeditionInTierData(desiredReverse, RundownDataBlock)

    # print(ExpeditionInTierData)
    # print(rundown,levelTier,levelIndex)

    if (rundown == -1 or levelTier == "" or levelIndex == -1 or ExpeditionInTierData==[]):
        # if no such level exists
        if not(silent):print("No level found, searched: "+desiredReverse)
        return False

    try:
        # get the name of the level if it exists (so the file name can be the name of the level)
        levelName = RundownDataBlock.data["Blocks"][RundownDataBlock.find(rundown)][levelTier][levelIndex]["Descriptive"]["PublicName"]
    except KeyError:
        levelName = desiredReverse

    # TODO remove all xml formatting from level names
    # TODO make the template to be pulled from a constant at the top of the file
    shutil.copy("Template for Generator R4.5.xlsx",levelName+".xlsx")
    fxlsx = open(levelName+".xlsx", 'rb+')

    iMeta = XlsxInterfacer.interface(pandas.read_excel(fxlsx, "Meta", header=None))
    iExpeditionInTier = XlsxInterfacer.interface(pandas.read_excel(fxlsx, "ExpeditionInTier", header=None))

    fxlsx.close()

    # sheets that need to be written
    # Meta
    # ExpeditionInTier
    # LX ExpeditionZoneData
    # LX ExpeditionZoneData Lists
    # LX WardenObjective
    # LX WardenObjective ReactorWaves

    # TODO other json to dataFrame functions

    frameMeta(iMeta, rundown, levelTier, levelIndex)
    frameExpeditionInTier(iExpeditionInTier, ExpeditionInTierData)

    # writer = pandas.ExcelWriter(levelName+".xlsx", engine='xlsxwriter')
    # writer = pandas.ExcelWriter(fxlsx, engine="openpyxl", mode="a")
    # iMeta.frame.to_excel(writer, sheet_name="Meta")

    iMeta.save(levelName+".xlsx", "Meta")
    iExpeditionInTier.save(levelName+".xlsx", "ExpeditionInTier")

def main():
    # TODO get level via input or args
    # level to reverse
    desiredReverse = "Septic"

    # Open Datablocks to get level from
    RundownDataBlock =  DatablockIO.datablock(open(blockpath+"RundownDataBlock.json", 'r', encoding="utf-8"))
    LevelLayoutDataBlock = DatablockIO.datablock(open(blockpath+"LevelLayoutDataBlock.json", 'r', encoding="utf8"))
    WardenObjectiveDataBlock = DatablockIO.datablock(open(blockpath+"WardenObjectiveDataBlock.json", 'r', encoding="utf8"))

    # TODO allow for program to take a list of levels to run utility on
    UtilityJob(desiredReverse, RundownDataBlock, LevelLayoutDataBlock, WardenObjectiveDataBlock)

if __name__ == "__main__":
    main()
