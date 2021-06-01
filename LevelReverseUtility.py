"""
This is a tool created by DPK
"""

import argparse
import io
import json
import re
import shutil
import typing

import numpy
import openpyxl
import pandas
import xlrd

import DatablockIO
import EnumConverter
import XlsxInterfacer

# argparse: used to get arguments in CLI (to decide which files to turn into levels encoding/decoding and which file)
# io:       used to read from and write to files
# json:     used to export the data to a json
# re:       used to preform regex searches and replaces
# shutil:   used to copy the template
# typing:   used to give types to function parameters
# numpy:    used to manipulate the inconsistant numpy data read by pandas
# openpyxl: used to copy templated sheets
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
# S: Sheet (minor changes to the sheet are insignificant to the utility)
Version = "5.1"
# relative path to location for datablocks, defaultly its folder should be on the same layer as this project's folder
blockpath = "../Datablocks/" # TODO create an argument to change the blockpath
# path to template file
templatepath = "Template for Generator R5.xlsx"
# default paths to xlsx files when running the program
defaultpaths = ["in.xlsx"]
# persistentID of the default rundown to insert levels into
rundowndefault = 26 # R5
#####

def writePublicNameFromDict(datablock:DatablockIO.datablock, interface:XlsxInterfacer.interface, x:int, y:int, dictionary:dict, key:str):
    """
    Takes a datablock and writes the publicName associated to the persisentID in the specified cell
    """
    try:
        if(str(dictionary[key])=="0"):return
        # This is to catch and not write any "0" datablocks used by the devs
        # This should let past all non-zero, even ones not in the blocks
    except KeyError:return
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


# load all datablock files
if True:
    # DATABLOCK_Rundown = DatablockIO.datablock(open(blockpath+"RundownDataBlock.json", 'r', encoding="utf8"))
    # DATABLOCK_LevelLayout = DatablockIO.datablock(open(blockpath+"LevelLayoutDataBlock.json", 'r', encoding="utf8"))
    # DATABLOCK_WardenObjective = DatablockIO.datablock(open(blockpath+"WardenObjectiveDataBlock.json", 'r', encoding="utf8"))
    DATABLOCK_ArtifactDistributionDataBlock = DatablockIO.datablock(open(blockpath+"ArtifactDistributionDataBlock.json", 'r', encoding="utf8"))
    DATABLOCK_BigPickupDistribution = DatablockIO.datablock(open(blockpath+"BigPickupDistributionDataBlock.json", 'r', encoding="utf8"))
    DATABLOCK_ChainedPuzzle = DatablockIO.datablock(open(blockpath+"ChainedPuzzleDataBlock.json", 'r', encoding="utf8"))
    DATABLOCK_ComplexResourceSet = DatablockIO.datablock(open(blockpath+"ComplexResourceSetDataBlock.json", 'r', encoding="utf8"))
    DATABLOCK_ConsumableDistribution = DatablockIO.datablock(open(blockpath+"ConsumableDistributionDataBlock.json", 'r', encoding="utf8"))
    DATABLOCK_Enemy = DatablockIO.datablock(open(blockpath+"EnemyDataBlock.json", 'r', encoding="utf8"))
    DATABLOCK_EnemyGroup = DatablockIO.datablock(open(blockpath+"EnemyGroupDataBlock.json", 'r', encoding="utf8"))
    DATABLOCK_EnemyPopulation = DatablockIO.datablock(open(blockpath+"EnemyPopulationDataBlock.json", 'r', encoding="utf8"))
    DATABLOCK_ExpeditionBalance = DatablockIO.datablock(open(blockpath+"ExpeditionBalanceDataBlock.json", 'r', encoding="utf8"))
    DATABLOCK_FogSettings = DatablockIO.datablock(open(blockpath+"FogSettingsDataBlock.json", 'r', encoding="utf8"))
    DATABLOCK_Item = DatablockIO.datablock(open(blockpath+"ItemDataBlock.json", 'r', encoding="utf8"))
    DATABLOCK_LightSettings = DatablockIO.datablock(open(blockpath+"LightSettingsDataBlock.json", 'r', encoding="utf8"))
    DATABLOCK_StaticSpawn = DatablockIO.datablock(open(blockpath+"StaticSpawnDataBlock.json", 'r', encoding="utf8"))
    DATABLOCK_SurvivalWavePopulation = DatablockIO.datablock(open(blockpath+"SurvivalWavePopulationDataBlock.json", 'r', encoding="utf8"))
    DATABLOCK_SurvivalWaveSettings = DatablockIO.datablock(open(blockpath+"SurvivalWaveSettingsDataBlock.json", 'r', encoding="utf8"))

# load all enum files
# TODO use dummy files in place of enum files as the enum files aren't used in R5
if True:
    ENUMFILE_LG_LayerType = open(blockpath+"TypeList/Enums/LG_LayerType.txt",'r')
    ENUMFILE_LG_StaticDistributionWeightType = open(blockpath+"TypeList/Enums/LG_StaticDistributionWeightType.txt",'r')
    ENUMFILE_SubComplex = open(blockpath+"TypeList/Enums/SubComplex.txt",'r')
    ENUMFILE_TERM_State = open(blockpath+"TypeList/Enums/TERM_State.txt",'r')
    ENUMFILE_eEnemyGroupType = open(blockpath+"TypeList/Enums/eEnemyGroupType.txt",'r')
    ENUMFILE_eEnemyRoleDifficulty = open(blockpath+"TypeList/Enums/eEnemyRoleDifficulty.txt",'r')
    ENUMFILE_eEnemyZoneDistribution = open(blockpath+"TypeList/Enums/eEnemyZoneDistribution.txt",'r')
    ENUMFILE_eExpeditionAccessibility = open(blockpath+"TypeList/Enums/eExpeditionAccessibility.txt",'r')
    ENUMFILE_eLocalZoneIndex = open(blockpath+"TypeList/Enums/eLocalZoneIndex.txt",'r')
    ENUMFILE_eProgressionPuzzleType = open(blockpath+"TypeList/Enums/eProgressionPuzzleType.txt",'r')
    ENUMFILE_eReactorWaveSpawnType = open(blockpath+"TypeList/Enums/eReactorWaveSpawnType.txt",'r')
    ENUMFILE_eRetrieveExitWaveTrigger = open(blockpath+"TypeList/Enums/eRetrieveExitWaveTrigger.txt",'r')
    ENUMFILE_eSecurityGateType = open(blockpath+"TypeList/Enums/eSecurityGateType.txt",'r')
    ENUMFILE_eWantedZoneHeighs = open(blockpath+"TypeList/Enums/eWantedZoneHeighs.txt",'r')
    ENUMFILE_eWardenObjectiveEventTrigger = open(blockpath+"TypeList/Enums/eWardenObjectiveEventTrigger.txt",'r')
    ENUMFILE_eWardenObjectiveEventType = open(blockpath+"TypeList/Enums/eWardenObjectiveEventType.txt",'r')
    ENUMFILE_eWardenObjectiveType = open(blockpath+"TypeList/Enums/eWardenObjectiveType.txt",'r')
    ENUMFILE_eWardenObjectiveWinCondition = open(blockpath+"TypeList/Enums/eWardenObjectiveWinCondition.txt",'r')
    ENUMFILE_eZoneBuildFromExpansionType = open(blockpath+"TypeList/Enums/eZoneBuildFromExpansionType.txt",'r')
    ENUMFILE_eZoneBuildFromType = open(blockpath+"TypeList/Enums/eZoneBuildFromType.txt",'r')
    ENUMFILE_eZoneDistributionAmount = open(blockpath+"TypeList/Enums/eZoneDistributionAmount.txt",'r')
    ENUMFILE_eZoneExpansionType = open(blockpath+"TypeList/Enums/eZoneExpansionType.txt",'r')

def ZonePlacementData(interface:XlsxInterfacer.interface, data:dict, col:int, row:int, horizontal=True):
    """
    add ZonePlacementData to the specified interface \n
    col and row define the upper left value (not header) \n
    horizontal is true if the values are in the same row
    """
    writeEnumFromDict(ENUMFILE_eLocalZoneIndex, interface, col, row, data, "LocalIndex")
    try:ZonePlacementWeights(interface, data["Weights"], col+horizontal, row+(not horizontal), horizontal)
    except KeyError:pass

def BulkheadDoorPlacementData(interface:XlsxInterfacer.interface, data:dict, col:int, row:int, horizontal=False):
    """
    add BulkheadDoorPlacementData to the specified interface
    col and row define the upper left value (not header) \n
    horizontal is true if the values are in the same row
    """
    writeEnumFromDict(ENUMFILE_eLocalZoneIndex, interface, col, row, data, "ZoneIndex")
    try:ZonePlacementWeights(interface, data["PlacementWeights"], col+horizontal, row+(not horizontal), horizontal)
    except KeyError:pass
    interface.writeFromDict(col+4*horizontal, row+4*(not horizontal), data, "AreaSeedOffset")
    interface.writeFromDict(col+5*horizontal, row+5*(not horizontal), data, "MarkerSeedOffset")

def FunctionPlacementData(interface:XlsxInterfacer.interface, data:dict, col:int, row:int, horizontal=True):
    """
    add FunctionPlacementData to the specified interface
    col and row define the upper left value (not header) \n
    horizontal is true if the values are in the same row
    """
    try:ZonePlacementData(interface, data["PlacementWeights"], col, row, horizontal)
    except KeyError:pass
    interface.writeFromDict(col+3*horizontal, row+3*(not horizontal), data, "AreaSeedOffset")
    interface.writeFromDict(col+4*horizontal, row+4*(not horizontal), data, "MarkerSeedOffset")

def ZonePlacementWeightsList(interface:XlsxInterfacer.interface, data:typing.List[typing.List[dict]], col:int, row:int, horizontal:bool=True):
    """
    add ZonePlacementWeightsList to the specified interface
    col and row define the upper left value (not header) \n
    horizontal describes the direction to iterate to write weights
    """
    letter = 'A'
    for group in data:
        for placement in group:
            interface.write(letter, col, row)
            writeEnumFromDict(ENUMFILE_eLocalZoneIndex, interface, col+(not horizontal), row+horizontal, placement, "LocalIndex")
            ZonePlacementWeights(interface, placement["Weights"], col+2*(not horizontal), row+2*horizontal, horizontal=(not horizontal))
            col+= horizontal 
            row+= not horizontal
        letter = chr(ord(letter)+1)

def ZonePlacementWeights(interface:XlsxInterfacer.interface, data:dict, col:int, row:int, horizontal:bool=True):
    """
    add ZonePlacementWeights to the specified interface
    col and row define the upper left value (not header) \n
    horizontal is true if the values are in the same row
    """
    interface.writeFromDict(col, row, data, "Start")
    interface.writeFromDict(col+horizontal, row+(not horizontal), data, "Middle")
    interface.writeFromDict(col+2*horizontal, row+2*(not horizontal), data, "End")

def GenericEnemyWaveData(interface:XlsxInterfacer.interface, data:dict, col:int, row:int, horizontal:bool=False):
    """
    add GenericEnemyWaveData to the specified interface
    col and row define the upper left value (not header) \n
    horizontal is true if the values are in the same row
    """
    writePublicNameFromDict(DATABLOCK_SurvivalWaveSettings, interface, col, row, data, "WaveSettings")
    writePublicNameFromDict(DATABLOCK_SurvivalWavePopulation, interface, col+horizontal, row+(not horizontal), data, "WavePopulation")
    interface.writeFromDict(col+2*horizontal, row+2*(not horizontal), data, "SpawnDelay")
    interface.writeFromDict(col+3*horizontal, row+3*(not horizontal), data, "TriggerAlarm")
    interface.writeFromDict(col+4*horizontal, row+4*(not horizontal), data, "IntelMessage")

def GenericEnemyWaveDataList(interface:XlsxInterfacer.interface, data:typing.List[dict], col:int, row:int, horizontal:bool=True):
    """
    add GenericEnemyWaveDataList to the specified interface
    col and row define the upper left value (not header) \n
    horizontal describes the direction to iterate to look for waves
    """
    for wave in data:
        GenericEnemyWaveData(interface, wave, col, row, horizontal=(not horizontal))
        col+= horizontal
        row+= not horizontal

def ArtifactZoneDistribution(interface:XlsxInterfacer.interface, data:dict, col:int, row:int, horizontal:bool=False):
    """
    add ArtifactZoneDistribution to the specified interface
    col and row define the upper left value (not header) \n
    horizontal is true if the values are in the same row
    """
    writeEnumFromDict(ENUMFILE_eLocalZoneIndex, interface, col, row, data, "Zone")
    interface.writeFromDict(col+horizontal, row+(not horizontal), data, "BasicArtifactWeight")
    interface.writeFromDict(col+2*horizontal, row+2*(not horizontal), data, "AdvancedArtifactWeight")
    interface.writeFromDict(col+3*horizontal, row+3*(not horizontal), data, "SpecializedArtifactWeight")

def WardenObjectiveEventData(interface:XlsxInterfacer.interface, data:dict, col:int, row:int, horizontal:bool=False):
    """
    add WardenObjectiveEventData to the specified interface
    col and row define the upper left value (not header) \n
    horizontal is true if the values are in the same row
    """
    writeEnumFromDict(ENUMFILE_eWardenObjectiveEventTrigger, interface, col, row, data, "Trigger")
    writeEnumFromDict(ENUMFILE_eWardenObjectiveEventType, interface, col+horizontal, row+(not horizontal), data, "Type")
    writeEnumFromDict(ENUMFILE_LG_LayerType, interface, col+2*horizontal, row+2*(not horizontal), data, "Type")
    writeEnumFromDict(ENUMFILE_eLocalZoneIndex, interface, col+3*horizontal, row+3*(not horizontal), data, "LocalIndex")
    interface.writeFromDict(col+4*horizontal, row+4*(not horizontal), dict, "Delay")
    interface.writeFromDict(col+5*horizontal, row+5*(not horizontal), dict, "WardenIntel")
    interface.writeFromDict(col+6*horizontal, row+6*(not horizontal), dict, "SoundID")
    # TODO convert sound id to name of sound

def GeneralFogDataStep(interface:XlsxInterfacer.interface, data:dict, col:int, row:int, horizontal:bool=False):
    """
    add GeneralFogDataStep to the specified interface
    col and row define the upper left value (not header) \n
    horizontal is true if the values are in the same row
    """
    writePublicNameFromDict(DATABLOCK_FogSettings, interface, col, row, data, "m_fogDataId")
    interface.writeFromDict(col+horizontal, row+(not horizontal), data, "m_transitionToTime")


def frameMeta(iMeta:XlsxInterfacer.interface, rundownID:int, tier:str, index:int):
    """
    edit the iMeta pandas dataFrame
    """
    iMeta.write(rundownID, 0, 2)
    iMeta.write(tier, 1, 2)
    iMeta.write(index, 2, 2)

def LayerData(interface:XlsxInterfacer.interface, data:dict, col:int, row:int):
    """
    add LayerData to the specified interface \n
    col and row define the upper left value, SHOULD be the header \n
    horizontal is true if the values are in the same row
    """
    interface.writeFromDict(col+5, row, data, "ZoneAliasStart")
    try:
        itercol,iterrow = col+5, row+1
        for zone in data["ZonesWithBulkheadEntrance"]:
            # NOTE textmode may need a toggle in this file for whether the json should have text enums
            interface.write(EnumConverter.enumToIndex(ENUMFILE_eLocalZoneIndex, zone, textmode=True), itercol, iterrow)
            itercol+= 1
    except KeyError:pass
    try:
        itercol,iterrow = col+5, row+2
        for placement in data["BulkheadDoorControllerPlacements"]:
            BulkheadDoorPlacementData(interface, placement, itercol, iterrow, horizontal=False)
            itercol+= 1
    except KeyError:pass
    ZonePlacementWeightsList(interface, data["BulkheadKeyPlacements"], col+5, row+8, horizontal=True)
    try:
        interface.writeFromDict(col+5, row+13, data["ObjectiveData"], "DataBlockId")
        writeEnumFromDict(ENUMFILE_eWardenObjectiveWinCondition, interface, col+5, row+14, data["ObjectiveData"], "WinCondition")
        ZonePlacementWeightsList(interface, data["ObjectiveData"]["ZonePlacementDatas"], col+5, row+15, horizontal=True)
    except KeyError:pass
    try:
        interface.writeFromDict(col+5, row+20, data["ArtifactData"], "ArtifactAmountMulti")
        writePublicNameFromDict(DATABLOCK_ArtifactDistributionDataBlock, interface, col+5, row+21, data["ArtifactData"], "ArtifactLayerDistributionDataID")
        itercol,iterrow = col+5, row+22
        for distribution in data["ArtifactData"]["ArtifactZoneDistributions"]:
            ArtifactZoneDistribution(interface, distribution, itercol, iterrow, horizontal=False)
            itercol+= 1
    except KeyError:pass

def frameExpeditionInTier(iExpeditionInTier:XlsxInterfacer.interface, ExpeditionInTierData:dict):
    """
    edit the iExpeditionInTier pandas dataFrame
    """
    iExpeditionInTier.writeFromDict(0, 2, ExpeditionInTierData, "Enabled")
    writeEnumFromDict(ENUMFILE_eExpeditionAccessibility, iExpeditionInTier, 1, 2, ExpeditionInTierData, "Accessibility")

    try:
        iExpeditionInTier.writeFromDict(12, 2, ExpeditionInTierData["CustomProgressionLock"], "MainSectors")
        iExpeditionInTier.writeFromDict(12, 3, ExpeditionInTierData["CustomProgressionLock"], "SecondarySectors")
        iExpeditionInTier.writeFromDict(12, 4, ExpeditionInTierData["CustomProgressionLock"], "ThirdSectors")
        iExpeditionInTier.writeFromDict(12, 5, ExpeditionInTierData["CustomProgressionLock"], "AllClearedSectors")
    except KeyError:pass

    try:
        iExpeditionInTier.writeFromDict(12, 8, ExpeditionInTierData["Descriptive"], "Prefix")
        iExpeditionInTier.writeFromDict(12, 9, ExpeditionInTierData["Descriptive"], "PublicName")
        iExpeditionInTier.writeFromDict(12, 10, ExpeditionInTierData["Descriptive"], "IsExtraExpedition")
        iExpeditionInTier.writeFromDict(12, 11, ExpeditionInTierData["Descriptive"], "ExpeditionDepth")
        iExpeditionInTier.writeFromDict(12, 12, ExpeditionInTierData["Descriptive"], "EstimatedDuration")
        # TODO regex replace new lines to be "\n"
        iExpeditionInTier.writeFromDict(12, 13, ExpeditionInTierData["Descriptive"], "ExpeditionDescription")
        iExpeditionInTier.writeFromDict(12, 14, ExpeditionInTierData["Descriptive"], "RoleplayedWardenIntel")
        iExpeditionInTier.writeFromDict(12, 15, ExpeditionInTierData["Descriptive"], "DevInfo")
    except KeyError:pass

    try:
        iExpeditionInTier.writeFromDict(0, 6, ExpeditionInTierData["Seeds"], "BuildSeed")
        iExpeditionInTier.writeFromDict(1, 6, ExpeditionInTierData["Seeds"], "FunctionMarkerOffset")
        iExpeditionInTier.writeFromDict(2, 6, ExpeditionInTierData["Seeds"], "StandardMarkerOffset")
        iExpeditionInTier.writeFromDict(3, 6, ExpeditionInTierData["Seeds"], "LightJobSeedOffset")
    except KeyError:pass

    try:
        writePublicNameFromDict(DATABLOCK_ComplexResourceSet, iExpeditionInTier, 0, 10, ExpeditionInTierData["Expedition"], "ComplexResourceData")
        writePublicNameFromDict(DATABLOCK_LightSettings, iExpeditionInTier, 1, 10, ExpeditionInTierData["Expedition"], "LightSettings")
        writePublicNameFromDict(DATABLOCK_FogSettings, iExpeditionInTier, 2, 10, ExpeditionInTierData["Expedition"], "FogSettings")
        writePublicNameFromDict(DATABLOCK_EnemyPopulation, iExpeditionInTier, 3, 10, ExpeditionInTierData["Expedition"], "EnemyPopulation")
        writePublicNameFromDict(DATABLOCK_ExpeditionBalance, iExpeditionInTier, 4, 10, ExpeditionInTierData["Expedition"], "ExpeditionBalance")
        writePublicNameFromDict(DATABLOCK_SurvivalWaveSettings, iExpeditionInTier, 5, 10, ExpeditionInTierData["Expedition"], "ScoutWaveSettings")
        writePublicNameFromDict(DATABLOCK_SurvivalWavePopulation, iExpeditionInTier, 6, 10, ExpeditionInTierData["Expedition"], "ScoutWavePopulation")
    except KeyError:pass

    iExpeditionInTier.writeFromDict(0, 13, ExpeditionInTierData, "LevelLayoutData")
    try:
        LayerData(iExpeditionInTier, ExpeditionInTierData["MainLayerData"], 0, 20)
    except KeyError:pass

    iExpeditionInTier.writeFromDict(3, 13, ExpeditionInTierData, "SecondaryLayerEnabled")
    iExpeditionInTier.writeFromDict(4, 13, ExpeditionInTierData, "SecondaryLayout")
    try:
        writeEnumFromDict(ENUMFILE_LG_LayerType, iExpeditionInTier, 3, 17, ExpeditionInTierData["BuildSecondaryFrom"], "LayerType")
        iExpeditionInTier.writeFromDict(4, 17, ExpeditionInTierData["BuildSecondaryFrom"], "Zone")
    except KeyError:pass
    try:
        LayerData(iExpeditionInTier, ExpeditionInTierData["SecondaryLayerData"], 0, 48)
    except KeyError:pass

    iExpeditionInTier.writeFromDict(7, 13, ExpeditionInTierData, "ThirdLayerEnabled")
    iExpeditionInTier.writeFromDict(8, 13, ExpeditionInTierData, "ThirdLayout")
    try:
        writeEnumFromDict(ENUMFILE_LG_LayerType, iExpeditionInTier, 7, 17, ExpeditionInTierData["BuildThirdFrom"], "LayerType")
        iExpeditionInTier.writeFromDict(8, 17, ExpeditionInTierData["BuildThirdFrom"], "Zone")
    except KeyError:pass
    try:
        LayerData(iExpeditionInTier, ExpeditionInTierData["ThirdLayerData"], 0, 76)
    except KeyError:pass

    try:
        iExpeditionInTier.writeFromDict(5, 6, ExpeditionInTierData["SpecialOverrideData"], "WeakResourceContainerWithPackChanceForLocked")
    except KeyError:pass

class ExpeditionZoneDataLists:
    """a class that decreases the dimentions of the dictionaries in ExpeditionZoneData (since the sheet cannot contain 2d-3d data)"""

    def __init__(self, LevelLayoutBlock:dict):
        """Generates numerous stubs to be iterated through and written to the interface"""

        self.stubEventsOnEnter = []
        self.stubProgressionPuzzleToEnter = []
        self.stubEnemySpawningInZone = []
        self.stubEnemyRespawnExcludeList = []
        self.stubTerminalPlacements = []
        self.stubLocalLogFiles = []
        groupedLocalLogFiles = []
        self.stubParsedLog = []
        groupedParsedLog = []
        self.stubPowerGeneratorPlacements = []
        self.stubDisinfectionStationPlacements = []
        self.stubStaticSpawnDataContainers = []

        loggroupstart = 'A'
        parsedgroupstart = 'A'

        for ZoneData in LevelLayoutBlock["Zones"]:
            try:
                zone = EnumConverter.indexToEnum(ENUMFILE_eLocalZoneIndex, ZoneData["LocalIndex"], False)
            except KeyError:
                zone = None # TODO possibly put a warning that a zone is missing a local index

            # EventsOnEnter
            try:
                for e in ZoneData["EventsOnEnter"]:
                    self.stubEventsOnEnter.append(dict({"ZoneIndex":zone},**e))
            except KeyError:pass

            # ProgressionPuzzleToEnter
            try:
                for e in ZoneData["ProgressionPuzzleToEnter"]["ZonePlacementData"]:
                    self.stubProgressionPuzzleToEnter.append({"ZoneIndex":zone, "ZonePlacementData":e})
            except KeyError:pass

            # EnemySpawningInZone
            try:
                for e in ZoneData["EnemySpawningInZone"]:
                    self.stubEnemySpawningInZone.append(dict({"ZoneIndex":zone},**e))
                    # TODO have the reverse utility write the Reminder as all possible groups of enemies that can spawn as a comma separated list
            except KeyError:pass

            # EnemyRespawnExcludeList
            try:
                for e in ZoneData["EnemyRespawnExcludeList"]: # e is an enum, so...
                    self.stubEnemyRespawnExcludeList.append({"ZoneIndex":zone, "value":e}) # it should be added as a value instead of adding dicts together
            except KeyError:pass

            # PowerGeneratorPlacements
            try:
                for e in ZoneData["PowerGeneratorPlacements"]:
                    self.stubPowerGeneratorPlacements.append(dict({"ZoneIndex":zone},**e))
            except KeyError:pass

            # DisinfectionStationPlacements
            try:
                for e in ZoneData["DisinfectionStationPlacements"]:
                    self.stubDisinfectionStationPlacements.append(dict({"ZoneIndex":zone},**e))
            except KeyError:pass

            # StaticSpawnDataContainers
            try:
                for e in ZoneData["StaticSpawnDataContainers"]:
                    self.stubStaticSpawnDataContainers.append(dict({"ZoneIndex":zone},**e))
            except KeyError:pass

            # Unlike the other lists, the terminal placement is several dimentions deep and must be handled piece by piece to find unique ParsedLogs and Log files
            try:
                # TerminalPlacements
                for placement in ZoneData["TerminalPlacements"]:

                    try: # if there are log files...
                        logfiles = placement["LocalLogFiles"] # will jump to the except if no logs exist

                        if logfiles == []:
                            # if there are no logs, keep the group blank
                            placement["LocalLogFiles"] = ""
                            raise KeyError # jump to the no logs exist

                        for logfile in logfiles:

                            # ParsedLog
                            try: # if there are pared logs...
                                parseds = logfile["ParsedLog"] # will jump to the except if no parsed exist

                                if parseds == []:
                                    # if there are no parsed logs, keep the group blank
                                    logfile["Parsed Group"] = ""
                                    raise KeyError # jump to the no parsed exist

                                try: # if parseds group has been handled, find it's group
                                    parsedsindex = groupedParsedLog.index(parseds)
                                except ValueError: # parseds has not already been handled
                                    groupedParsedLog.append(parseds)
                                    parsedsindex = len(groupedParsedLog) - 1
                                    for parsed in parseds: # add each parsed log to the stub
                                        self.stubParsedLog.append({"Parsed Group":chr(ord(parsedgroupstart)+parsedsindex),"value":parsed})
                                finally: # finally, change the log file to reflect the parsed group
                                    logfile["Parsed Group"] = chr(ord(parsedgroupstart)+parsedsindex)

                            except KeyError:pass # no parsed exist, pass

                        # XXX something is wrong with the logs
                        # LocalLogFiles
                        try: # if log goup has been handled, find it's group
                            logfilesindex = groupedLocalLogFiles.index(logfiles)
                        except ValueError: # log group has not already been handled
                            groupedLocalLogFiles.append(logfiles)
                            logfilesindex = len(groupedLocalLogFiles) - 1
                            for logfile in logfiles:
                                self.stubLocalLogFiles.append(dict({"Log Group":chr(ord(loggroupstart)+logfilesindex)},**logfile))
                        finally: # finally, change the log file to reflect the parsed group
                            placement["LocalLogFiles"] = chr(ord(loggroupstart)+logfilesindex)

                    except KeyError:pass # no logs exist, pass

                    self.stubTerminalPlacements.append(dict({"ZoneIndex":zone},**placement))

            except KeyError:pass # no terminals exist, pass


    def write(self, iExpeditionZoneDataLists:XlsxInterfacer.interface):
        startrow = 2

        startcolEventsOnEnter = XlsxInterfacer.ctn("A")
        startcolProgressionPuzzleToEnter = XlsxInterfacer.ctn("K")
        startcolEnemySpawningInZone = XlsxInterfacer.ctn("R")
        startcolEnemyRespawnExcludeList = XlsxInterfacer.ctn("Y")
        startcolTerminalPlacements = XlsxInterfacer.ctn("AB")
        startcolLocalLogFiles = XlsxInterfacer.ctn("AM")
        startcolParsedLog = XlsxInterfacer.ctn("AU")
        startcolPowerGeneratorPlacements = XlsxInterfacer.ctn("AX")
        startcolDisinfectionStationPlacements = XlsxInterfacer.ctn("BE")
        startcolStaticSpawnDataContainers = XlsxInterfacer.ctn("BL")

        row = startrow
        # EventsOnEnter
        for Snippet in self.stubEventsOnEnter:
            writeEnumFromDict(ENUMFILE_eLocalZoneIndex, iExpeditionZoneDataLists, startcolEventsOnEnter, row, Snippet, "ZoneIndex")
            iExpeditionZoneDataLists.writeFromDict(startcolEventsOnEnter+1, row, Snippet, "Delay")

            try:
                iExpeditionZoneDataLists.writeFromDict(startcolEventsOnEnter+2, row, Snippet["Noise"], "Enabled")
                iExpeditionZoneDataLists.writeFromDict(startcolEventsOnEnter+3, row, Snippet["Noise"], "RadiusMin")
                iExpeditionZoneDataLists.writeFromDict(startcolEventsOnEnter+4, row, Snippet["Noise"], "RadiusMax")
            except KeyError:pass

            try:
                iExpeditionZoneDataLists.writeFromDict(startcolEventsOnEnter+5, row, Snippet["Intel"], "Enabled")
                iExpeditionZoneDataLists.writeFromDict(startcolEventsOnEnter+6, row, Snippet["Intel"], "IntelMessage")
            except KeyError:pass

            try:
                iExpeditionZoneDataLists.writeFromDict(startcolEventsOnEnter+7, row, Snippet["Sound"], "Enabled")
                iExpeditionZoneDataLists.writeFromDict(startcolEventsOnEnter+8, row, Snippet["Sound"], "SoundEvent")
                # TODO convert sound placeholders
            except KeyError:pass

            row+= 1

        row = startrow
        # ProgressionPuzzleToEnter
        for Snippet in self.stubProgressionPuzzleToEnter:
            writeEnumFromDict(ENUMFILE_eLocalZoneIndex, iExpeditionZoneDataLists, startcolProgressionPuzzleToEnter, row, Snippet, "ZoneIndex")

            try:
                ZonePlacementData(iExpeditionZoneDataLists, Snippet["ZonePlacementData"], startcolProgressionPuzzleToEnter+2, row, horizontal=True)
            except KeyError:pass

            row+= 1

        row = startrow
        # EnemySpawningInZone
        for Snippet in self.stubEnemySpawningInZone:
            writeEnumFromDict(ENUMFILE_eLocalZoneIndex, iExpeditionZoneDataLists, startcolEnemySpawningInZone, row, Snippet, "ZoneIndex")
            iExpeditionZoneDataLists.writeFromDict(startcolEnemySpawningInZone+1, row, Snippet, "Reminder")
            writeEnumFromDict(ENUMFILE_eEnemyGroupType, iExpeditionZoneDataLists, startcolEnemySpawningInZone+2, row, Snippet, "GroupType")
            writeEnumFromDict(ENUMFILE_eEnemyRoleDifficulty, iExpeditionZoneDataLists, startcolEnemySpawningInZone+3, row, Snippet, "Difficulty")
            try: # for some reason the distribution of 7 is used when it doesn't exist so this bodge is used, thanks 10cc
                writeEnumFromDict(ENUMFILE_eEnemyZoneDistribution, iExpeditionZoneDataLists, startcolEnemySpawningInZone+4, row, Snippet, "Distribution")
            except TypeError:pass
            iExpeditionZoneDataLists.writeFromDict(startcolEnemySpawningInZone+5, row, Snippet, "DistributionValue")

            row+= 1

        row = startrow
        # EnemyRespawnExcludeList
        for Snippet in self.stubEnemyRespawnExcludeList:
            writeEnumFromDict(ENUMFILE_eLocalZoneIndex, iExpeditionZoneDataLists, startcolEnemyRespawnExcludeList, row, Snippet, "ZoneIndex")
            writePublicNameFromDict(DATABLOCK_Enemy, iExpeditionZoneDataLists, startcolEnemyRespawnExcludeList+1, row, Snippet, "value")
            row+= 1

        row = startrow
        # TerminalPlacements
        for Snippet in self.stubTerminalPlacements:
            writeEnumFromDict(ENUMFILE_eLocalZoneIndex, iExpeditionZoneDataLists, startcolTerminalPlacements, row, Snippet, "ZoneIndex")

            try:
                ZonePlacementWeights(iExpeditionZoneDataLists, Snippet["PlacementWeights"], startcolTerminalPlacements+1, row, horizontal=True)
            except KeyError:pass

            iExpeditionZoneDataLists.writeFromDict(startcolTerminalPlacements+4, row, Snippet, "AreaSeedOffset")
            iExpeditionZoneDataLists.writeFromDict(startcolTerminalPlacements+5, row, Snippet, "MarkerSeedOffset")

            iExpeditionZoneDataLists.writeFromDict(startcolTerminalPlacements+6, row, Snippet, "LocalLogFiles")

            writeEnumFromDict(ENUMFILE_TERM_State, iExpeditionZoneDataLists, startcolTerminalPlacements+7, row, Snippet, "StartingState")

            iExpeditionZoneDataLists.writeFromDict(startcolTerminalPlacements+8, row, Snippet, "AudioEventEnter")
            iExpeditionZoneDataLists.writeFromDict(startcolTerminalPlacements+9, row, Snippet, "AudioEventExit")
            # TODO convert sound placeholders

            row+= 1

        row = startrow
        # LocalLogFiles
        for Snippet in self.stubLocalLogFiles:
            iExpeditionZoneDataLists.writeFromDict(startcolLocalLogFiles, row, Snippet, "Log Group")

            iExpeditionZoneDataLists.writeFromDict(startcolLocalLogFiles+1, row, Snippet, "Parsed Group")
            iExpeditionZoneDataLists.writeFromDict(startcolLocalLogFiles+2, row, Snippet, "FileName")
            iExpeditionZoneDataLists.writeFromDict(startcolLocalLogFiles+3, row, Snippet, "FileContent")
            iExpeditionZoneDataLists.writeFromDict(startcolLocalLogFiles+4, row, Snippet, "AttachedAudioFile")
            iExpeditionZoneDataLists.writeFromDict(startcolLocalLogFiles+5, row, Snippet, "AttachedAudioByteSize")
            iExpeditionZoneDataLists.writeFromDict(startcolLocalLogFiles+6, row, Snippet, "PlayerDialogToTriggerAfterAudio")
            # TODO convert sound placeholders

            row+= 1

        row = startrow
        # ParsedLog
        for Snippet in self.stubParsedLog:
            iExpeditionZoneDataLists.writeFromDict(startcolParsedLog, row, Snippet, "Parsed Group")
            iExpeditionZoneDataLists.writeFromDict(startcolParsedLog+1, row, Snippet, "value")
            row+= 1

        row = startrow
        # PowerGeneratorPlacements
        for Snippet in self.stubPowerGeneratorPlacements:
            writeEnumFromDict(ENUMFILE_eLocalZoneIndex, iExpeditionZoneDataLists, startcolPowerGeneratorPlacements, row, Snippet, "ZoneIndex")
            FunctionPlacementData(iExpeditionZoneDataLists, Snippet, startcolPowerGeneratorPlacements+1, row, horizontal=True)
            row+= 1

        row = startrow
        # DisinfectionStationPlacements
        for Snippet in self.stubDisinfectionStationPlacements:
            writeEnumFromDict(ENUMFILE_eLocalZoneIndex, iExpeditionZoneDataLists, startcolDisinfectionStationPlacements, row, Snippet, "ZoneIndex")
            FunctionPlacementData(iExpeditionZoneDataLists, Snippet, startcolDisinfectionStationPlacements+1, row, horizontal=True)
            row+= 1

        row = startrow
        # StaticSpawnDataContainers
        for Snippet in self.stubStaticSpawnDataContainers:
            writeEnumFromDict(ENUMFILE_eLocalZoneIndex, iExpeditionZoneDataLists, startcolStaticSpawnDataContainers, row, Snippet, "ZoneIndex")
            iExpeditionZoneDataLists.writeFromDict(startcolStaticSpawnDataContainers+1, row, Snippet, "Count")
            writeEnumFromDict(ENUMFILE_LG_StaticDistributionWeightType, iExpeditionZoneDataLists, startcolStaticSpawnDataContainers+2, row ,Snippet, "DistributionWeightType")
            iExpeditionZoneDataLists.writeFromDict(startcolStaticSpawnDataContainers+3, row, Snippet, "DistributionWeight")
            iExpeditionZoneDataLists.writeFromDict(startcolStaticSpawnDataContainers+4, row, Snippet, "DistributionRandomBlend")
            iExpeditionZoneDataLists.writeFromDict(startcolStaticSpawnDataContainers+5, row, Snippet, "DistributionResultPow")
            writePublicNameFromDict(DATABLOCK_StaticSpawn, iExpeditionZoneDataLists, startcolStaticSpawnDataContainers+6, row, Snippet, "StaticSpawnDataId")
            iExpeditionZoneDataLists.writeFromDict(startcolStaticSpawnDataContainers+7, row, Snippet, "FixedSeed")
            row+= 1

def ExpeditionZoneData(iExpeditionZoneData:XlsxInterfacer.interface, ExpeditionZoneData:dict, row:int):
    """
    adds a zone to the iExpeditionZoneData (does not include any lists)
    (this would end up getting called once per layer)
    """
    # set up some checkpoints so if some of the data gets reformatted, not the entire function needs to be altered,
    # just the headings and contents of the section will need edited column values
    colPuzzleType = XlsxInterfacer.ctn("Q")
    colHSUClustersInZone = XlsxInterfacer.ctn("AH")
    colHealthMulti = XlsxInterfacer.ctn("AY")

    writeEnumFromDict(ENUMFILE_eLocalZoneIndex, iExpeditionZoneData, 0, row, ExpeditionZoneData, "LocalIndex")
    iExpeditionZoneData.writeFromDict(1, row, ExpeditionZoneData, "SubSeed")
    iExpeditionZoneData.writeFromDict(2, row, ExpeditionZoneData, "BulkheadDCScanSeed")
    writeEnumFromDict(ENUMFILE_SubComplex, iExpeditionZoneData, 3, row, ExpeditionZoneData, "SubComplex")
    iExpeditionZoneData.writeFromDict(4, row, ExpeditionZoneData, "CustomGeomorph")
    try:
        iExpeditionZoneData.writeFromDict(5, row, ExpeditionZoneData["CoverageMinMax"], "x")
        iExpeditionZoneData.writeFromDict(6, row, ExpeditionZoneData["CoverageMinMax"], "y")
    except KeyError:pass
    writeEnumFromDict(ENUMFILE_eLocalZoneIndex, iExpeditionZoneData, 7, row, ExpeditionZoneData, "BuildFromLocalIndex")
    writeEnumFromDict(ENUMFILE_eZoneBuildFromType, iExpeditionZoneData, 8, row, ExpeditionZoneData, "StartPosition")
    iExpeditionZoneData.writeFromDict(9, row, ExpeditionZoneData, "StartPosition_IndexWeight")
    writeEnumFromDict(ENUMFILE_eZoneBuildFromExpansionType, iExpeditionZoneData, 10, row, ExpeditionZoneData, "StartExpansion")
    writeEnumFromDict(ENUMFILE_eZoneExpansionType, iExpeditionZoneData, 11, row, ExpeditionZoneData, "ZoneExpansion")
    writePublicNameFromDict(DATABLOCK_LightSettings, iExpeditionZoneData, 12, row, ExpeditionZoneData, "LightSettings")
    try:
        writeEnumFromDict(ENUMFILE_eWantedZoneHeighs, iExpeditionZoneData, 13, row, ExpeditionZoneData["AltitudeData"], "AllowedZoneAltitude")
        iExpeditionZoneData.writeFromDict(14, row, ExpeditionZoneData["AltitudeData"], "ChanceToChange")
    except KeyError:pass
    # EventsOnEnter in lists

    try:
        writeEnumFromDict(ENUMFILE_eProgressionPuzzleType, iExpeditionZoneData, colPuzzleType, row, ExpeditionZoneData["ProgressionPuzzleToEnter"], "PuzzleType")
        iExpeditionZoneData.writeFromDict(colPuzzleType+1, row, ExpeditionZoneData["ProgressionPuzzleToEnter"], "CustomText")
        iExpeditionZoneData.writeFromDict(colPuzzleType+2, row, ExpeditionZoneData["ProgressionPuzzleToEnter"], "PlacementCount")
        # ProgressionPuzzleToEnter's ZonePlacementData in lists
    except KeyError:pass
    writePublicNameFromDict(DATABLOCK_ChainedPuzzle, iExpeditionZoneData, colPuzzleType+4, row, ExpeditionZoneData, "ChainedPuzzleToEnter")
    writeEnumFromDict(ENUMFILE_eSecurityGateType, iExpeditionZoneData, colPuzzleType+5, row, ExpeditionZoneData, "SecurityGateToEnter")
    try:
        iExpeditionZoneData.writeFromDict(colPuzzleType+6, row, ExpeditionZoneData["ActiveEnemyWave"], "HasActiveEnemyWave")
        writePublicNameFromDict(DATABLOCK_EnemyGroup, iExpeditionZoneData, colPuzzleType+7, row, ExpeditionZoneData["ActiveEnemyWave"], "EnemyGroupInfrontOfDoor")
        writePublicNameFromDict(DATABLOCK_EnemyGroup, iExpeditionZoneData, colPuzzleType+8, row, ExpeditionZoneData["ActiveEnemyWave"], "EnemyGroupInArea")
        iExpeditionZoneData.writeFromDict(colPuzzleType+9, row, ExpeditionZoneData["ActiveEnemyWave"], "EnemyGroupsInArea")
    except KeyError:pass
    # EnemySpawningInZone in lists
    iExpeditionZoneData.writeFromDict(colPuzzleType+11, row, ExpeditionZoneData, "EnemyRespawning")
    iExpeditionZoneData.writeFromDict(colPuzzleType+12, row, ExpeditionZoneData, "EnemyRespawnRequireOtherZone")
    iExpeditionZoneData.writeFromDict(colPuzzleType+13, row, ExpeditionZoneData, "EnemyRespawnRoomDistance")
    iExpeditionZoneData.writeFromDict(colPuzzleType+14, row, ExpeditionZoneData, "EnemyRespawnTimeInterval")
    iExpeditionZoneData.writeFromDict(colPuzzleType+15, row, ExpeditionZoneData, "EnemyRespawnCountMultiplier")
    # EnemyRespawnExcludeList in lists

    iExpeditionZoneData.writeFromDict(colHSUClustersInZone, row, ExpeditionZoneData, "HSUClustersInZone")
    iExpeditionZoneData.writeFromDict(colHSUClustersInZone+1, row, ExpeditionZoneData, "CorpseClustersInZone")
    iExpeditionZoneData.writeFromDict(colHSUClustersInZone+2, row, ExpeditionZoneData, "ResourceContainerClustersInZone")
    iExpeditionZoneData.writeFromDict(colHSUClustersInZone+3, row, ExpeditionZoneData, "GeneratorClustersInZone")
    writeEnumFromDict(ENUMFILE_eZoneDistributionAmount, iExpeditionZoneData, colHSUClustersInZone+4, row, ExpeditionZoneData, "CorpsesInZone")
    writeEnumFromDict(ENUMFILE_eZoneDistributionAmount, iExpeditionZoneData, colHSUClustersInZone+5, row, ExpeditionZoneData, "GroundSpawnersInZone")
    writeEnumFromDict(ENUMFILE_eZoneDistributionAmount, iExpeditionZoneData, colHSUClustersInZone+6, row, ExpeditionZoneData, "HSUsInZone")
    writeEnumFromDict(ENUMFILE_eZoneDistributionAmount, iExpeditionZoneData, colHSUClustersInZone+7, row, ExpeditionZoneData, "DeconUnitsInZone")
    iExpeditionZoneData.writeFromDict(colHSUClustersInZone+8, row, ExpeditionZoneData, "AllowSmallPickupsAllocation")
    iExpeditionZoneData.writeFromDict(colHSUClustersInZone+9, row, ExpeditionZoneData, "AllowResourceContainerAllocation")
    iExpeditionZoneData.writeFromDict(colHSUClustersInZone+10, row, ExpeditionZoneData, "ForceBigPickupsAllocation")
    writePublicNameFromDict(DATABLOCK_ConsumableDistribution, iExpeditionZoneData, colHSUClustersInZone+11, row, ExpeditionZoneData, "ConsumableDistributionInZone")
    writePublicNameFromDict(DATABLOCK_BigPickupDistribution, iExpeditionZoneData, colHSUClustersInZone+12, row, ExpeditionZoneData, "BigPickupDistributionInZone")
    # TerminalPlacements in lists
    iExpeditionZoneData.writeFromDict(colHSUClustersInZone+14, row, ExpeditionZoneData, "ForbidTerminalsInZone")
    # PowerGeneratorPlacements in lists
    # DisinfectionStationPlacements in lists

    iExpeditionZoneData.writeFromDict(colHealthMulti, row, ExpeditionZoneData, "HealthMulti")
    try:
        ZonePlacementWeights(iExpeditionZoneData, ExpeditionZoneData["HealthPlacement"], colHealthMulti+1, row, True)
    except KeyError:pass
    iExpeditionZoneData.writeFromDict(colHealthMulti+4, row, ExpeditionZoneData, "WeaponAmmoMulti")
    try:
        ZonePlacementWeights(iExpeditionZoneData, ExpeditionZoneData["WeaponAmmoPlacement"], colHealthMulti+5, row, True)
    except KeyError:pass
    iExpeditionZoneData.writeFromDict(colHealthMulti+8, row, ExpeditionZoneData, "ToolAmmoMulti")
    try:
        ZonePlacementWeights(iExpeditionZoneData, ExpeditionZoneData["ToolAmmoPlacement"], colHealthMulti+9, row, True)
    except KeyError:pass
    iExpeditionZoneData.writeFromDict(colHealthMulti+12, row, ExpeditionZoneData, "DisinfectionMulti")
    try:
        ZonePlacementWeights(iExpeditionZoneData, ExpeditionZoneData["DisinfectionPlacement"], colHealthMulti+13, row, True)
    except KeyError:pass
    # StaticSpawnDataContainers in lists

def LevelLayoutBlockframes(iExpeditionZoneData:XlsxInterfacer.interface, iExpeditionZoneDataLists:XlsxInterfacer.interface, LevelLayoutBlock:dict):
    """
    edit the iExpeditionZoneData and iExpeditionZoneDataLists pandas dataFrame
    """

    ExpeditionZoneDataLists(LevelLayoutBlock).write(iExpeditionZoneDataLists)

    row = 2

    for ZoneData in LevelLayoutBlock["Zones"]:
        ExpeditionZoneData(iExpeditionZoneData, ZoneData, row)
        row+= 1

def getExpeditionInTierData(levelIdentifier:str, RundownDataBlock:DatablockIO.datablock):
    """
    Outputs the ExpeditionInTierData for a specified level \n
    The levelIdentifier can be either the name the level OR the RUNDOWN,TIER,INDEX of a level delimited by commas as shown \n
    e.g. \n
    "Cuernos" \n
    "Contact,Cuernos" \n
    "Contact,C,2" \n
    "Rundown 004 - EA,C,2" \n
    "Rundown 004 - EA,2,2" \n
    "25,2,2" \n
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
        found = rundownIndex != None
        if not(found):
            rundownIndex = -1
            # if rundown index is -1, then the value of rundown must describe a portion of the title
            for block in RundownDataBlock.data["Blocks"]:
                if found: break
                rundownIndex+= 1
                # if the value of rundown describes some portion of the title, this is the rundown to return
                if block["StorytellingData"]["Title"].lower().find(rundown.lower()) != None: found = True
        return [None,rundownIndex][found]

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
        return [[],None,"",None]

    # convert numerical tier to A-E
    try: levelTier = chr(65+int(levelTier))
    except: pass
    # make sure the tier letter is upper cased
    levelTier = levelTier.upper()

    # make sure the levelIndex is a number
    try:
        levelIndex = int(levelIndex)
    except ValueError:
        return [[],None,"",None]

    rundownIndex = rundownValueToIndex(RundownDataBlock,rundown)
    # if no such rundown exists, return a blank array
    if rundownIndex in [-1, None]:
        return [[],None,"",None]

    # get the persistentID of the rundown
    rundown = RundownDataBlock.data["Blocks"][rundownIndex]["persistentID"]

    try:
        return RundownDataBlock.data["Blocks"][rundownIndex]["Tier"+levelTier][levelIndex],rundown,"Tier"+levelTier,levelIndex
    except KeyError:
        return [[],None,"",None]
    except IndexError:
        return [[],None,"",None]

def UtilityJob(desiredReverse:str, RundownDataBlock:DatablockIO.datablock, LevelLayoutBlock:DatablockIO.datablock, WardenObjectiveDataBlock:DatablockIO.datablock, silent:bool=True, debug:bool=False):
    """
    Have the utility start a job \n
    Take an identifier of which level to reverse (see below) \n
    Will output an xlsx file in the template format (same format to be fed into the LevelUtility) \n
    desiredReverse must follow one of the following example's format: \n
    "Cuernos" \n
    "Contact,Cuernos" \n
    "Contact,C,2" \n
    "Rundown 004 - EA,C,2" \n
    "Rundown 004 - EA,2,2" \n
    "25,2,2" \n
    (DO NOT CONFLATE DATA/TYPES BETWEEN EXAMPLES, IT MAY NOT WORK)
    """

    # check the template version before doing any searching
    iKey = XlsxInterfacer.interface(pandas.read_excel(templatepath, "Key", header=None))
    if iKey.read(str, 0, 5).split(".")[0:2] != Version.split(".")[0:2]:
        raise Exception("Version mismatch between utility and sheet, incompatible.")

    ExpeditionInTierData, rundown, levelTier, levelIndex = getExpeditionInTierData(desiredReverse, RundownDataBlock)

    # print(ExpeditionInTierData)
    # print(rundown,levelTier,levelIndex)

    if (rundown == None or levelTier == "" or levelIndex == None or ExpeditionInTierData==[]):
        # if no such level exists
        if not(silent):print("The search for \""+desiredReverse+"\" matched no level.")
        return False

    try:
        # get the name of the level if it exists (so the file name can be the name of the level)
        levelName = RundownDataBlock.data["Blocks"][RundownDataBlock.find(rundown)][levelTier][levelIndex]["Descriptive"]["PublicName"]
        if debug:print("The search for \""+desiredReverse+"\" found:",rundown,levelTier,levelIndex,levelName)
    except KeyError:
        levelName = desiredReverse
        if debug:print("The search for \""+desiredReverse+"\" found a nameless level:",rundown,levelTier,levelIndex)

    # XXX remove all xml formatting from level names (causes crashes)
    strippedLevelName = levelName
    try:
        shutil.copy(templatepath,strippedLevelName+".xlsx")
        fxlsx = open(strippedLevelName+".xlsx", 'rb+')
    except PermissionError:
        if not(silent):print("PermissionError opening \""+strippedLevelName+".xlsx"+"\", is it open? Cannot run utility.")
        return False

    iMeta = XlsxInterfacer.interface(pandas.read_excel(fxlsx, "Meta", header=None))
    iExpeditionInTier = XlsxInterfacer.interface(pandas.read_excel(fxlsx, "ExpeditionInTier", header=None))

    # XXX bodge for testing
    LayerDataL1 = LevelLayoutBlock.data["Blocks"][LevelLayoutBlock.find(ExpeditionInTierData["LevelLayoutData"])]
    iExpeditionZoneDataL1 = XlsxInterfacer.interface(pandas.read_excel(fxlsx, "LX ExpeditionZoneData", header=None))
    iExpeditionZoneDataListsL1 = XlsxInterfacer.interface(pandas.read_excel(fxlsx, "LX ExpeditionZoneData Lists", header=None))

    fxlsx.close()

    # XXX other json to dataFrame functions

    frameMeta(iMeta, rundown, levelTier, levelIndex)
    frameExpeditionInTier(iExpeditionInTier, ExpeditionInTierData)
    # sheets that need to be written
    # LX ExpeditionZoneData
    # LX ExpeditionZoneData Lists
    # LX WardenObjective
    # LX WardenObjective ReactorWaves

    # XXX bodge for testing
    LevelLayoutBlockframes(iExpeditionZoneDataL1, iExpeditionZoneDataListsL1, LayerDataL1)

    # writer = pandas.ExcelWriter(levelName+".xlsx", engine='xlsxwriter')
    # writer = pandas.ExcelWriter(fxlsx, engine="openpyxl", mode="a")
    # iMeta.frame.to_excel(writer, sheet_name="Meta")

    iMeta.save(strippedLevelName+".xlsx", "Meta")
    iExpeditionInTier.save(strippedLevelName+".xlsx", "ExpeditionInTier")

    # XXX bodge for testing
    workbook = openpyxl.load_workbook(filename = strippedLevelName+".xlsx")
    workbook.copy_worksheet(workbook["LX ExpeditionZoneData"]).title = "L1 ExpeditionZoneData"
    # TODO because the lists can be longer than 20 items long in total, the formatted portion should be copied down to cover all cells with values
    workbook["LX ExpeditionZoneData Lists"].title = "a" # this weird renaming is used to avoid UserWarnings by openpyxl because "LX ExpeditionZoneData Lists Copy" is too long
    workbook.copy_worksheet(workbook["a"]).title = "L1 ExpeditionZoneData Lists"
    workbook["a"].title = "LX ExpeditionZoneData Lists"
    workbook.save(filename = strippedLevelName+".xlsx")

    # XXX bodge for testing
    iExpeditionZoneDataL1.save(strippedLevelName+".xlsx", "L1 ExpeditionZoneData")
    iExpeditionZoneDataListsL1.save(strippedLevelName+".xlsx", "L1 ExpeditionZoneData Lists")

    return True

def SearchJob(desiredReverse:str, RundownDataBlock, LevelLayoutBlock, WardenObjectiveDataBlock, silent:bool=True, debug:bool=False):
    """
    Secondary job meant specifically to search for and display the search result for a level
    """
    # TODO finish the SearchJob
    pass

def main():
    # TODO get level via input or args
    # level to reverse
    desiredReverse = "Rundown 005,A,0"

    # Open Datablocks to get level from
    RundownDataBlock =  DatablockIO.datablock(open(blockpath+"RundownDataBlock.json", 'r', encoding="utf-8"))
    LevelLayoutDataBlock = DatablockIO.datablock(open(blockpath+"LevelLayoutDataBlock.json", 'r', encoding="utf8"))
    WardenObjectiveDataBlock = DatablockIO.datablock(open(blockpath+"WardenObjectiveDataBlock.json", 'r', encoding="utf8"))

    # TODO allow for program to take a list of levels to run utility on
    UtilityJob(desiredReverse, RundownDataBlock, LevelLayoutDataBlock, WardenObjectiveDataBlock, silent=False, debug=True)

    # TODO allow for a secondary job just to search for a level but not run the reverse tool on it

if __name__ == "__main__":
    main()