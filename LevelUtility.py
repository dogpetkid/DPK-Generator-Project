"""
This is a tool created by DPK

This tool can convert xlsx to a bunch of GTFO Datablock pieces and also convert levels from the Datablocks back into the templated form
Template: https://docs.google.com/spreadsheets/d/1ENa6McJnomHa5ugB-VBFjMF62nslj4VwdgM_5ERVRqw/edit?usp=sharing
"""

import argparse
import io
import json
import logging
import os
import re
import textwrap
import time
import typing

import numpy
import pandas
import xlrd

import ConfigManager
import DatablockIO
import EnumConverter
import XlsxInterfacer

# argparse: used to get arguments in CLI (to decide which files to turn into levels encoding/decoding and which file)
# io:       used to read from and write to files
# json:     used to export the data to a json
# logging:  used to create log files for the tool
# os:       used to create log directory
# re:       used to preform regex searches
# textwrap: used to help format argparse description
# time:     used to give the log file's name a timestamp and set time to gmttime
# typing:   used to give types to function parameters
# numpy:    used to manipulate the inconsistant numpy data read by pandas
# pandas:   used to read the excel data
# xlrd:     used to catch and throw excel errors when initially reading the sheets

# a regex to capture the newlines used in the sheets
sheetnewlnregex = "(\\\\r)?\\\\n"
devlf = "\n"
devcrlf = "\r\n"
# a regex to capture the tabs used in the sheets
sheettabregex = "\\\\t"
devtb = "\t"

# Settings
#####
# Version number meaning:
# R.G.S
# R: Rundown
# G: Generator
# S: Sheet (minor changes to the sheet are insignificant to the utility)
Version = ConfigManager.config["Project"]["Version"]
# relative path to location for datablocks, defaultly its folder should be on the same layer as this project's folder
blockpath = os.path.join(os.path.dirname(__file__), ConfigManager.config["Project"]["blockpath"])
# default paths to xlsx files when running the program
defaultpaths = ConfigManager.config["LevelUtility"]["defaultpaths"]
#####

def EnsureKeyInDictArray(dictionary:dict, key:str):
    """this function will ensure that an array exists in a key if there is not already a value"""
    try:_ = dictionary[key]
    except KeyError:dictionary[key] = []


# load all datablock files
if True:
    try:
        # DATABLOCK_Rundown = DatablockIO.datablock(open(os.path.join(blockpath,ConfigManager.config["Project"]["blockprefix"]+"RundownDataBlock"+ConfigManager.config["Project"]["blocksuffix"]), 'r', encoding="utf8"))
        # DATABLOCK_LevelLayout = DatablockIO.datablock(open(os.path.join(blockpath,ConfigManager.config["Project"]["blockprefix"]+"LevelLayoutDataBlock"+ConfigManager.config["Project"]["blocksuffix"]), 'r', encoding="utf8"))
        # DATABLOCK_WardenObjective = DatablockIO.datablock(open(os.path.join(blockpath,ConfigManager.config["Project"]["blockprefix"]+"WardenObjectiveDataBlock"+ConfigManager.config["Project"]["blocksuffix"]), 'r', encoding="utf8"))
        DATABLOCK_ArtifactDistributionDataBlock = DatablockIO.datablock(open(os.path.join(blockpath,ConfigManager.config["Project"]["blockprefix"]+"ArtifactDistributionDataBlock"+ConfigManager.config["Project"]["blocksuffix"]), 'r', encoding="utf8"))
        DATABLOCK_BigPickupDistribution = DatablockIO.datablock(open(os.path.join(blockpath,ConfigManager.config["Project"]["blockprefix"]+"BigPickupDistributionDataBlock"+ConfigManager.config["Project"]["blocksuffix"]), 'r', encoding="utf8"))
        DATABLOCK_ChainedPuzzle = DatablockIO.datablock(open(os.path.join(blockpath,ConfigManager.config["Project"]["blockprefix"]+"ChainedPuzzleDataBlock"+ConfigManager.config["Project"]["blocksuffix"]), 'r', encoding="utf8"))
        DATABLOCK_ComplexResourceSet = DatablockIO.datablock(open(os.path.join(blockpath,ConfigManager.config["Project"]["blockprefix"]+"ComplexResourceSetDataBlock"+ConfigManager.config["Project"]["blocksuffix"]), 'r', encoding="utf8"))
        DATABLOCK_ConsumableDistribution = DatablockIO.datablock(open(os.path.join(blockpath,ConfigManager.config["Project"]["blockprefix"]+"ConsumableDistributionDataBlock"+ConfigManager.config["Project"]["blocksuffix"]), 'r', encoding="utf8"))
        DATABLOCK_Enemy = DatablockIO.datablock(open(os.path.join(blockpath,ConfigManager.config["Project"]["blockprefix"]+"EnemyDataBlock"+ConfigManager.config["Project"]["blocksuffix"]), 'r', encoding="utf8"))
        DATABLOCK_EnemyGroup = DatablockIO.datablock(open(os.path.join(blockpath,ConfigManager.config["Project"]["blockprefix"]+"EnemyGroupDataBlock"+ConfigManager.config["Project"]["blocksuffix"]), 'r', encoding="utf8"))
        DATABLOCK_EnemyPopulation = DatablockIO.datablock(open(os.path.join(blockpath,ConfigManager.config["Project"]["blockprefix"]+"EnemyPopulationDataBlock"+ConfigManager.config["Project"]["blocksuffix"]), 'r', encoding="utf8"))
        DATABLOCK_ExpeditionBalance = DatablockIO.datablock(open(os.path.join(blockpath,ConfigManager.config["Project"]["blockprefix"]+"ExpeditionBalanceDataBlock"+ConfigManager.config["Project"]["blocksuffix"]), 'r', encoding="utf8"))
        DATABLOCK_FogSettings = DatablockIO.datablock(open(os.path.join(blockpath,ConfigManager.config["Project"]["blockprefix"]+"FogSettingsDataBlock"+ConfigManager.config["Project"]["blocksuffix"]), 'r', encoding="utf8"))
        DATABLOCK_Item = DatablockIO.datablock(open(os.path.join(blockpath,ConfigManager.config["Project"]["blockprefix"]+"ItemDataBlock"+ConfigManager.config["Project"]["blocksuffix"]), 'r', encoding="utf8"))
        DATABLOCK_LightSettings = DatablockIO.datablock(open(os.path.join(blockpath,ConfigManager.config["Project"]["blockprefix"]+"LightSettingsDataBlock"+ConfigManager.config["Project"]["blocksuffix"]), 'r', encoding="utf8"))
        DATABLOCK_StaticSpawn = DatablockIO.datablock(open(os.path.join(blockpath,ConfigManager.config["Project"]["blockprefix"]+"StaticSpawnDataBlock"+ConfigManager.config["Project"]["blocksuffix"]), 'r', encoding="utf8"))
        DATABLOCK_SurvivalWavePopulation = DatablockIO.datablock(open(os.path.join(blockpath,ConfigManager.config["Project"]["blockprefix"]+"SurvivalWavePopulationDataBlock"+ConfigManager.config["Project"]["blocksuffix"]), 'r', encoding="utf8"))
        DATABLOCK_SurvivalWaveSettings = DatablockIO.datablock(open(os.path.join(blockpath,ConfigManager.config["Project"]["blockprefix"]+"SurvivalWaveSettingsDataBlock"+ConfigManager.config["Project"]["blocksuffix"]), 'r', encoding="utf8"))
    except FileNotFoundError as e:
        if __name__ == "__main__":
            print("Missing a DataBlock: " + str(e))
            input()
        raise FileNotFoundError("Missing a DataBlock: " + str(e))

# load all enum files
if True:
    try:
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
    except FileNotFoundError:
        # use none as the enum file if the enums are not present
        ENUMFILE_LG_LayerType = None
        ENUMFILE_LG_StaticDistributionWeightType = None
        ENUMFILE_SubComplex = None
        ENUMFILE_TERM_State = None
        ENUMFILE_eEnemyGroupType = None
        ENUMFILE_eEnemyRoleDifficulty = None
        ENUMFILE_eEnemyZoneDistribution = None
        ENUMFILE_eExpeditionAccessibility = None
        ENUMFILE_eLocalZoneIndex = None
        ENUMFILE_eProgressionPuzzleType = None
        ENUMFILE_eReactorWaveSpawnType = None
        ENUMFILE_eRetrieveExitWaveTrigger = None
        ENUMFILE_eSecurityGateType = None
        ENUMFILE_eWantedZoneHeighs = None
        ENUMFILE_eWardenObjectiveEventTrigger = None
        ENUMFILE_eWardenObjectiveEventType = None
        ENUMFILE_eWardenObjectiveType = None
        ENUMFILE_eWardenObjectiveWinCondition = None
        ENUMFILE_eZoneBuildFromExpansionType = None
        ENUMFILE_eZoneBuildFromType = None
        ENUMFILE_eZoneDistributionAmount = None
        ENUMFILE_eZoneExpansionType = None

def ZonePlacementData(interface:XlsxInterfacer.interface, col:int, row:int, horizontal:bool=True):
    """
    return a ZonePlacementData dict \n
    col and row define the upper left value (not header) \n
    horizontal is true if the values are in the same row
    """
    data = {}
    interface.readIntoDict(str, col, row, data, "LocalIndex")
    EnumConverter.enumInDict(ENUMFILE_eLocalZoneIndex, data, "LocalIndex")
    data["Weights"] = ZonePlacementWeights(interface, col+horizontal, row+(not horizontal), horizontal)
    return data

def BulkheadDoorPlacementData(interface:XlsxInterfacer.interface, col:int, row:int, horizontal:bool=False):
    """
    return a FunctionPlacementData dict \n
    col and row define the upper left value (not header) \n
    horizontal is true if the values are in the same row
    """
    data = {}
    interface.readIntoDict(str, col, row, data, "ZoneIndex")
    EnumConverter.enumInDict(ENUMFILE_eLocalZoneIndex, data, "ZoneIndex")
    data["PlacementWeights"] = ZonePlacementWeights(interface, col+horizontal, row+(not horizontal), horizontal)
    interface.readIntoDict(int, col+4*horizontal, row+4*(not horizontal), data, "AreaSeedOffset")
    interface.readIntoDict(int, col+5*horizontal, row+5*(not horizontal), data, "MarkerSeedOffset")
    return data

def FunctionPlacementData(interface:XlsxInterfacer.interface, col:int, row:int, horizontal:bool=True):
    """
    return a FunctionPlacementData dict \n
    col and row define the upper left value (not header) \n
    horizontal is true if the values are in the same row
    """
    data = {}
    data["PlacementWeights"] = ZonePlacementWeights(interface, col, row, horizontal)
    interface.readIntoDict(int, col+3*horizontal, row+3*(not horizontal), data, "AreaSeedOffset")
    interface.readIntoDict(int, col+4*horizontal, row+4*(not horizontal), data, "MarkerSeedOffset")
    return data

def ZonePlacementWeightsList(interface:XlsxInterfacer.interface, col:int, row:int, horizontal:bool=True):
    """
    return a list of lists containing ZonePlacementWeights \n
    col and row define the upper left value (not header) \n
    horizontal describes the direction to iterate to look for weights
    """
    data = {}
    while not(interface.isEmpty(col, row)):
        Snippet = {}
        identifier = interface.read(str, col, row)
        interface.readIntoDict(str, col+(not horizontal), row+horizontal, Snippet, "LocalIndex")
        EnumConverter.enumInDict(ENUMFILE_eLocalZoneIndex, Snippet, "LocalIndex")
        # the direction of the set of weights and values in the data are perpendicular
        Snippet["Weights"] = ZonePlacementWeights(interface, col+2*(not horizontal), row+2*horizontal, horizontal=(not horizontal))
        EnsureKeyInDictArray(data, identifier)
        data[identifier].append(Snippet)
        col+= horizontal
        row+= not horizontal
    return list(data.values())

def ZonePlacementWeights(interface:XlsxInterfacer.interface, col:int, row:int, horizontal:bool=True):
    """
    return a ZonePlacementWeights dict \n
    col and row define the upper left value (not header) \n
    horizontal is true if the values are in the same row
    """
    weights = {}
    interface.readIntoDict(float, col, row, weights, "Start")
    interface.readIntoDict(float, col+horizontal, row+(not horizontal), weights, "Middle")
    interface.readIntoDict(float, col+2*horizontal, row+2*(not horizontal), weights, "End")
    return weights

def GenericEnemyWaveData(interface:XlsxInterfacer.interface, col:int, row:int, horizontal:bool=False):
    """
    return a GenericEnemyWaveData dict \n
    col and row define the upper left value (not header) \n
    horizontal is true if the values are in the same row
    """
    data = {}
    interface.readIntoDict(str, col, row, data, "WaveSettings")
    DatablockIO.nameInDict(DATABLOCK_SurvivalWaveSettings, data, "WaveSettings")
    interface.readIntoDict(str, col+horizontal, row+(not horizontal), data, "WavePopulation")
    DatablockIO.nameInDict(DATABLOCK_SurvivalWavePopulation, data, "WavePopulation")
    interface.readIntoDict(float, col+2*horizontal, row+2*(not horizontal), data, "SpawnDelay")
    interface.readIntoDict(bool, col+3*horizontal, row+3*(not horizontal), data, "TriggerAlarm")
    interface.readIntoDict(str, col+4*horizontal, row+4*(not horizontal), data, "IntelMessage")
    return data

def GenericEnemyWaveDataList(interface:XlsxInterfacer.interface, col:int, row:int, horizontal:bool=True):
    """
    return a GenericEnemyWaveData dict \n
    col and row define the upper left value (not header) \n
    horizontal describes the direction to iterate to look for waves
    """
    datalist = []
    while not(interface.isEmpty(col,row) and interface.isEmpty(col+2*horizontal,row+2*(not horizontal))):
        # the direction of the set of waves and values in the data are perpendicular
        datalist.append(GenericEnemyWaveData(interface, col, row, horizontal=(not horizontal)))
        col+= horizontal
        row+= not horizontal
    return datalist

def ArtifactZoneDistribution(interface:XlsxInterfacer.interface, col:int, row:int, horizontal:bool=False):
    """
    return a ArtifactZoneDistribution dict \n
    col and row define the upper left value (not header) \n
    horizontal is true if the values are in the same row
    """
    data = {}
    interface.readIntoDict(str, col, row, data, "Zone")
    EnumConverter.enumInDict(ENUMFILE_eLocalZoneIndex, data, "Zone")
    interface.readIntoDict(float, col+horizontal, row+(not horizontal), data, "BasicArtifactWeight")
    interface.readIntoDict(float, col+2*horizontal, row+2*(not horizontal), data, "AdvancedArtifactWeight")
    interface.readIntoDict(float, col+3*horizontal, row+3*(not horizontal), data, "SpecializedArtifactWeight")
    return data

def WardenObjectiveEventData(interface:XlsxInterfacer.interface, col:int, row:int, horizontal:bool=False):
    """
    return a WardenObjectiveEventData dict \n
    col and row define the upper left value (not header) \n
    horizontal is true if the values are in the same row
    """
    data = {}
    interface.readIntoDict(str, col, row, data, "Trigger")
    EnumConverter.enumInDict(ENUMFILE_eWardenObjectiveEventTrigger, data, "Trigger")
    interface.readIntoDict(str, col+horizontal, row+(not horizontal), data, "Type")
    EnumConverter.enumInDict(ENUMFILE_eWardenObjectiveEventType, data, "Type")
    interface.readIntoDict(str, col+2*horizontal, row+2*(not horizontal), data, "Layer")
    EnumConverter.enumInDict(ENUMFILE_LG_LayerType, data, "Layer")
    interface.readIntoDict(str, col+3*horizontal, row+3*(not horizontal), data, "LocalIndex")
    EnumConverter.enumInDict(ENUMFILE_eLocalZoneIndex, data, "LocalIndex")
    interface.readIntoDict(float, col+4*horizontal, row+4*(not horizontal), data, "Delay")
    interface.readIntoDict(str, col+5*horizontal, row+5*(not horizontal), data, "WardenIntel")
    interface.readIntoDict(int, col+6*horizontal, row+6*(not horizontal), data, "SoundID")
    # TODO convert sound placeholders
    return data

def GeneralFogDataStep(interface:XlsxInterfacer.interface, col:int, row:int, horizontal:bool=False):
    """
    return a GeneralFogDataStep dict \n
    col and row define the upper left value (not header) \n
    horizontal is true if the values are in the same row
    """
    data = {}
    interface.readIntoDict(str, col, row, data, "m_fogDataId")
    DatablockIO.nameInDict(DATABLOCK_FogSettings, data, "m_fogDataId")
    interface.readIntoDict(float, col+horizontal, row+(not horizontal), data, "m_transitionToTime")
    return data


def LayerData(interface:XlsxInterfacer.interface, col:int, row:int):
    """
    return a LayerData dict \n
    col and row define the upper left value, SHOULD be the header \n
    horizontal is true if the values are in the same row
    """
    data = {}
    interface.readIntoDict(int, col+5, row, data, "ZoneAliasStart")
    data["ZonesWithBulkheadEntrance"] = []
    itercol,iterrow = col+5, row+1
    while not(interface.isEmpty(itercol, iterrow)):
        # NOTE textmode may need a toggle in this file for whether the json should have text enums
        data["ZonesWithBulkheadEntrance"].append(EnumConverter.enumToIndex(ENUMFILE_eLocalZoneIndex, interface.read(str, itercol, iterrow), textmode=True))
        itercol+= 1
    data["BulkheadDoorControllerPlacements"] = []
    itercol,iterrow = col+5, row+2
    while not(interface.isEmpty(itercol, iterrow)):
        data["BulkheadDoorControllerPlacements"].append(BulkheadDoorPlacementData(interface, itercol, iterrow, horizontal=False))
        itercol+= 1
    data["BulkheadKeyPlacements"] = ZonePlacementWeightsList(interface, col+5, row+8, horizontal=True)
    data["ObjectiveData"] = {}
    interface.readIntoDict(int, col+5, row+13, data["ObjectiveData"], "DataBlockId")
    interface.readIntoDict(str, col+5, row+14, data["ObjectiveData"], "WinCondition")
    EnumConverter.enumInDict(ENUMFILE_eWardenObjectiveWinCondition, data["ObjectiveData"], "WinCondition")
    data["ObjectiveData"]["ZonePlacementDatas"] = ZonePlacementWeightsList(interface, col+5, row+15, horizontal=True)
    data["ArtifactData"] = {}
    interface.readIntoDict(float, col+5, row+20, data["ArtifactData"], "ArtifactAmountMulti")
    interface.readIntoDict(str, col+5, row+21, data["ArtifactData"], "ArtifactLayerDistributionDataID")
    DatablockIO.nameInDict(DATABLOCK_ArtifactDistributionDataBlock, data["ArtifactData"], "ArtifactLayerDistributionDataID")
    data["ArtifactData"]["ArtifactZoneDistributions"] = []
    itercol,iterrow = col+5, row+22
    while not(interface.isEmpty(itercol, iterrow)):
        data["ArtifactData"]["ArtifactZoneDistributions"].append(ArtifactZoneDistribution(interface, itercol, iterrow, horizontal=False))
        itercol+= 1
    return data

def ExpeditionInTier(iExpeditionInTier:XlsxInterfacer.interface):
    """returns the expedition in tier piece to be inserted into the rundown data block"""
    data = {}
    data["Enabled"] = iExpeditionInTier.read(bool, 0, 2)
    data["Accessibility"] = iExpeditionInTier.read(str, 1, 2)
    EnumConverter.enumInDict(ENUMFILE_eExpeditionAccessibility, data, "Accessibility")
    data["CustomProgressionLock"] = {}
    iExpeditionInTier.readIntoDict(int, 12, 2, data["CustomProgressionLock"], "MainSectors")
    iExpeditionInTier.readIntoDict(int, 12, 3, data["CustomProgressionLock"], "SecondarySectors")
    iExpeditionInTier.readIntoDict(int, 12, 4, data["CustomProgressionLock"], "ThirdSectors")
    iExpeditionInTier.readIntoDict(int, 12, 5, data["CustomProgressionLock"], "AllClearedSectors")
    data["Descriptive"] = {}
    data["Descriptive"]["Prefix"] = iExpeditionInTier.read(str, 12, 8)
    data["Descriptive"]["PublicName"] = iExpeditionInTier.read(str, 12, 9)
    iExpeditionInTier.readIntoDict(bool, 12, 10, data["Descriptive"], "IsExtraExpedition")
    iExpeditionInTier.readIntoDict(int, 12, 11, data["Descriptive"], "ExpeditionDepth")
    data["Descriptive"]["EstimatedDuration"] = iExpeditionInTier.read(XlsxInterfacer.blankable, 12, 12)
    data["Descriptive"]["ExpeditionDescription"] = re.sub(sheetnewlnregex, devcrlf, iExpeditionInTier.read(XlsxInterfacer.blankable, 12, 13))
    data["Descriptive"]["RoleplayedWardenIntel"] = re.sub(sheetnewlnregex, devcrlf, iExpeditionInTier.read(XlsxInterfacer.blankable, 12, 14))
    data["Descriptive"]["DevInfo"] = re.sub(sheetnewlnregex, devlf, iExpeditionInTier.read(XlsxInterfacer.blankable, 12, 15))
    data["Seeds"] = {}
    iExpeditionInTier.readIntoDict(int, 0, 6, data["Seeds"], "BuildSeed")
    iExpeditionInTier.readIntoDict(int, 1, 6, data["Seeds"], "FunctionMarkerOffset")
    iExpeditionInTier.readIntoDict(int, 2, 6, data["Seeds"], "StandardMarkerOffset")
    iExpeditionInTier.readIntoDict(int, 3, 6, data["Seeds"], "LightJobSeedOffset")
    data["Expedition"] = {}
    iExpeditionInTier.readIntoDict(str, 0, 10, data["Expedition"], "ComplexResourceData")
    DatablockIO.nameInDict(DATABLOCK_ComplexResourceSet, data["Expedition"], "ComplexResourceData")
    iExpeditionInTier.readIntoDict(str, 1, 10, data["Expedition"], "LightSettings")
    DatablockIO.nameInDict(DATABLOCK_LightSettings, data["Expedition"], "LightSettings")
    iExpeditionInTier.readIntoDict(str, 2, 10, data["Expedition"], "FogSettings")
    DatablockIO.nameInDict(DATABLOCK_FogSettings, data["Expedition"], "FogSettings")
    iExpeditionInTier.readIntoDict(str, 3, 10, data["Expedition"], "EnemyPopulation")
    DatablockIO.nameInDict(DATABLOCK_EnemyPopulation, data["Expedition"], "EnemyPopulation")
    iExpeditionInTier.readIntoDict(str, 4, 10, data["Expedition"], "ExpeditionBalance")
    DatablockIO.nameInDict(DATABLOCK_ExpeditionBalance, data["Expedition"], "ExpeditionBalance")
    iExpeditionInTier.readIntoDict(str, 5, 10, data["Expedition"], "ScoutWaveSettings")
    DatablockIO.nameInDict(DATABLOCK_SurvivalWaveSettings, data["Expedition"], "ScoutWaveSettings")
    iExpeditionInTier.readIntoDict(str, 6, 10, data["Expedition"], "ScoutWavePopulation")
    DatablockIO.nameInDict(DATABLOCK_SurvivalWavePopulation, data["Expedition"], "ScoutWavePopulation")
    data["LevelLayoutData"] = iExpeditionInTier.read(int, 0, 13)
    data["MainLayerData"] = LayerData(iExpeditionInTier, 0, 20)
    iExpeditionInTier.readIntoDict(bool, 3, 13, data, "SecondaryLayerEnabled")
    iExpeditionInTier.readIntoDict(int, 4, 13, data, "SecondaryLayout")
    data["BuildSecondaryFrom"] = {}
    iExpeditionInTier.readIntoDict(str, 3, 17, data["BuildSecondaryFrom"], "LayerType")
    EnumConverter.enumInDict(ENUMFILE_LG_LayerType, data["BuildSecondaryFrom"], "LayerType")
    iExpeditionInTier.readIntoDict(str, 4, 17, data["BuildSecondaryFrom"], "Zone")
    EnumConverter.enumInDict(ENUMFILE_eLocalZoneIndex, data["BuildSecondaryFrom"], "Zone")
    data["SecondaryLayerData"] = LayerData(iExpeditionInTier, 0, 48)
    iExpeditionInTier.readIntoDict(bool, 7, 13, data, "ThirdLayerEnabled")
    iExpeditionInTier.readIntoDict(int, 8, 13, data, "ThirdLayout")
    data["BuildThirdFrom"] = {}
    iExpeditionInTier.readIntoDict(str, 7, 17, data["BuildThirdFrom"], "LayerType")
    EnumConverter.enumInDict(ENUMFILE_LG_LayerType, data["BuildThirdFrom"], "LayerType")
    iExpeditionInTier.readIntoDict(str, 8, 17, data["BuildThirdFrom"], "Zone")
    EnumConverter.enumInDict(ENUMFILE_eLocalZoneIndex, data["BuildThirdFrom"], "Zone")
    data["ThirdLayerData"] = LayerData(iExpeditionInTier, 0, 76)
    data["SpecialOverrideData"] = {}
    iExpeditionInTier.readIntoDict(float, 5, 6, data["SpecialOverrideData"], "WeakResourceContainerWithPackChanceForLocked")
    return data

def LevelLayoutBlock(iExpeditionZoneData:XlsxInterfacer.interface, iExpeditionZoneDataLists:XlsxInterfacer.interface):
    """returns a Level Layout block (name, internalEnabled, and persistentID are set to defaults as their data comes from elsewhere)"""
    data = {}
    data["Zones"] = []
    listdata = ExpeditionZoneDataLists(iExpeditionZoneDataLists)

    row = 2
    while not (iExpeditionZoneData.isEmpty(0, row)):
        data["Zones"].append(ExpeditionZoneData(iExpeditionZoneData, listdata, row))
        row+= 1

    # Set values to be filled
    data["name"] = "DPK Utility Generated Layout"
    data["internalEnabled"] = False
    data["persistentID"] = -1

    return data

class ExpeditionZoneDataLists:
    """a class that increases the dimentions of the dictionaries in ExpeditionZoneData (since the sheet cannot contain 2d-3d data)"""

    def __init__(self, iExpeditionZoneDataLists:XlsxInterfacer.interface):
        """Generates numerous stubs that can have zone specific data request from the object getters"""
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

        self.stubEventsOnEnter = {}
        self.stubProgressionPuzzleToEnter = {}
        self.stubEnemySpawningInZone = {}
        self.stubEnemyRespawnExcludeList = {}
        self.stubTerminalPlacements = {}
        self.stubLocalLogFiles = {}
        self.stubParsedLog = {}
        self.stubPowerGeneratorPlacements = {}
        self.stubDisinfectionStationPlacements = {}
        self.stubStaticSpawnDataContainers = {}
        
        row = startrow
        # EventsOnEnter
        while not(iExpeditionZoneDataLists.isEmpty(startcolEventsOnEnter,row)):
            Snippet = {}
            iExpeditionZoneDataLists.readIntoDict(float, startcolEventsOnEnter+1, row, Snippet, "Delay")
            Snippet["Noise"] = {}
            iExpeditionZoneDataLists.readIntoDict(bool, startcolEventsOnEnter+2, row, Snippet["Noise"], "Enabled")
            iExpeditionZoneDataLists.readIntoDict(float, startcolEventsOnEnter+3, row, Snippet["Noise"], "RadiusMin")
            iExpeditionZoneDataLists.readIntoDict(float, startcolEventsOnEnter+4, row, Snippet["Noise"], "RadiusMax")
            Snippet["Intel"] = {}
            iExpeditionZoneDataLists.readIntoDict(bool, startcolEventsOnEnter+5, row, Snippet["Intel"], "Enabled")
            iExpeditionZoneDataLists.readIntoDict(str, startcolEventsOnEnter+6, row, Snippet["Intel"], "IntelMessage")
            Snippet["Sound"] = {}
            iExpeditionZoneDataLists.readIntoDict(bool, startcolEventsOnEnter+7, row, Snippet["Sound"], "Enabled")
            iExpeditionZoneDataLists.readIntoDict(int, startcolEventsOnEnter+8, row, Snippet["Sound"], "SoundEvent")
            # TODO convert sound placeholders
            EnsureKeyInDictArray(self.stubEventsOnEnter, iExpeditionZoneDataLists.read(str, startcolEventsOnEnter, row))
            self.stubEventsOnEnter[iExpeditionZoneDataLists.read(str, startcolEventsOnEnter, row)].append(Snippet)
            row+= 1

        row = startrow
        # ProgressionPuzzleToEnter
        while not(iExpeditionZoneDataLists.isEmpty(startcolProgressionPuzzleToEnter,row)):
            Snippet = ZonePlacementData(iExpeditionZoneDataLists, startcolProgressionPuzzleToEnter+2,row, horizontal=True)
            EnsureKeyInDictArray(self.stubProgressionPuzzleToEnter, iExpeditionZoneDataLists.read(str, startcolProgressionPuzzleToEnter, row))
            self.stubProgressionPuzzleToEnter[iExpeditionZoneDataLists.read(str, startcolProgressionPuzzleToEnter, row)].append(Snippet)
            row+= 1

        row = startrow
        # EnemySpawningInZone
        while not(iExpeditionZoneDataLists.isEmpty(startcolEnemySpawningInZone,row)):
            Snippet = {}
            iExpeditionZoneDataLists.readIntoDict(str, startcolEnemySpawningInZone+2, row, Snippet, "GroupType")
            EnumConverter.enumInDict(ENUMFILE_eEnemyGroupType, Snippet, "GroupType")
            iExpeditionZoneDataLists.readIntoDict(str, startcolEnemySpawningInZone+3, row, Snippet, "Difficulty")
            EnumConverter.enumInDict(ENUMFILE_eEnemyRoleDifficulty, Snippet, "Difficulty")
            iExpeditionZoneDataLists.readIntoDict(str, startcolEnemySpawningInZone+4, row, Snippet, "Distribution")
            EnumConverter.enumInDict(ENUMFILE_eEnemyZoneDistribution, Snippet, "Distribution")
            iExpeditionZoneDataLists.readIntoDict(float, startcolEnemySpawningInZone+5, row, Snippet, "DistributionValue")
            EnsureKeyInDictArray(self.stubEnemySpawningInZone, iExpeditionZoneDataLists.read(str, startcolEnemySpawningInZone, row))
            self.stubEnemySpawningInZone[iExpeditionZoneDataLists.read(str, startcolEnemySpawningInZone, row)].append(Snippet)
            row+= 1

        row = startrow
        # EnemyRespawnExcludeList
        while not(iExpeditionZoneDataLists.isEmpty(startcolEnemyRespawnExcludeList,row)):
            Snippet = {}
            # using a key named value saves rewriting several try except statements here
            iExpeditionZoneDataLists.readIntoDict(str, startcolEnemyRespawnExcludeList+1, row, Snippet, "value")
            DatablockIO.nameInDict(DATABLOCK_Enemy, Snippet, "value") # using nameInDict saves writing more try statements here
            EnsureKeyInDictArray(self.stubEnemyRespawnExcludeList, iExpeditionZoneDataLists.read(str, startcolEnemyRespawnExcludeList, row))
            try:
                # if the key for snippet does not exist in the dictionary, then it was not present
                self.stubEnemyRespawnExcludeList[iExpeditionZoneDataLists.read(str, startcolEnemyRespawnExcludeList, row)].append(Snippet["value"])
            except KeyError:
                pass
            row+= 1

        row = startrow
        # ParsedLog
        while not(iExpeditionZoneDataLists.isEmpty(startcolParsedLog,row)):
            Snippet = iExpeditionZoneDataLists.read(XlsxInterfacer.blankable, startcolParsedLog+1,row)
            EnsureKeyInDictArray(self.stubParsedLog, iExpeditionZoneDataLists.read(str, startcolParsedLog, row))
            self.stubLocalLogFiles[iExpeditionZoneDataLists.read(str, startcolParsedLog, row)].append(Snippet)
            row+= 1

        row = startrow
        # LocalLogFiles
        while not(iExpeditionZoneDataLists.isEmpty(startcolLocalLogFiles,row)):
            Snippet = {}
            iExpeditionZoneDataLists.readIntoDict(str, startcolLocalLogFiles+2, row, Snippet, "FileName")
            iExpeditionZoneDataLists.readIntoDict(str, startcolLocalLogFiles+3, row, Snippet, "FileContent")
            try:
                Snippet["FileContent"] = re.sub(sheetnewlnregex, devcrlf, Snippet["FileContent"])
                Snippet["FileContent"] = re.sub(sheettabregex, devtb, Snippet["FileContent"])
            except KeyError:pass
            iExpeditionZoneDataLists.readIntoDict(int, startcolLocalLogFiles+4, row, Snippet, "AttachedAudioFile")
            # TODO convert sound placeholders
            iExpeditionZoneDataLists.readIntoDict(int, startcolLocalLogFiles+5, row, Snippet, "AttachedAudioByteSize")
            iExpeditionZoneDataLists.readIntoDict(int, startcolLocalLogFiles+6, row, Snippet, "PlayerDialogToTriggerAfterAudio")
            # TODO convert sound placeholders
            EnsureKeyInDictArray(self.stubLocalLogFiles, iExpeditionZoneDataLists.read(str, startcolLocalLogFiles, row))
            self.stubLocalLogFiles[iExpeditionZoneDataLists.read(str, startcolLocalLogFiles, row)].append(Snippet)
            row+= 1

        row = startrow
        # TerminalPlacements
        while not(iExpeditionZoneDataLists.isEmpty(startcolTerminalPlacements,row)):
            Snippet = {}
            Snippet["PlacementWeights"] = ZonePlacementWeights(iExpeditionZoneDataLists, startcolTerminalPlacements+1, row, horizontal=True)
            iExpeditionZoneDataLists.readIntoDict(int, startcolTerminalPlacements+4, row, Snippet, "AreaSeedOffset")
            iExpeditionZoneDataLists.readIntoDict(int, startcolTerminalPlacements+5, row, Snippet, "MarkerSeedOffset")
            Snippet["LocalLogFiles"] = self.LocalLogFiles(iExpeditionZoneDataLists.read(XlsxInterfacer.blankable, startcolTerminalPlacements+6, row))
            Snippet["StartingStateData"] = {}
            iExpeditionZoneDataLists.readIntoDict(str, startcolTerminalPlacements+7, row, Snippet["StartingStateData"], "StartingState")
            EnumConverter.enumInDict(ENUMFILE_TERM_State, Snippet["StartingStateData"], "StartingState")
            iExpeditionZoneDataLists.readIntoDict(int, startcolTerminalPlacements+8, row, Snippet["StartingStateData"], "AudioEventEnter")
            iExpeditionZoneDataLists.readIntoDict(int, startcolTerminalPlacements+9, row, Snippet["StartingStateData"], "AudioEventExit")
            # TODO convert sound placeholders
            EnsureKeyInDictArray(self.stubTerminalPlacements, iExpeditionZoneDataLists.read(str, startcolTerminalPlacements, row))
            self.stubTerminalPlacements[iExpeditionZoneDataLists.read(str, startcolTerminalPlacements, row)].append(Snippet)
            row+= 1

        row = startrow
        # PowerGeneratorPlacements
        while not(iExpeditionZoneDataLists.isEmpty(startcolPowerGeneratorPlacements,row)):
            Snippet = {}
            Snippet.update(FunctionPlacementData(iExpeditionZoneDataLists, startcolPowerGeneratorPlacements+1, row, horizontal=True))
            EnsureKeyInDictArray(self.stubPowerGeneratorPlacements, iExpeditionZoneDataLists.read(str, startcolPowerGeneratorPlacements, row))
            self.stubPowerGeneratorPlacements[iExpeditionZoneDataLists.read(str, startcolPowerGeneratorPlacements, row)].append(Snippet)
            row+= 1

        row = startrow
        # DisinfectionStationPlacements
        while not(iExpeditionZoneDataLists.isEmpty(startcolDisinfectionStationPlacements,row)):
            Snippet = {}
            Snippet.update(FunctionPlacementData(iExpeditionZoneDataLists, startcolDisinfectionStationPlacements+1, row, horizontal=True))
            EnsureKeyInDictArray(self.stubDisinfectionStationPlacements, iExpeditionZoneDataLists.read(str, startcolDisinfectionStationPlacements, row))
            self.stubDisinfectionStationPlacements[iExpeditionZoneDataLists.read(str, startcolDisinfectionStationPlacements, row)].append(Snippet)
            row+= 1

        row = startrow
        # StaticSpawnDataContainers
        while not(iExpeditionZoneDataLists.isEmpty(startcolStaticSpawnDataContainers,row)):
            Snippet = {}
            iExpeditionZoneDataLists.readIntoDict(int, startcolStaticSpawnDataContainers+1, row, Snippet, "Count")
            iExpeditionZoneDataLists.readIntoDict(str, startcolStaticSpawnDataContainers+2, row, Snippet, "DistributionWeightType")
            EnumConverter.enumInDict(ENUMFILE_LG_StaticDistributionWeightType, Snippet, "DistributionWeightType")
            iExpeditionZoneDataLists.readIntoDict(float, startcolStaticSpawnDataContainers+3, row, Snippet, "DistributionWeight")
            iExpeditionZoneDataLists.readIntoDict(float, startcolStaticSpawnDataContainers+4, row, Snippet, "DistributionRandomBlend")
            iExpeditionZoneDataLists.readIntoDict(float, startcolStaticSpawnDataContainers+5, row, Snippet, "DistributionResultPow")
            iExpeditionZoneDataLists.readIntoDict(str, startcolStaticSpawnDataContainers+6, row, Snippet, "StaticSpawnDataId")
            DatablockIO.nameInDict(DATABLOCK_StaticSpawn, Snippet, "StaticSpawnDataId")
            iExpeditionZoneDataLists.readIntoDict(int, startcolStaticSpawnDataContainers+7, row, Snippet, "FixedSeed")
            EnsureKeyInDictArray(self.stubStaticSpawnDataContainers, iExpeditionZoneDataLists.read(str, startcolStaticSpawnDataContainers, row))
            self.stubStaticSpawnDataContainers[iExpeditionZoneDataLists.read(str, startcolStaticSpawnDataContainers, row)].append(Snippet)
            row+= 1


    def EventsOnEnter(self, identifier:str):
        """returns the EventsOnEnter array for a specific zone"""
        try:return self.stubEventsOnEnter[identifier]
        except KeyError:pass
        return []

    def ProgressionPuzzleToEnterZonePlacementData(self, identifier:str):
        """returns the ZonePlacementData for the ProgressionPuzzleToEnter for a specific zone"""
        try:return self.stubProgressionPuzzleToEnter[identifier]
        except KeyError:pass
        return []

    def EnemySpawningInZone(self, identifier:str):
        """returns the EnemySpawningInZone array for a specific zone"""
        try:return self.stubEnemySpawningInZone[identifier]
        except KeyError:pass
        return []

    def EnemyRespawnExcludeList(self, identifier:str):
        """returns the EnemyRespawnExcludeList array for a specific zone"""
        try:return self.stubEnemyRespawnExcludeList[identifier]
        except KeyError:pass
        return []

    def TerminalPlacements(self, identifier:str):
        """returns the TerminalPlacements array for a specific zone"""
        try:return self.stubTerminalPlacements[identifier]
        except KeyError:pass
        return []

    def LocalLogFiles(self, group:str):
        """returns the LocalLogFiles array for a specific grouping to be used in the TerminalPlacements"""
        try:return self.stubLocalLogFiles[group]
        except KeyError:pass
        return []

    def PowerGeneratorPlacements(self, identifier:str):
        """returns the ZonePlacementWeights for the PowerGeneratorPlacements for a specific zone"""
        try:return self.stubPowerGeneratorPlacements[identifier]
        except KeyError:pass
        return []
    
    def DisinfectionStationPlacements(self, identifier:str):
        """returns the ZonePlacementWeights for the DisinfectionStationPlacements for a specific zone"""
        try:return self.stubDisinfectionStationPlacements[identifier]
        except KeyError:pass
        return []

    def StaticSpawnDataContainers(self, identifier:str):
        """returns the StaticSpawnDataContainers array for a specific zone"""
        try:return self.stubStaticSpawnDataContainers[identifier]
        except KeyError:pass
        return []

def ExpeditionZoneData(iExpeditionZoneData:XlsxInterfacer.interface, listdata:ExpeditionZoneDataLists, row:int):
    """returns the ExpeditionZoneData for a particular row"""
    # set up some checkpoints so if some of the data gets reformatted, not the entire function needs to be altered,
    # just the headings and contents of the section will need edited column values
    colPuzzleType = XlsxInterfacer.ctn("Q")
    colHSUClustersInZone = XlsxInterfacer.ctn("AH")
    colHealthMulti = XlsxInterfacer.ctn("AY")

    data = {}

    zonestr = iExpeditionZoneData.read(str, 0, row)
    # NOTE textmode may need a toggle in this file for whether the json should have text enums
    data["LocalIndex"] =  EnumConverter.enumToIndex(ENUMFILE_eLocalZoneIndex, zonestr, textmode=True)

    iExpeditionZoneData.readIntoDict(int, 1, row, data, "SubSeed")
    iExpeditionZoneData.readIntoDict(int, 2, row, data, "BulkheadDCScanSeed")
    iExpeditionZoneData.readIntoDict(str, 3, row, data, "SubComplex")
    EnumConverter.enumInDict(ENUMFILE_SubComplex, data, "SubComplex")
    iExpeditionZoneData.readIntoDict(str, 4, row, data, "CustomGeomorph")
    data["CoverageMinMax"] = {}
    data["CoverageMinMax"]["x"] = iExpeditionZoneData.read(float, 5, row)
    data["CoverageMinMax"]["y"] = iExpeditionZoneData.read(float, 6, row)
    iExpeditionZoneData.readIntoDict(str, 7, row, data, "BuildFromLocalIndex")
    EnumConverter.enumInDict(ENUMFILE_eLocalZoneIndex, data, "BuildFromLocalIndex")
    iExpeditionZoneData.readIntoDict(str, 8, row, data, "StartPosition")
    EnumConverter.enumInDict(ENUMFILE_eZoneBuildFromType, data, "StartPosition")
    iExpeditionZoneData.readIntoDict(float, 9, row, data, "StartPosition_IndexWeight")
    iExpeditionZoneData.readIntoDict(str, 10, row, data, "StartExpansion")
    EnumConverter.enumInDict(ENUMFILE_eZoneBuildFromExpansionType, data, "StartExpansion")
    iExpeditionZoneData.readIntoDict(str, 11, row, data, "ZoneExpansion")
    EnumConverter.enumInDict(ENUMFILE_eZoneExpansionType, data, "ZoneExpansion")
    iExpeditionZoneData.readIntoDict(str, 12, row, data, "LightSettings")
    DatablockIO.nameInDict(DATABLOCK_LightSettings, data, "LightSettings")
    data["AltitudeData"] = {}
    iExpeditionZoneData.readIntoDict(str, 13, row, data["AltitudeData"], "AllowedZoneAltitude")
    EnumConverter.enumInDict(ENUMFILE_eWantedZoneHeighs, data["AltitudeData"], "AllowedZoneAltitude")
    iExpeditionZoneData.readIntoDict(float, 14, row, data["AltitudeData"], "ChanceToChange")
    data["EventsOnEnter"] = listdata.EventsOnEnter(zonestr)

    data["ProgressionPuzzleToEnter"] = {}
    iExpeditionZoneData.readIntoDict(str, colPuzzleType, row, data["ProgressionPuzzleToEnter"], "PuzzleType")
    EnumConverter.enumInDict(ENUMFILE_eProgressionPuzzleType, data["ProgressionPuzzleToEnter"], "PuzzleType")
    iExpeditionZoneData.readIntoDict(str, colPuzzleType+1, row, data["ProgressionPuzzleToEnter"], "CustomText")
    iExpeditionZoneData.readIntoDict(int, colPuzzleType+2, row, data["ProgressionPuzzleToEnter"], "PlacementCount")
    data["ProgressionPuzzleToEnter"]["ZonePlacementData"] = listdata.ProgressionPuzzleToEnterZonePlacementData(zonestr)
    iExpeditionZoneData.readIntoDict(str, colPuzzleType+4, row, data, "ChainedPuzzleToEnter")
    DatablockIO.nameInDict(DATABLOCK_ChainedPuzzle, data, "ChainedPuzzleToEnter")
    iExpeditionZoneData.readIntoDict(str, colPuzzleType+5, row, data, "SecurityGateToEnter")
    EnumConverter.enumInDict(ENUMFILE_eSecurityGateType, data, "SecurityGateToEnter")
    data["ActiveEnemyWave"] = {}
    iExpeditionZoneData.readIntoDict(bool, colPuzzleType+6, row, data["ActiveEnemyWave"], "HasActiveEnemyWave")
    iExpeditionZoneData.readIntoDict(str, colPuzzleType+7, row, data["ActiveEnemyWave"], "EnemyGroupInfrontOfDoor")
    DatablockIO.nameInDict(DATABLOCK_EnemyGroup, data["ActiveEnemyWave"], "EnemyGroupInfrontOfDoor")
    iExpeditionZoneData.readIntoDict(str, colPuzzleType+8, row, data["ActiveEnemyWave"], "EnemyGroupInArea")
    DatablockIO.nameInDict(DATABLOCK_EnemyGroup, data["ActiveEnemyWave"], "EnemyGroupInArea")
    iExpeditionZoneData.readIntoDict(int, colPuzzleType+9, row, data["ActiveEnemyWave"], "EnemyGroupsInArea")
    data["EnemySpawningInZone"] = listdata.EnemySpawningInZone(zonestr)
    iExpeditionZoneData.readIntoDict(bool, colPuzzleType+11, row, data, "EnemyRespawning")
    iExpeditionZoneData.readIntoDict(bool, colPuzzleType+12, row, data, "EnemyRespawnRequireOtherZone")
    iExpeditionZoneData.readIntoDict(int, colPuzzleType+13, row, data, "EnemyRespawnRoomDistance")
    iExpeditionZoneData.readIntoDict(float, colPuzzleType+14, row, data, "EnemyRespawnTimeInterval")
    iExpeditionZoneData.readIntoDict(float, colPuzzleType+15, row, data, "EnemyRespawnCountMultiplier")
    data["EnemyRespawnExcludeList"] = listdata.EnemyRespawnExcludeList(zonestr)

    iExpeditionZoneData.readIntoDict(int, colHSUClustersInZone, row, data, "HSUClustersInZone")
    iExpeditionZoneData.readIntoDict(int, colHSUClustersInZone+1, row, data, "CorpseClustersInZone")
    iExpeditionZoneData.readIntoDict(int, colHSUClustersInZone+2, row, data, "ResourceContainerClustersInZone")
    iExpeditionZoneData.readIntoDict(int, colHSUClustersInZone+3, row, data, "GeneratorClustersInZone")
    iExpeditionZoneData.readIntoDict(str, colHSUClustersInZone+4, row, data, "CorpsesInZone")
    EnumConverter.enumInDict(ENUMFILE_eZoneDistributionAmount, data, "CorpsesInZone")
    iExpeditionZoneData.readIntoDict(str, colHSUClustersInZone+5, row, data, "GroundSpawnersInZone")
    EnumConverter.enumInDict(ENUMFILE_eZoneDistributionAmount, data, "GroundSpawnersInZone")
    iExpeditionZoneData.readIntoDict(str, colHSUClustersInZone+6, row, data, "HSUsInZone")
    EnumConverter.enumInDict(ENUMFILE_eZoneDistributionAmount, data, "HSUsInZone")
    iExpeditionZoneData.readIntoDict(str, colHSUClustersInZone+7, row, data, "DeconUnitsInZone")
    EnumConverter.enumInDict(ENUMFILE_eZoneDistributionAmount, data, "DeconUnitsInZone")
    iExpeditionZoneData.readIntoDict(bool, colHSUClustersInZone+8, row, data, "AllowSmallPickupsAllocation")
    iExpeditionZoneData.readIntoDict(bool, colHSUClustersInZone+9, row, data, "AllowResourceContainerAllocation")
    iExpeditionZoneData.readIntoDict(bool, colHSUClustersInZone+10, row, data, "ForceBigPickupsAllocation")
    iExpeditionZoneData.readIntoDict(str, colHSUClustersInZone+11, row, data, "ConsumableDistributionInZone")
    DatablockIO.nameInDict(DATABLOCK_ConsumableDistribution, data, "ConsumableDistributionInZone")
    iExpeditionZoneData.readIntoDict(str, colHSUClustersInZone+12, row, data, "BigPickupDistributionInZone")
    DatablockIO.nameInDict(DATABLOCK_BigPickupDistribution, data, "BigPickupDistributionInZone")
    data["TerminalPlacements"] = listdata.TerminalPlacements(zonestr)
    iExpeditionZoneData.readIntoDict(bool, colHSUClustersInZone+14, row, data, "ForbidTerminalsInZone")
    data["PowerGeneratorPlacements"] = listdata.PowerGeneratorPlacements(zonestr)
    data["DisinfectionStationPlacements"] = listdata.DisinfectionStationPlacements(zonestr)

    iExpeditionZoneData.readIntoDict(float, colHealthMulti, row, data, "HealthMulti")
    data["HealthPlacement"] = ZonePlacementWeights(iExpeditionZoneData, colHealthMulti+1, row, horizontal=True)
    iExpeditionZoneData.readIntoDict(float, colHealthMulti+4, row, data, "WeaponAmmoMulti")
    data["WeaponAmmoPlacement"] = ZonePlacementWeights(iExpeditionZoneData, colHealthMulti+5, row, horizontal=True)
    iExpeditionZoneData.readIntoDict(float, colHealthMulti+8, row, data, "ToolAmmoMulti")
    data["ToolAmmoPlacement"] = ZonePlacementWeights(iExpeditionZoneData, colHealthMulti+9, row, horizontal=True)
    iExpeditionZoneData.readIntoDict(float, colHealthMulti+12, row, data, "DisinfectionMulti")
    data["DisinfectionPlacement"] = ZonePlacementWeights(iExpeditionZoneData, colHealthMulti+13, row, horizontal=True)
    data["StaticSpawnDataContainers"] = listdata.StaticSpawnDataContainers(zonestr)

    return data

class ReactorWaveData:
    """
    a class that contains a dictionaries for the ReactorWaveData (since the sheet cannot contain 2d-3d data) \n
    access wave data by referencing ReactorWaveData.waves
    """

    def __init__(self, iWardenObjectiveReactorWaves:XlsxInterfacer.interface):
        """Generates stubs and the wave data to be returned from the getters"""
        startrow = 2

        startcolReactorWaves = XlsxInterfacer.ctn("B")
        startcolEnemyWaves = XlsxInterfacer.ctn("K")
        startcolEvents = XlsxInterfacer.ctn("Q")

        self.waves = []
        self.stubEnemyWaves = {}
        self.stubEvents = {}

        # EnemyWaves
        row = startrow
        while not(iWardenObjectiveReactorWaves.isEmpty(startcolEnemyWaves-1, row)):
            Snippet = {}
            waveNo = iWardenObjectiveReactorWaves.read(str, startcolEnemyWaves-1, row)
            iWardenObjectiveReactorWaves.readIntoDict(str, startcolEnemyWaves, row, Snippet, "WaveSettings")
            DatablockIO.nameInDict(DATABLOCK_SurvivalWaveSettings, Snippet, "WaveSettings")
            iWardenObjectiveReactorWaves.readIntoDict(str, startcolEnemyWaves+1, row, Snippet, "WavePopulation")
            DatablockIO.nameInDict(DATABLOCK_SurvivalWavePopulation, Snippet, "WavePopulation")
            iWardenObjectiveReactorWaves.readIntoDict(float, startcolEnemyWaves+2, row, Snippet, "SpawnTimeRel")
            iWardenObjectiveReactorWaves.readIntoDict(str, startcolEnemyWaves+3, row, Snippet, "SpawnType")
            EnumConverter.enumInDict(ENUMFILE_eReactorWaveSpawnType, Snippet, "SpawnType")
            EnsureKeyInDictArray(self.stubEnemyWaves, waveNo)
            self.stubEnemyWaves[waveNo].append(Snippet)
            row+= 1

        # Events
        row = startrow
        while not(iWardenObjectiveReactorWaves.isEmpty(startcolEvents-1, row)):
            Snippet = {}
            waveNo = iWardenObjectiveReactorWaves.read(str, startcolEvents-1, row)
            Snippet = WardenObjectiveEventData(iWardenObjectiveReactorWaves, startcolEvents, row, horizontal=True)
            EnsureKeyInDictArray(self.stubEvents, waveNo)
            self.stubEvents[waveNo].append(Snippet)
            row+= 1

        # ReactorWaves
        row = startrow
        while not(iWardenObjectiveReactorWaves.isEmpty(startcolReactorWaves, row)):
            wave = {}
            waveNo = iWardenObjectiveReactorWaves.read(str, startcolReactorWaves-1, row)
            iWardenObjectiveReactorWaves.readIntoDict(float, startcolReactorWaves, row, wave, "Warmup")
            iWardenObjectiveReactorWaves.readIntoDict(float, startcolReactorWaves+1, row, wave, "WarmupFail")
            iWardenObjectiveReactorWaves.readIntoDict(float, startcolReactorWaves+2, row, wave, "Wave")
            iWardenObjectiveReactorWaves.readIntoDict(float, startcolReactorWaves+3, row, wave, "Verify")
            iWardenObjectiveReactorWaves.readIntoDict(float, startcolReactorWaves+4, row, wave, "VerifyFail")
            iWardenObjectiveReactorWaves.readIntoDict(bool, startcolReactorWaves+5, row, wave, "VerifyInOtherZone")
            iWardenObjectiveReactorWaves.readIntoDict(str, startcolReactorWaves+6, row, wave, "ZoneForVerification")
            EnumConverter.enumInDict(ENUMFILE_eLocalZoneIndex, wave, "ZoneForVerification")
            wave["EnemyWaves"] = self.EnemyWaves(waveNo)
            wave["Events"] = self.Events(waveNo)
            self.waves.append(wave)
            row+= 1


    def EnemyWaves(self, identifier:str):
        """returns the EnemyWaves array for a specific zone"""
        try:return self.stubEnemyWaves[identifier]
        except KeyError:pass
        return []

    def Events(self, identifier:str):
        """returns the EnemyWaves array for a specific zone"""
        try:return self.stubEvents[identifier]
        except KeyError:pass
        return []

def WardenObjectiveBlock(iWardenObjective:XlsxInterfacer.interface, iWardenObjectiveReactorWaves:XlsxInterfacer.interface):
    """returns the Warden Objective"""
    # set up some checkpoints so if some of the data gets reformatted, not the entire function needs to be altered,
    # just the headings and contents of the section will need edited column values
    rowWavesOnElevatorLand = 22-1
    rowChainedPuzzleToActive = 70-1
    rowLightsOnFromBeginning = 84-1
    rowActivateHSU_ItemFromStart = 103-1

    data = {}

    data["Type"] = iWardenObjective.read(str, 1, 1)
    EnumConverter.enumInDict(ENUMFILE_eWardenObjectiveType, data, "Type")
    data["Header"] = iWardenObjective.read(XlsxInterfacer.blankable, 1, 3)
    data["MainObjective"] = iWardenObjective.read(XlsxInterfacer.blankable, 1, 4)
    data["FindLocationInfo"] = iWardenObjective.read(XlsxInterfacer.blankable, 1, 5)
    data["FindLocationInfoHelp"] = iWardenObjective.read(XlsxInterfacer.blankable, 1, 6)
    data["GoToZone"] = iWardenObjective.read(XlsxInterfacer.blankable, 1, 7)
    data["GoToZoneHelp"] = iWardenObjective.read(XlsxInterfacer.blankable, 1, 8)
    data["InZoneFindItem"] = iWardenObjective.read(XlsxInterfacer.blankable, 1, 9)
    data["InZoneFindItemHelp"] = iWardenObjective.read(XlsxInterfacer.blankable, 1, 10)
    data["SolveItem"] = iWardenObjective.read(XlsxInterfacer.blankable, 1, 11)
    data["SolveItemHelp"] = iWardenObjective.read(XlsxInterfacer.blankable, 1, 12)
    data["GoToWinCondition_Elevator"] = iWardenObjective.read(XlsxInterfacer.blankable, 1, 13)
    data["GoToWinConditionHelp_Elevator"] = iWardenObjective.read(XlsxInterfacer.blankable, 1, 14)
    data["GoToWinCondition_CustomGeo"] = iWardenObjective.read(XlsxInterfacer.blankable, 1, 15)
    data["GoToWinConditionHelp_CustomGeo"] = iWardenObjective.read(XlsxInterfacer.blankable, 1, 16)
    data["GoToWinCondition_ToMainLayer"] = iWardenObjective.read(XlsxInterfacer.blankable, 1, 17)
    data["GoToWinConditionHelp_ToMainLayer"] = iWardenObjective.read(XlsxInterfacer.blankable, 1, 18)
    iWardenObjective.readIntoDict(float, 1, 19, data, "ShowHelpDelay")

    data["WavesOnElevatorLand"] = GenericEnemyWaveDataList(iWardenObjective, 2, rowWavesOnElevatorLand+1, horizontal=True)
    iWardenObjective.readIntoDict(str, 1, rowWavesOnElevatorLand+6, data, "WaveOnElevatorWardenIntel")
    iWardenObjective.readIntoDict(str, 1, rowWavesOnElevatorLand+8, data, "FogTransitionDataOnElevatorLand")
    DatablockIO.nameInDict(DATABLOCK_FogSettings, data, "FogTransitionDataOnElevatorLand")
    iWardenObjective.readIntoDict(float, 1, rowWavesOnElevatorLand+9, data, "FogTransitionDurationOnElevatorLand")
    data["WavesOnActivate"] = GenericEnemyWaveDataList(iWardenObjective, 2, rowWavesOnElevatorLand+12, horizontal=True)
    iWardenObjective.readIntoDict(bool, 1, rowWavesOnElevatorLand+17, data, "StopAllWavesBeforeGotoWin")
    data["EventsOnActivate"] = []
    col,row = 2,rowWavesOnElevatorLand+20
    while not(iWardenObjective.isEmpty(col, row)):
        data["EventsOnActivate"].append(WardenObjectiveEventData(iWardenObjective, col, row, horizontal=False))
        col+= 1
    data["WavesOnGotoWin"] = GenericEnemyWaveDataList(iWardenObjective, 2, rowWavesOnElevatorLand+29, horizontal=True)
    iWardenObjective.readIntoDict(str, 1, rowWavesOnElevatorLand+34, data, "WaveOnGotoWinTrigger")
    EnumConverter.enumInDict(ENUMFILE_eRetrieveExitWaveTrigger, data, "WaveOnGotoWinTrigger")
    data["EventsOnGotoWin"] = []
    col,row = 2,rowWavesOnElevatorLand+37
    while not(iWardenObjective.isEmpty(col, row)):
        data["EventsOnGotoWin"].append(WardenObjectiveEventData(iWardenObjective, col, row, horizontal=False))
        col+= 1
    iWardenObjective.readIntoDict(str, 1, rowWavesOnElevatorLand+45, data, "FogTransitionDataOnGotoWin")
    DatablockIO.nameInDict(DATABLOCK_FogSettings, data, "FogTransitionDataOnGotoWin")
    iWardenObjective.readIntoDict(float, 1, rowWavesOnElevatorLand+46, data, "FogTransitionDurationOnGotoWin")

    iWardenObjective.readIntoDict(str, 1, rowChainedPuzzleToActive, data, "ChainedPuzzleToActive")
    DatablockIO.nameInDict(DATABLOCK_ChainedPuzzle, data, "ChainedPuzzleToActive")
    iWardenObjective.readIntoDict(str, 1, rowChainedPuzzleToActive+1, data, "ChainedPuzzleMidObjective")
    DatablockIO.nameInDict(DATABLOCK_ChainedPuzzle, data, "ChainedPuzzleMidObjective")
    iWardenObjective.readIntoDict(str, 1, rowChainedPuzzleToActive+2, data, "ChainedPuzzleAtExit")
    DatablockIO.nameInDict(DATABLOCK_ChainedPuzzle, data, "ChainedPuzzleAtExit")
    iWardenObjective.readIntoDict(float, 1, rowChainedPuzzleToActive+3, data, "ChainedPuzzleAtExitScanSpeedMultiplier")
    iWardenObjective.readIntoDict(int, 1, rowChainedPuzzleToActive+5, data, "Gather_RequiredCount")
    iWardenObjective.readIntoDict(str, 1, rowChainedPuzzleToActive+6, data, "Gather_ItemId")
    DatablockIO.nameInDict(DATABLOCK_Item, data, "Gather_ItemId")
    iWardenObjective.readIntoDict(int, 1, rowChainedPuzzleToActive+7, data, "Gather_SpawnCount")
    iWardenObjective.readIntoDict(int, 1, rowChainedPuzzleToActive+8, data, "Gather_MaxPerZone")
    data["Retrieve_Items"] = []
    col,row = 1,rowChainedPuzzleToActive+10
    while not(iWardenObjective.isEmpty(col, row)):
        data["Retrieve_Items"].append(DatablockIO.nameToId(DATABLOCK_Item, iWardenObjective.read(str, col, row)))
        col+= 1
    data["ReactorWaves"] = ReactorWaveData(iWardenObjectiveReactorWaves).waves

    iWardenObjective.readIntoDict(bool, 1, rowLightsOnFromBeginning, data, "LightsOnFromBeginning")
    iWardenObjective.readIntoDict(bool, 1, rowLightsOnFromBeginning+1, data, "LightsOnDuringIntro")
    iWardenObjective.readIntoDict(bool, 1, rowLightsOnFromBeginning+2, data, "LightsOnWhenStartupComplete")
    iWardenObjective.readIntoDict(str, 1, rowLightsOnFromBeginning+4, data, "SpecialTerminalCommand")
    iWardenObjective.readIntoDict(str, 1, rowLightsOnFromBeginning+5, data, "SpecialTerminalCommandDesc")
    data["PostCommandOutput"] = []
    col,row = 1,rowLightsOnFromBeginning+6
    while not(iWardenObjective.isEmpty(col, row)):
        data["PostCommandOutput"].append(iWardenObjective.read(str, col, row))
        col+= 1
    iWardenObjective.readIntoDict(int, 1, rowLightsOnFromBeginning+8, data, "PowerCellsToDistribute")
    iWardenObjective.readIntoDict(int, 1, rowLightsOnFromBeginning+10, data, "Uplink_NumberOfVerificationRounds")
    iWardenObjective.readIntoDict(int, 1, rowLightsOnFromBeginning+11, data, "Uplink_NumberOfTerminals")
    iWardenObjective.readIntoDict(int, 1, rowLightsOnFromBeginning+13, data, "CentralPowerGenClustser_NumberOfGenerators")
    iWardenObjective.readIntoDict(int, 1, rowLightsOnFromBeginning+14, data, "CentralPowerGenClustser_NumberOfPowerCells")
    data["CentralPowerGenClustser_FogDataSteps"] = []
    col,row = 1,rowLightsOnFromBeginning+16
    while not(iWardenObjective.isEmpty(col,row)):
        data["CentralPowerGenClustser_FogDataSteps"].append(GeneralFogDataStep(iWardenObjective, col, row, horizontal=False))
        col+= 1

    iWardenObjective.readIntoDict(str, 1, rowActivateHSU_ItemFromStart, data, "ActivateHSU_ItemFromStart")
    DatablockIO.nameInDict(DATABLOCK_Item, data, "ActivateHSU_ItemFromStart")
    iWardenObjective.readIntoDict(str, 1, rowActivateHSU_ItemFromStart+1, data, "ActivateHSU_ItemAfterActivation")
    DatablockIO.nameInDict(DATABLOCK_Item, data, "ActivateHSU_ItemAfterActivation")
    iWardenObjective.readIntoDict(bool, 1, rowActivateHSU_ItemFromStart+2, data, "ActivateHSU_StopEnemyWavesOnActivation")
    iWardenObjective.readIntoDict(bool, 1, rowActivateHSU_ItemFromStart+3, data, "ActivateHSU_ObjectiveCompleteAfterInsertion")
    iWardenObjective.readIntoDict(bool, 1, rowActivateHSU_ItemFromStart+4, data, "ActivateHSU_RequireItemAfterActivationInExitScan")
    data["ActivateHSU_Events"] = []
    col,row = 2,rowActivateHSU_ItemFromStart+7
    while not(iWardenObjective.isEmpty(col,row)):
        data["ActivateHSU_Events"].append(WardenObjectiveEventData(iWardenObjective, col, row, horizontal=False))
        col+= 1
    
    # Set default values
    data["name"] = "DPK Utility Objective"
    data["internalEnabled"] = False
    data["persistentID"] = 0
    # Attempt to fill default values with those from the table
    iWardenObjective.readIntoDict(str,1, rowActivateHSU_ItemFromStart+15, data, "name")
    iWardenObjective.readIntoDict(bool,1, rowActivateHSU_ItemFromStart+16, data, "internalEnabled")
    iWardenObjective.readIntoDict(int,1, rowActivateHSU_ItemFromStart+17, data, "persistentID")
    return data

def finalizeData(dictExpeditionInTier:dict, arrdictLevelLayoutBlock:typing.List[dict], arrdictWardenObjectiveBlock:typing.List[dict]):
    """
    finalizeData takes the ExpeditionInTier, LevelLayoutBlocks, and WardenObjectiveBlocks \n
    for the LevelLayoutBlocks, it fills in the default block metadata \n
    for the WardenObjectiveBlocks, it overrides the block metadata
    (Both Arrays should be of length 3 and contain a dictionary or None)
    """
    utilityname = dictExpeditionInTier["Descriptive"]["PublicName"] + " - " + dictExpeditionInTier["Descriptive"]["Prefix"]
    levelEnabled = dictExpeditionInTier["Enabled"]

    if arrdictLevelLayoutBlock[0] != None:
        arrdictLevelLayoutBlock[0]["name"] = utilityname + "L1 Layout"
        arrdictLevelLayoutBlock[0]["internalEnabled"] = levelEnabled
        arrdictLevelLayoutBlock[0]["persistentID"] = dictExpeditionInTier["LevelLayoutData"]

    if arrdictLevelLayoutBlock[1] != None:
        arrdictLevelLayoutBlock[1]["name"] = utilityname + "L2 Layout"
        arrdictLevelLayoutBlock[1]["internalEnabled"] = levelEnabled and dictExpeditionInTier["SecondaryLayerEnabled"]
        arrdictLevelLayoutBlock[1]["persistentID"] = dictExpeditionInTier["SecondaryLayout"]

    if arrdictLevelLayoutBlock[2] != None:
        arrdictLevelLayoutBlock[2]["name"] = utilityname + "L3 Layout"
        arrdictLevelLayoutBlock[2]["internalEnabled"] = levelEnabled and dictExpeditionInTier["ThirdLayerEnabled"]
        arrdictLevelLayoutBlock[2]["persistentID"] = dictExpeditionInTier["ThirdLayout"]

    if arrdictWardenObjectiveBlock[0] != None:
        arrdictWardenObjectiveBlock[0]["name"] = utilityname + "L1 Objective"
        arrdictWardenObjectiveBlock[0]["internalEnabled"] = levelEnabled
        arrdictWardenObjectiveBlock[0]["persistentID"] = dictExpeditionInTier["MainLayerData"]["ObjectiveData"]["DataBlockId"]

    if arrdictWardenObjectiveBlock[1] != None:
        arrdictWardenObjectiveBlock[1]["name"] = utilityname + "L2 Objective"
        arrdictWardenObjectiveBlock[1]["internalEnabled"] = levelEnabled and dictExpeditionInTier["SecondaryLayerEnabled"]
        arrdictWardenObjectiveBlock[1]["persistentID"] = dictExpeditionInTier["SecondaryLayerData"]["ObjectiveData"]["DataBlockId"]

    if arrdictWardenObjectiveBlock[2] != None:
        arrdictWardenObjectiveBlock[2]["name"] = utilityname + "L3 Objective"
        arrdictWardenObjectiveBlock[2]["internalEnabled"] = levelEnabled and dictExpeditionInTier["ThirdLayerEnabled"]
        arrdictWardenObjectiveBlock[2]["persistentID"] = dictExpeditionInTier["ThirdLayerData"]["ObjectiveData"]["DataBlockId"]

def UtilityJob(LevelXlsxFile:io.BytesIO, RundownBlock:DatablockIO.datablock, LevelLayoutDataBlock:DatablockIO.datablock, WardenObjectiveDataBlock:DatablockIO.datablock, tier:typing.Union[int,str], index:int, logger:logging.Logger=None):
    """
    Have the utility start a job
    This will take an xlsx file as input (use open(file, 'rb'))
    In addition it will take the Rundown Block, level tier (0-4), and index of the level in the tier
    It will automatically insert the items 
    """
    if logger == None:
        logger = logging.getLevelName("none")
        logger.addHandler(logging.NullHandler())
        logger.propagate = False

    logger.info("Starting level utilty job: \""+LevelXlsxFile.name+"\"")

    if (isinstance(tier, int)):
        try:
            tierName = ["TierA","TierB","TierC","TierD","TierE"][tier]
        except IndexError:
            raise IndexError("Invalid level tier: "+tier)
    else:
        if tier in ["TierA","TierB","TierC","TierD","TierE"]:
            tierName = tier
        else:
            raise Exception("Invalid level tier: "+tier)

    # get all interfaces
    iKey = XlsxInterfacer.interface(pandas.read_excel(LevelXlsxFile, "Key", header=None))
    if iKey.read(str, 0, 5).split(".")[0:2] != Version.split(".")[0:2]:
        raise Exception("Version mismatch between utility and sheet, incompatible.")

    # Load all sheets (and allow missing Zone and Objective data)
    iExpeditionInTier = XlsxInterfacer.interface(pandas.read_excel(LevelXlsxFile, "ExpeditionInTier", header=None))

    try:
        iL1ExpeditionZoneData = XlsxInterfacer.interface(pandas.read_excel(LevelXlsxFile, "L1 ExpeditionZoneData", header=None))
        iL1ExpeditionZoneDataLists = XlsxInterfacer.interface(pandas.read_excel(LevelXlsxFile, "L1 ExpeditionZoneData Lists", header=None))
        logger.debug("Found L1 ExpeditionZoneData")
    except (xlrd.biffh.XLRDError, ValueError):
        logger.debug("No L1 ExpeditionZoneData")
    try:
        iL2ExpeditionZoneData = XlsxInterfacer.interface(pandas.read_excel(LevelXlsxFile, "L2 ExpeditionZoneData", header=None))
        iL2ExpeditionZoneDataLists = XlsxInterfacer.interface(pandas.read_excel(LevelXlsxFile, "L2 ExpeditionZoneData Lists", header=None))
        logger.debug("Found L2 ExpeditionZoneData")
    except (xlrd.biffh.XLRDError, ValueError):
        logger.debug("No L2 ExpeditionZoneData")
    try:
        iL3ExpeditionZoneData = XlsxInterfacer.interface(pandas.read_excel(LevelXlsxFile, "L3 ExpeditionZoneData", header=None))
        iL3ExpeditionZoneDataLists = XlsxInterfacer.interface(pandas.read_excel(LevelXlsxFile, "L3 ExpeditionZoneData Lists", header=None))
        logger.debug("Found L3 ExpeditionZoneData")
    except (xlrd.biffh.XLRDError, ValueError):
        logger.debug("No L3 ExpeditionZoneData")

    try:
        iL1WardenObjective = XlsxInterfacer.interface(pandas.read_excel(LevelXlsxFile, "L1 WardenObjective", header=None))
        iL1WardenObjectiveReactorWaves = XlsxInterfacer.interface(pandas.read_excel(LevelXlsxFile, "L1 WardenObjective ReactorWaves", header=None))
        logger.debug("Found L1 WardenObjective")
    except (xlrd.biffh.XLRDError, ValueError):
        logger.debug("No L1 WardenObjective")
    try:
        iL2WardenObjective = XlsxInterfacer.interface(pandas.read_excel(LevelXlsxFile, "L2 WardenObjective", header=None))
        iL2WardenObjectiveReactorWaves = XlsxInterfacer.interface(pandas.read_excel(LevelXlsxFile, "L2 WardenObjective ReactorWaves", header=None))
        logger.debug("Found L2 WardenObjective")
    except (xlrd.biffh.XLRDError, ValueError):
        logger.debug("No L2 WardenObjective")
    try:
        iL3WardenObjective = XlsxInterfacer.interface(pandas.read_excel(LevelXlsxFile, "L3 WardenObjective", header=None))
        iL3WardenObjectiveReactorWaves = XlsxInterfacer.interface(pandas.read_excel(LevelXlsxFile, "L3 WardenObjective ReactorWaves", header=None))
        logger.debug("Found L3 WardenObjective")
    except (xlrd.biffh.XLRDError, ValueError):
        logger.debug("No L3 WardenObjective")

    # Convert sheets into dictionaries
    try:dictExpeditionInTier = ExpeditionInTier(iExpeditionInTier)
    except Exception as e:raise Exception("Problem creating ExpeditionInTier: "+str(e))

    arrdictLevelLayoutBlock = [None,None,None]
    try:arrdictLevelLayoutBlock[0] = LevelLayoutBlock(iL1ExpeditionZoneData, iL1ExpeditionZoneDataLists)
    except NameError:pass
    except Exception as e:raise Exception("Problem reading L1 LevelLayout: "+str(e))
    try:arrdictLevelLayoutBlock[1] = LevelLayoutBlock(iL2ExpeditionZoneData, iL2ExpeditionZoneDataLists)
    except NameError:pass
    except Exception as e:
        logger.error("Problem reading L2 LevelLayout (skipping layout): "+str(e))
    try:arrdictLevelLayoutBlock[2] = LevelLayoutBlock(iL3ExpeditionZoneData, iL3ExpeditionZoneDataLists)
    except NameError:pass
    except Exception as e:
        logger.error("Problem reading L3 LevelLayout (skipping layout): "+str(e))

    arrdictWardenObjectiveBlock = [None,None,None]
    try:arrdictWardenObjectiveBlock[0] = WardenObjectiveBlock(iL1WardenObjective, iL1WardenObjectiveReactorWaves)
    except NameError:pass
    except Exception as e:raise Exception("Problem reading L1 WardenOjbective: "+str(e))
    try:arrdictWardenObjectiveBlock[1] = WardenObjectiveBlock(iL2WardenObjective, iL2WardenObjectiveReactorWaves)
    except NameError:pass
    except Exception as e:
        logger.error("Problem reading L2 WardenOjbective (skipping objective): "+str(e))
    try:arrdictWardenObjectiveBlock[2] = WardenObjectiveBlock(iL3WardenObjective, iL3WardenObjectiveReactorWaves)
    except NameError:pass
    except Exception as e:
        logger.error("Problem reading L3 WardenOjbective (skipping objective): "+str(e))

    # copy descriptive from ExpeditionInTier into LevelLayout and WardenObjective blocks
    finalizeData(dictExpeditionInTier, arrdictLevelLayoutBlock, arrdictWardenObjectiveBlock)

    logger.debug("Block data finalized")

    # Add ExpeditionInTier to RundownBlock
    try:
        # if a level exists at the specified index, overwrite it on the rundown
        RundownBlock[tierName][index] = dictExpeditionInTier
    except IndexError:
        RundownBlock[tierName].append(dictExpeditionInTier)

    # Add/Edit the LevelLayout and WardenObjective Blocks
    if arrdictLevelLayoutBlock[0] != None:
        LevelLayoutDataBlock.writeblock(arrdictLevelLayoutBlock[0])
    if arrdictLevelLayoutBlock[1] != None:
        LevelLayoutDataBlock.writeblock(arrdictLevelLayoutBlock[1])
    if arrdictLevelLayoutBlock[2] != None:
        LevelLayoutDataBlock.writeblock(arrdictLevelLayoutBlock[2])
    if arrdictWardenObjectiveBlock[0] != None:
        WardenObjectiveDataBlock.writeblock(arrdictWardenObjectiveBlock[0])
    if arrdictWardenObjectiveBlock[1] != None:
        WardenObjectiveDataBlock.writeblock(arrdictWardenObjectiveBlock[1])
    if arrdictWardenObjectiveBlock[2] != None:
        WardenObjectiveDataBlock.writeblock(arrdictWardenObjectiveBlock[2])

    logger.debug("Blocks written to dictionaries")

    logger.info("Finished level utilty job: \""+LevelXlsxFile.name+"\"")
    return True

def main():
    parser = argparse.ArgumentParser(
        prog="DPK LevelUtility",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        description=textwrap.dedent("""\
        This is a tool created by DPK.
        This tool can convert xlsx to a bunch of GTFO Datablock pieces.
        """)
    )

    parser.add_argument('path', type=str, nargs='*', help='Specific xlsx file(s) to add to datablocks.')
    parser.add_argument('-n', "--noinput", action='store_true', help='[N]o inputs (which could be annoying in CLI and scripts)')
    parser.add_argument('-v', "--verbosity", type=str.upper, help='Changes console [v]erbosity', choices=['DEBUG','INFO','WARNING','ERROR','CRITICAL'], default='INFO')

    # allow the arguments to be used anywhere needed
    global args
    args = parser.parse_args()

    # create a logs folder if it does not exist already
    logdir = os.path.join(os.path.dirname(__file__),"./logs/")
    if not os.path.exists(logdir):os.makedirs(logdir)

    logformatter = logging.Formatter(fmt="%(asctime)s\t: %(name)s\t: %(levelname)s\t: %(message)s")
    logformatter.converter = time.gmtime
    consoleformatter = logging.Formatter(fmt="%(levelname)s : %(message)s")
    logformatter.converter = time.gmtime

    logger = logging.getLogger("LevelUtilty")
    logger.setLevel(logging.DEBUG)

    logfilehandler = logging.FileHandler(os.path.join(logdir,time.strftime("%Y.%m.%d.%H.%M.%S",time.gmtime())+".LevelUtility.log"))
    logfilehandler.setFormatter(logformatter)
    logger.addHandler(logfilehandler)

    consoleloghandler = logging.StreamHandler()
    consoleloghandler.setLevel(getattr(logging, args.verbosity))
    consoleloghandler.setFormatter(consoleformatter)
    logger.addHandler(consoleloghandler)

    joblogger = logger.getChild("job")

    # Wait for hit return to continue
    def waitUser():
        input("HIT ENTER TO CONTINUE. ") # waiting on the user won't be written to the log
        return

    logger.debug("Running DPK's LevelUtilty with the given arguments:\n\t"+str(args))

    paths = args.path

    anythingDone = False

    pathsDefault = False # True for when default paths are being used
    if paths==[]:
        pathsDefault = True
        paths = defaultpaths
        logger.warning("No files given, using default paths.")

    # Open Datablocks to be manipulated
    try:
        RundownDataBlock = DatablockIO.datablock(open(os.path.join(blockpath,ConfigManager.config["Project"]["blockprefix"]+"RundownDataBlock"+ConfigManager.config["Project"]["blocksuffix"]), 'r+', encoding="utf8"))
        LevelLayoutDataBlock = DatablockIO.datablock(open(os.path.join(blockpath,ConfigManager.config["Project"]["blockprefix"]+"LevelLayoutDataBlock"+ConfigManager.config["Project"]["blocksuffix"]), 'r+', encoding="utf8"))
        WardenObjectiveDataBlock = DatablockIO.datablock(open(os.path.join(blockpath,ConfigManager.config["Project"]["blockprefix"]+"WardenObjectiveDataBlock"+ConfigManager.config["Project"]["blocksuffix"]), 'r+', encoding="utf8"))
    except FileNotFoundError as e:
        if not args.noinput:
            print("Missing a DataBlock: " + str(e))
            input()
        raise FileNotFoundError("Missing a DataBlock: " + str(e))

    for path in paths:
        logger.info("Working on: \""+path+"\"")
        try:
            fxlsx = open(path, 'rb')
        except FileNotFoundError:
            if (pathsDefault): logger.debug("No default file, skipping: \""+path+"\"")
            else: logger.error("Path does not have a file: \""+path+"\"")
            continue
        try:
            iMeta = XlsxInterfacer.interface(pandas.read_excel(fxlsx, "Meta", header=None))
        except xlrd.biffh.XLRDError:
            logger.error("Missing meta sheet for level: \""+path+"\"")
            continue
        try:
            rundownID = iMeta.read(int, 0, 2)
            tierName = iMeta.read(str, 1, 2)
            expeditionIndex = iMeta.read(int, 2, 2)
        except XlsxInterfacer.EmptyCell:
            logger.error("Missing data on meta sheet: \""+path+"\"")
            continue

        try:
            RundownBlock = RundownDataBlock.data["Blocks"][RundownDataBlock.find(rundownID)]
        except TypeError as e:
            logger.error("Current blocks lack rundown with id "+str(rundownID)+": \""+path+"\"\n\t"+str(e))
            continue

        try:
            UtilityJob(fxlsx, RundownBlock, LevelLayoutDataBlock, WardenObjectiveDataBlock, tierName, expeditionIndex, logger=joblogger) # Utilty job should stay silenced because it is currently unable to write to the log file
        except Exception as e:
            # This if condition is to not write this twice in the debug log when something goes wrong
            logger.error("Something went wrong reading the sheet: \""+path+"\"\n\t"+str(e))
            continue
        logger.info("Finished with: \""+path+"\"")
        anythingDone = True


    # Save manipulated datablocks
    if anythingDone:
        logger.info("Writing blocks...")
        RundownDataBlock.writedatablock()
        LevelLayoutDataBlock.writedatablock()
        WardenObjectiveDataBlock.writedatablock()
        logger.info("Blocks written.")
    else:
        logger.info("Nothing to write.")

    # handle end of program commands
    if (not anythingDone):
        logger.warning("Nothing happened... are you sure you didn't do anything wrong?")
    logger.info("Done.")
    if not(args.noinput):waitUser()

if __name__ == "__main__":
    main()
