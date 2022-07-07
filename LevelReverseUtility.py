"""
This is a tool created by DPK
"""

import argparse
import io
import logging
import os
import re
import shutil
import textwrap
import time
import typing

import numpy
import openpyxl
import pandas
import xlrd

import ConfigManager
import DatablockIO
import EnumConverter
import XlsxInterfacer

# argparse: used to get arguments in CLI (to decide which files to turn into levels encoding/decoding and which file)
# io:       used to read from and write to files
# logging:  used to create log files for the tool
# os:       used to create log directory
# re:       used to sanitize text of newlines and Windows reserved characters
# shutil:   used to copy the template
# textwrap: used to help format argparse description
# time:     used to give the log file's name a timestamp and set time to gmttime
# typing:   used to annotate function parameters
# numpy:    used to manipulate the inconsistant numpy data read by pandas
# openpyxl: used to copy templated sheets
# pandas:   used to read the excel data
# xlrd:     used to catch and throw excel errors when initially reading the sheets

# a regex to capture the newlines the devs put into the json
devnewlnregex = "\r?\n" # TODO figure out if this regex hurts the integrity of audio logs' text
sheetlf = "\\\\n"
sheetcrlf = "\\\\r\\\\n"
# a regex to capture the tabs the debs put into the json
devtabregex = "\t"
sheettb = "\\\\t"
# a regex to capture the Windows Reserved characters; see https://docs.microsoft.com/en-us/windows/win32/fileio/naming-a-file#naming-conventions
winreserveregex = "[<>:\"/\\\\|?*]"

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
# path to template file
templatepath = os.path.join(os.path.dirname(__file__), ConfigManager.config["LevelReverseUtility"]["templatepath"])
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

def ZonePlacementData(interface:XlsxInterfacer.interface, data:dict, col:int, row:int, horizontal=True):
    """
    add ZonePlacementData to the specified interface \n
    col and row define the upper left value (not header) \n
    horizontal is true if the values are in the same row
    """
    writeEnumFromDict(ENUMFILE_eLocalZoneIndex, interface, col, row, data, "LocalIndex")
    try:ZonePlacementWeights(interface, data["Weights"], col+2*horizontal, row+2*(not horizontal), horizontal)
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
    try:ZonePlacementWeights(interface, data["PlacementWeights"], col, row, horizontal)
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
            ZonePlacementWeights(interface, placement["Weights"], col+3*(not horizontal), row+3*horizontal, horizontal=(not horizontal))
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
    writeEnumFromDict(ENUMFILE_LG_LayerType, interface, col+2*horizontal, row+2*(not horizontal), data, "Layer")
    writeEnumFromDict(ENUMFILE_eLocalZoneIndex, interface, col+3*horizontal, row+3*(not horizontal), data, "LocalIndex")
    interface.writeFromDict(col+4*horizontal, row+4*(not horizontal), data, "Delay")
    interface.writeFromDict(col+5*horizontal, row+5*(not horizontal), data, "WardenIntel")
    interface.writeFromDict(col+6*horizontal, row+6*(not horizontal), data, "SoundID")
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
    try:
        itercol,iterrow = col+5, row
        for zone in data["ZonesWithBulkheadEntrance"]:
            # NOTE textmode may need a toggle in this file for whether the json should have text enums
            interface.write(EnumConverter.enumToIndex(ENUMFILE_eLocalZoneIndex, zone, textmode=True), itercol, iterrow)
            itercol+= 1
    except KeyError:pass
    try:
        itercol,iterrow = col+5, row+1
        for placement in data["BulkheadDoorControllerPlacements"]:
            BulkheadDoorPlacementData(interface, placement, itercol, iterrow, horizontal=False)
            itercol+= 1
    except KeyError:pass
    ZonePlacementWeightsList(interface, data["BulkheadKeyPlacements"], col+5, row+7, horizontal=True)
    try:
        interface.writeFromDict(col+5, row+13, data["ObjectiveData"], "DataBlockId")
        writeEnumFromDict(ENUMFILE_eWardenObjectiveWinCondition, interface, col+5, row+14, data["ObjectiveData"], "WinCondition")
        ZonePlacementWeightsList(interface, data["ObjectiveData"]["ZonePlacementDatas"], col+5, row+15, horizontal=True)
    except KeyError:pass
    try:
        interface.writeFromDict(col+5, row+22, data["ArtifactData"], "ArtifactAmountMulti")
        writePublicNameFromDict(DATABLOCK_ArtifactDistributionDataBlock, interface, col+5, row+23, data["ArtifactData"], "ArtifactLayerDistributionDataID")
        itercol,iterrow = col+5, row+24
        for distribution in data["ArtifactData"]["ArtifactZoneDistributions"]:
            ArtifactZoneDistribution(interface, distribution, itercol, iterrow, horizontal=False)
            itercol+= 1
    except KeyError:pass

def frameExpeditionInTier(iExpeditionInTier:XlsxInterfacer.interface, ExpeditionInTierData:dict):
    """
    edit the iExpeditionInTier pandas dataFrame
    """
    iExpeditionInTier.writeFromDict(0, 2, ExpeditionInTierData, "Enabled")
    writeEnumFromDict(ENUMFILE_eExpeditionAccessibility, iExpeditionInTier, 6, 2, ExpeditionInTierData, "Accessibility")

    try:
        iExpeditionInTier.writeFromDict(0, 10, ExpeditionInTierData["CustomProgressionLock"], "MainSectors")
        iExpeditionInTier.writeFromDict(1, 10, ExpeditionInTierData["CustomProgressionLock"], "SecondarySectors")
        iExpeditionInTier.writeFromDict(2, 10, ExpeditionInTierData["CustomProgressionLock"], "ThirdSectors")
        iExpeditionInTier.writeFromDict(3, 10, ExpeditionInTierData["CustomProgressionLock"], "AllClearedSectors")
    except KeyError:pass

    try:
        iExpeditionInTier.writeFromDict(14, 10, ExpeditionInTierData["Descriptive"], "Prefix")
        iExpeditionInTier.writeFromDict(14, 11, ExpeditionInTierData["Descriptive"], "PublicName")
        iExpeditionInTier.writeFromDict(14, 12, ExpeditionInTierData["Descriptive"], "IsExtraExpedition")
        iExpeditionInTier.writeFromDict(14, 17, ExpeditionInTierData["Descriptive"], "ExpeditionDepth")
        iExpeditionInTier.writeFromDict(14, 18, ExpeditionInTierData["Descriptive"], "EstimatedDuration")
        # TODO regex replace new lines to be "\n"
        try:
            desc = re.sub(devnewlnregex, sheetcrlf, str(ExpeditionInTierData["Descriptive"]["ExpeditionDescription"])) # TODO fix this bodge for localization
            iExpeditionInTier.write(desc, 14, 19)
        except KeyError:pass
        try:
            desc = re.sub(devnewlnregex, sheetcrlf, str(ExpeditionInTierData["Descriptive"]["RoleplayedWardenIntel"])) # TODO fix this bodge for localization
            iExpeditionInTier.write(desc, 14, 20)
        except KeyError:pass
        try:
            desc = re.sub(devnewlnregex, sheetlf, str(ExpeditionInTierData["Descriptive"]["DevInfo"])) # TODO fix this bodge for localization
            iExpeditionInTier.write(desc, 14, 21)
        except KeyError:pass
    except KeyError:pass

    try:
        iExpeditionInTier.writeFromDict(5, 10, ExpeditionInTierData["Seeds"], "BuildSeed")
        iExpeditionInTier.writeFromDict(6, 10, ExpeditionInTierData["Seeds"], "FunctionMarkerOffset")
        iExpeditionInTier.writeFromDict(7, 10, ExpeditionInTierData["Seeds"], "StandardMarkerOffset")
        iExpeditionInTier.writeFromDict(8, 10, ExpeditionInTierData["Seeds"], "LightJobSeedOffset")
    except KeyError:pass

    try:
        writePublicNameFromDict(DATABLOCK_ComplexResourceSet,       iExpeditionInTier, 0, 14, ExpeditionInTierData["Expedition"], "ComplexResourceData")
        writePublicNameFromDict(DATABLOCK_LightSettings,            iExpeditionInTier, 2, 14, ExpeditionInTierData["Expedition"], "LightSettings")
        writePublicNameFromDict(DATABLOCK_FogSettings,              iExpeditionInTier, 3, 14, ExpeditionInTierData["Expedition"], "FogSettings")
        writePublicNameFromDict(DATABLOCK_EnemyPopulation,          iExpeditionInTier, 4, 14, ExpeditionInTierData["Expedition"], "EnemyPopulation")
        writePublicNameFromDict(DATABLOCK_ExpeditionBalance,        iExpeditionInTier, 5, 14, ExpeditionInTierData["Expedition"], "ExpeditionBalance")
        writePublicNameFromDict(DATABLOCK_SurvivalWaveSettings,     iExpeditionInTier, 6, 14, ExpeditionInTierData["Expedition"], "ScoutWaveSettings")
        writePublicNameFromDict(DATABLOCK_SurvivalWavePopulation,   iExpeditionInTier, 7, 14, ExpeditionInTierData["Expedition"], "ScoutWavePopulation")
    except KeyError:pass

    iExpeditionInTier.writeFromDict(0, 21, ExpeditionInTierData, "LevelLayoutData")
    try:
        LayerData(iExpeditionInTier, ExpeditionInTierData["MainLayerData"], 0, 36)
    except KeyError:pass

    iExpeditionInTier.writeFromDict(2, 21, ExpeditionInTierData, "SecondaryLayerEnabled")
    iExpeditionInTier.writeFromDict(3, 21, ExpeditionInTierData, "SecondaryLayout")
    try:
        writeEnumFromDict(ENUMFILE_LG_LayerType, iExpeditionInTier, 2, 25, ExpeditionInTierData["BuildSecondaryFrom"], "LayerType")
        iExpeditionInTier.writeFromDict(3, 25, ExpeditionInTierData["BuildSecondaryFrom"], "Zone")
    except KeyError:pass
    try:
        LayerData(iExpeditionInTier, ExpeditionInTierData["SecondaryLayerData"], 0, 66)
    except KeyError:pass

    iExpeditionInTier.writeFromDict(5, 21, ExpeditionInTierData, "ThirdLayerEnabled")
    iExpeditionInTier.writeFromDict(6, 21, ExpeditionInTierData, "ThirdLayout")
    try:
        writeEnumFromDict(ENUMFILE_LG_LayerType, iExpeditionInTier, 5, 25, ExpeditionInTierData["BuildThirdFrom"], "LayerType")
        iExpeditionInTier.writeFromDict(6, 25, ExpeditionInTierData["BuildThirdFrom"], "Zone")
    except KeyError:pass
    try:
        LayerData(iExpeditionInTier, ExpeditionInTierData["ThirdLayerData"], 0, 95)
    except KeyError:pass

    try:
        iExpeditionInTier.writeFromDict(2, 29, ExpeditionInTierData["SpecialOverrideData"], "WeakResourceContainerWithPackChanceForLocked")
    except KeyError:pass

class ExpeditionZoneDataLists:
    """a class that decreases the dimentions of the dictionaries in ExpeditionZoneData (since the sheet cannot contain 2d-3d data)"""

    def __init__(self, LevelLayout:dict):
        """Generates numerous stubs to be iterated through and written to the interface"""

        self.stubEventsOnEnter = []
        self.stubProgressionPuzzleToEnter = []
        self.stubEnemySpawningInZone = []
        self.stubEnemyRespawnExcludeList = []
        self.stubTerminalPlacements = []
        self.stubLocalLogFiles = []
        groupedLocalLogFiles = []
        self.stubPowerGeneratorPlacements = []
        self.stubDisinfectionStationPlacements = []
        self.stubStaticSpawnDataContainers = []

        loggroupstart = 'A'

        for ZoneData in LevelLayout["Zones"]:
            zone = EnumConverter.indexToEnum(ENUMFILE_eLocalZoneIndex, ZoneData["LocalIndex"], False)

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

            # Unlike the other lists, the terminal placement is several dimentions deep and must be handled piece by piece to find unique Log files
            try:
                # TerminalPlacements
                for placement in ZoneData["TerminalPlacements"]:

                    try: # if there are log files...
                        logfiles = placement["LocalLogFiles"] # will jump to the except if no logs exist

                        if logfiles == []:
                            # if there are no logs, keep the group blank
                            placement["LocalLogFiles"] = ""
                            raise KeyError # jump to the no logs exist

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
        """Iterates through the numerous stubs and writes them to the specified interface"""
        startrow = 2

        startcolEventsOnEnter =                     XlsxInterfacer.ctn("C")
        startcolEventsOnPortalWarp =                XlsxInterfacer.ctn("M")
        startcolEventsOnApproachDoor =              XlsxInterfacer.ctn("AX")
        startcolEventsOnUnlockDoor =                XlsxInterfacer.ctn("CJ")
        startcolEventsOnOpenDoor =                  XlsxInterfacer.ctn("EJ")
        startcolEventsOnDoorScanStart =             XlsxInterfacer.ctn("FV")
        startcolEventsOnDoorScanDone =              XlsxInterfacer.ctn("HH")
        startcolEventsOnBossDeath =                 XlsxInterfacer.ctn("IT")
        startcolEventsOnTrigger =                   XlsxInterfacer.ctn("KF")
        startcolProgressionPuzzleToEnter =          XlsxInterfacer.ctn("LS")
        startcolEventsOnTerminalDeactivateAlarm =   XlsxInterfacer.ctn("MA")
        startcolWorldEventChainedPuzzleDatas =      XlsxInterfacer.ctn("NM")
        startcolEnemySpawningInZone =               XlsxInterfacer.ctn("PD")
        startcolEnemyRespawnExcludeList =           XlsxInterfacer.ctn("PK")
        startcolSpecificPickupSpawningDatas =       XlsxInterfacer.ctn("PN")
        startcolLocalLogFiles =                     XlsxInterfacer.ctn("QO")
        startcolUniqueCommands =                    XlsxInterfacer.ctn("QW")
        startcolCommandEvents =                     XlsxInterfacer.ctn("RD")
        startcolTerminalZoneSelectionDatas =        XlsxInterfacer.ctn("RI")
        startcolTerminalPlacements =                XlsxInterfacer.ctn("PR")
        startcolPowerGeneratorPlacements =          XlsxInterfacer.ctn("TD")
        startcolDisinfectionStationPlacements =     XlsxInterfacer.ctn("TK")
        startcolStaticSpawnDataContainers =         XlsxInterfacer.ctn("TR")

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

            try:
                writeEnumFromDict(ENUMFILE_TERM_State, iExpeditionZoneDataLists, startcolTerminalPlacements+7, row, Snippet["StartingStateData"], "StartingState")

                iExpeditionZoneDataLists.writeFromDict(startcolTerminalPlacements+11, row, Snippet["StartingStateData"], "AudioEventEnter")
                iExpeditionZoneDataLists.writeFromDict(startcolTerminalPlacements+12, row, Snippet["StartingStateData"], "AudioEventExit")
                # TODO convert sound placeholders
            except KeyError:pass

            row+= 1

        row = startrow
        # LocalLogFiles
        for Snippet in self.stubLocalLogFiles:
            iExpeditionZoneDataLists.writeFromDict(startcolLocalLogFiles, row, Snippet, "Log Group")

            iExpeditionZoneDataLists.writeFromDict(startcolLocalLogFiles+1, row, Snippet, "FileName")
            try:
                content = re.sub(devnewlnregex, sheetcrlf, str(Snippet["FileContent"])) # TODO fix this bodge for localization
                content = re.sub(devtabregex, sheettb, content)
                iExpeditionZoneDataLists.write(content, startcolLocalLogFiles+2, row)
            except KeyError:pass
            iExpeditionZoneDataLists.writeFromDict(startcolLocalLogFiles+4, row, Snippet, "AttachedAudioFile")
            iExpeditionZoneDataLists.writeFromDict(startcolLocalLogFiles+5, row, Snippet, "AttachedAudioByteSize")
            iExpeditionZoneDataLists.writeFromDict(startcolLocalLogFiles+6, row, Snippet, "PlayerDialogToTriggerAfterAudio")
            # TODO convert sound placeholders

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
    colPuzzleType = XlsxInterfacer.ctn("AF")
    colHSUClustersInZone = XlsxInterfacer.ctn("BH")
    colHealthMulti = XlsxInterfacer.ctn("BZ")

    writeEnumFromDict(ENUMFILE_eLocalZoneIndex, iExpeditionZoneData, 0, row, ExpeditionZoneData, "LocalIndex")
    iExpeditionZoneData.writeFromDict(5, row, ExpeditionZoneData, "SubSeed")
    iExpeditionZoneData.writeFromDict(8, row, ExpeditionZoneData, "BulkheadDCScanSeed")
    writeEnumFromDict(ENUMFILE_SubComplex, iExpeditionZoneData, 9, row, ExpeditionZoneData, "SubComplex")
    iExpeditionZoneData.writeFromDict(10, row, ExpeditionZoneData, "CustomGeomorph")
    try:
        iExpeditionZoneData.writeFromDict(12, row, ExpeditionZoneData["CoverageMinMax"], "x")
        iExpeditionZoneData.writeFromDict(13, row, ExpeditionZoneData["CoverageMinMax"], "y")
    except KeyError:pass
    writeEnumFromDict(ENUMFILE_eLocalZoneIndex, iExpeditionZoneData, 14, row, ExpeditionZoneData, "BuildFromLocalIndex")
    writeEnumFromDict(ENUMFILE_eZoneBuildFromType, iExpeditionZoneData, 15, row, ExpeditionZoneData, "StartPosition")
    iExpeditionZoneData.writeFromDict(16, row, ExpeditionZoneData, "StartPosition_IndexWeight")
    writeEnumFromDict(ENUMFILE_eZoneBuildFromExpansionType, iExpeditionZoneData, 17, row, ExpeditionZoneData, "StartExpansion")
    writeEnumFromDict(ENUMFILE_eZoneExpansionType, iExpeditionZoneData, 18, row, ExpeditionZoneData, "ZoneExpansion")
    writePublicNameFromDict(DATABLOCK_LightSettings, iExpeditionZoneData, 19, row, ExpeditionZoneData, "LightSettings")
    try:
        writeEnumFromDict(ENUMFILE_eWantedZoneHeighs, iExpeditionZoneData, 20, row, ExpeditionZoneData["AltitudeData"], "AllowedZoneAltitude")
        iExpeditionZoneData.writeFromDict(21, row, ExpeditionZoneData["AltitudeData"], "ChanceToChange")
    except KeyError:pass
    # EventsOnEnter in lists

    try:
        writeEnumFromDict(ENUMFILE_eProgressionPuzzleType, iExpeditionZoneData, colPuzzleType, row, ExpeditionZoneData["ProgressionPuzzleToEnter"], "PuzzleType")
        iExpeditionZoneData.writeFromDict(colPuzzleType+1, row, ExpeditionZoneData["ProgressionPuzzleToEnter"], "CustomText")
        iExpeditionZoneData.writeFromDict(colPuzzleType+2, row, ExpeditionZoneData["ProgressionPuzzleToEnter"], "PlacementCount")
        # ProgressionPuzzleToEnter's ZonePlacementData in lists
    except KeyError:pass
    writePublicNameFromDict(DATABLOCK_ChainedPuzzle, iExpeditionZoneData, colPuzzleType+4, row, ExpeditionZoneData, "ChainedPuzzleToEnter")
    writeEnumFromDict(ENUMFILE_eSecurityGateType, iExpeditionZoneData, colPuzzleType+8, row, ExpeditionZoneData, "SecurityGateToEnter")
    try:
        iExpeditionZoneData.writeFromDict(colPuzzleType+16, row, ExpeditionZoneData["ActiveEnemyWave"], "HasActiveEnemyWave")
        writePublicNameFromDict(DATABLOCK_EnemyGroup, iExpeditionZoneData, colPuzzleType+17, row, ExpeditionZoneData["ActiveEnemyWave"], "EnemyGroupInfrontOfDoor")
        writePublicNameFromDict(DATABLOCK_EnemyGroup, iExpeditionZoneData, colPuzzleType+18, row, ExpeditionZoneData["ActiveEnemyWave"], "EnemyGroupInArea")
        iExpeditionZoneData.writeFromDict(colPuzzleType+19, row, ExpeditionZoneData["ActiveEnemyWave"], "EnemyGroupsInArea")
    except KeyError:pass
    # EnemySpawningInZone in lists
    iExpeditionZoneData.writeFromDict(colPuzzleType+22, row, ExpeditionZoneData, "EnemyRespawning")
    iExpeditionZoneData.writeFromDict(colPuzzleType+23, row, ExpeditionZoneData, "EnemyRespawnRequireOtherZone")
    iExpeditionZoneData.writeFromDict(colPuzzleType+24, row, ExpeditionZoneData, "EnemyRespawnRoomDistance")
    iExpeditionZoneData.writeFromDict(colPuzzleType+25, row, ExpeditionZoneData, "EnemyRespawnTimeInterval")
    iExpeditionZoneData.writeFromDict(colPuzzleType+26, row, ExpeditionZoneData, "EnemyRespawnCountMultiplier")
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
    iExpeditionZoneData.writeFromDict(colHSUClustersInZone+15, row, ExpeditionZoneData, "ForbidTerminalsInZone")
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

def framesLevelLayoutBlock(iExpeditionZoneData:XlsxInterfacer.interface, iExpeditionZoneDataLists:XlsxInterfacer.interface, LevelLayout:dict):
    """
    edit the iExpeditionZoneData and iExpeditionZoneDataLists pandas dataFrames for a single level layout
    """

    ExpeditionZoneDataLists(LevelLayout).write(iExpeditionZoneDataLists)

    row = 2

    for ZoneData in LevelLayout["Zones"]:
        ExpeditionZoneData(iExpeditionZoneData, ZoneData, row)
        row+= 1

class ReactorWaveData:
    """a class that decreases the dimentions of the ReactorWaveData (since the sheet cannot contain 2d-3d data)"""

    def __init__(self, WardenObjective:dict):
        """Generates numerous stubs to be iterated through and written to the interface"""

        self.waves = []
        self.stubEnemyWaves = []
        self.stubEvents = []

        try:
            ReactorWaves = WardenObjective["ReactorWaves"]
        except KeyError:
            return # if there are no reactor waves for the objective, the process of filling them can be skipped

        waveNo = 1
        for wave in ReactorWaves:

            self.waves.append(dict({"WaveNo":waveNo},**wave))

            try:
                enemyWaves = wave["EnemyWaves"]
                for enemyWave in enemyWaves:
                    self.stubEnemyWaves.append(dict({"WaveNo":waveNo},**enemyWave))
            except KeyError:pass # no enemy waves exist, pass

            try:
                events = wave["Events"]
                for event in events:
                    self.stubEvents.append(dict({"WaveNo":waveNo},**event))
            except KeyError:pass # no events exist, pass

            waveNo+= 1


    def write(self, iWardenObjectiveReactorWaves:XlsxInterfacer.interface):
        """Iterates through the numerous stubs and writes them to the specified interface"""
        startrow = 2

        startcolReactorWaves = XlsxInterfacer.ctn("B")
        startcolEnemyWaves = XlsxInterfacer.ctn("K")
        startcolEvents = XlsxInterfacer.ctn("Q")

        # ReactorWaves
        row = startrow
        for Snippet in self.waves:
            iWardenObjectiveReactorWaves.writeFromDict(startcolReactorWaves-1, row, Snippet, "WaveNo")
            iWardenObjectiveReactorWaves.writeFromDict(startcolReactorWaves, row, Snippet, "Warmup")
            iWardenObjectiveReactorWaves.writeFromDict(startcolReactorWaves+1, row, Snippet, "WarmupFail")
            iWardenObjectiveReactorWaves.writeFromDict(startcolReactorWaves+2, row, Snippet, "Wave")
            iWardenObjectiveReactorWaves.writeFromDict(startcolReactorWaves+3, row, Snippet, "Verify")
            iWardenObjectiveReactorWaves.writeFromDict(startcolReactorWaves+4, row, Snippet, "VerifyFail")
            iWardenObjectiveReactorWaves.writeFromDict(startcolReactorWaves+5, row, Snippet, "VerifyInOtherZone")
            writeEnumFromDict(ENUMFILE_eLocalZoneIndex, iWardenObjectiveReactorWaves, startcolReactorWaves+6, row, Snippet, "ZoneForVerification")
            row+= 1

        # EnemyWaves
        row = startrow
        for Snippet in self.stubEnemyWaves:
            iWardenObjectiveReactorWaves.writeFromDict(startcolEnemyWaves-1, row, Snippet, "WaveNo")
            writePublicNameFromDict(DATABLOCK_SurvivalWaveSettings, iWardenObjectiveReactorWaves, startcolEnemyWaves, row, Snippet, "WaveSettings")
            writePublicNameFromDict(DATABLOCK_SurvivalWavePopulation, iWardenObjectiveReactorWaves, startcolEnemyWaves+1, row, Snippet, "WavePopulation")
            iWardenObjectiveReactorWaves.writeFromDict(startcolEnemyWaves+2, row, Snippet, "SpawnTimeRel")
            writeEnumFromDict(ENUMFILE_eReactorWaveSpawnType, iWardenObjectiveReactorWaves, startcolEnemyWaves+3, row, Snippet, "SpawnType")
            row+= 1

        # Events
        row = startrow
        for Snippet in self.stubEvents:
            iWardenObjectiveReactorWaves.writeFromDict(startcolEvents-1, row, Snippet, "WaveNo")
            WardenObjectiveEventData(iWardenObjectiveReactorWaves, Snippet, startcolEvents, row, horizontal=True)
            row+= 1

def framesWardenObjectiveBlock(iWardenObjective:XlsxInterfacer.interface, iWardenObjectiveReactorWaves:XlsxInterfacer.interface, WardenObjective:dict):
    """
    edits the iWardenObjective and iWardenObjectiveReactorWaves pandas dataFrames for a single warden objective
    """
    # set up some checkpoints so if some of the data gets reformatted, not the entire function needs to be altered,
    # just the headings and contents of the section will need edited column values
    rowWavesOnElevatorLand = 25-1
    rowChainedPuzzleToActive = 172-1
    rowLightsOnFromBeginning = 185-1
    rowActivateHSU_ItemFromStart = 206-1
    rowSurvival_TimeToActivate = 252-1
    rowname = 274-1

    writeEnumFromDict(ENUMFILE_eWardenObjectiveType, iWardenObjective, 1, 1, WardenObjective, "Type")
    iWardenObjective.writeFromDict(1, 3, WardenObjective, "Header")
    iWardenObjective.writeFromDict(1, 4, WardenObjective, "MainObjective")
    iWardenObjective.writeFromDict(1, 8, WardenObjective, "FindLocationInfo")
    iWardenObjective.writeFromDict(1, 9, WardenObjective, "FindLocationInfoHelp")
    iWardenObjective.writeFromDict(1, 10, WardenObjective, "GoToZone")
    iWardenObjective.writeFromDict(1, 11, WardenObjective, "GoToZoneHelp")
    iWardenObjective.writeFromDict(1, 12, WardenObjective, "InZoneFindItem")
    iWardenObjective.writeFromDict(1, 13, WardenObjective, "InZoneFindItemHelp")
    iWardenObjective.writeFromDict(1, 14, WardenObjective, "SolveItem")
    iWardenObjective.writeFromDict(1, 15, WardenObjective, "SolveItemHelp")
    iWardenObjective.writeFromDict(1, 16, WardenObjective, "GoToWinCondition_Elevator")
    iWardenObjective.writeFromDict(1, 17, WardenObjective, "GoToWinConditionHelp_Elevator")
    iWardenObjective.writeFromDict(1, 18, WardenObjective, "GoToWinCondition_CustomGeo")
    iWardenObjective.writeFromDict(1, 19, WardenObjective, "GoToWinConditionHelp_CustomGeo")
    iWardenObjective.writeFromDict(1, 20, WardenObjective, "GoToWinCondition_ToMainLayer")
    iWardenObjective.writeFromDict(1, 21, WardenObjective, "GoToWinConditionHelp_ToMainLayer")
    iWardenObjective.writeFromDict(1, 22, WardenObjective, "ShowHelpDelay")

    try:
        GenericEnemyWaveDataList(iWardenObjective, WardenObjective["WavesOnElevatorLand"], 2, rowWavesOnElevatorLand+1, horizontal=True)
    except KeyError:pass
    iWardenObjective.writeFromDict(1, rowWavesOnElevatorLand+44, WardenObjective, "WaveOnElevatorWardenIntel")
    writePublicNameFromDict(DATABLOCK_FogSettings, iWardenObjective, 1, rowWavesOnElevatorLand+46, WardenObjective, "FogTransitionDataOnElevatorLand")
    iWardenObjective.writeFromDict(1, rowWavesOnElevatorLand+47, WardenObjective, "FogTransitionDurationOnElevatorLand")
    try:
        GenericEnemyWaveDataList(iWardenObjective, WardenObjective["WavesOnActivate"], 2, rowWavesOnElevatorLand+52, horizontal=True)
    except KeyError:pass
    try:
        itercol,iterrow = 2, rowWavesOnElevatorLand+60
        for event in WardenObjective["EventsOnActivate"]:
            WardenObjectiveEventData(iWardenObjective, event, itercol, iterrow, horizontal=False)
            itercol+= 1
    except KeyError:pass
    iWardenObjective.writeFromDict(1, rowWavesOnElevatorLand+100, WardenObjective, "StopAllWavesBeforeGotoWin")
    try:
        GenericEnemyWaveDataList(iWardenObjective, WardenObjective["WavesOnGotoWin"], 2, rowWavesOnElevatorLand+103, horizontal=True)
    except KeyError:pass
    writeEnumFromDict(ENUMFILE_eRetrieveExitWaveTrigger, iWardenObjective, 1, rowWavesOnElevatorLand+104, WardenObjective, "WaveOnGotoWinTrigger")
    try:
        itercol,iterrow = 2, rowWavesOnElevatorLand+112
        for event in WardenObjective["EventsOnGotoWin"]:
            WardenObjectiveEventData(iWardenObjective, event, itercol, iterrow, horizontal=False)
            itercol+= 1
    except KeyError:pass
    writePublicNameFromDict(DATABLOCK_FogSettings, iWardenObjective, 1, rowWavesOnElevatorLand+143, WardenObjective, "FogTransitionDataOnGotoWin")
    iWardenObjective.writeFromDict(1, rowWavesOnElevatorLand+144, WardenObjective, "FogTransitionDurationOnGotoWin")

    writePublicNameFromDict(DATABLOCK_ChainedPuzzle, iWardenObjective, 1, rowChainedPuzzleToActive, WardenObjective, "ChainedPuzzleToActive")
    writePublicNameFromDict(DATABLOCK_ChainedPuzzle, iWardenObjective, 1, rowChainedPuzzleToActive+1, WardenObjective, "ChainedPuzzleMidObjective")
    writePublicNameFromDict(DATABLOCK_ChainedPuzzle, iWardenObjective, 1, rowChainedPuzzleToActive+2, WardenObjective, "ChainedPuzzleAtExit")
    iWardenObjective.writeFromDict(1, rowChainedPuzzleToActive+3, WardenObjective, "ChainedPuzzleAtExitScanSpeedMultiplier")
    iWardenObjective.writeFromDict(1, rowChainedPuzzleToActive+5, WardenObjective, "Gather_RequiredCount")
    writePublicNameFromDict(DATABLOCK_Item, iWardenObjective, 1, rowChainedPuzzleToActive+6, WardenObjective, "Gather_ItemId")
    iWardenObjective.writeFromDict(1, rowChainedPuzzleToActive+7, WardenObjective, "Gather_SpawnCount")
    iWardenObjective.writeFromDict(1, rowChainedPuzzleToActive+8, WardenObjective, "Gather_MaxPerZone")
    try:
        itercol,iterrow = 1, rowChainedPuzzleToActive+10
        for item in WardenObjective["Retrieve_Items"]:
            if str(item) == "0": continue # don't attempt to try to find blocks with id zero
            iWardenObjective.write(DatablockIO.idToName(DATABLOCK_Item, item), itercol, iterrow)
            itercol+= 1
    except KeyError:pass
    # XXX in order to run the test; don't run unfixed frame writer
    # ReactorWaveData(WardenObjective).write(iWardenObjectiveReactorWaves)

    iWardenObjective.writeFromDict(1, rowLightsOnFromBeginning, WardenObjective, "LightsOnFromBeginning")
    iWardenObjective.writeFromDict(1, rowLightsOnFromBeginning+1, WardenObjective, "LightsOnDuringIntro")
    iWardenObjective.writeFromDict(1, rowLightsOnFromBeginning+2, WardenObjective, "LightsOnWhenStartupComplete")
    iWardenObjective.writeFromDict(1, rowLightsOnFromBeginning+5, WardenObjective, "SpecialTerminalCommand")
    iWardenObjective.writeFromDict(1, rowLightsOnFromBeginning+6, WardenObjective, "SpecialTerminalCommandDesc")
    try:
        # TODO if an output string is blank, it will not be placed on the list which may leave a hole in the list
        itercol,iterrow = 1, rowLightsOnFromBeginning+7
        for output in WardenObjective["PostCommandOutput"]:
            iWardenObjective.write(output, itercol, iterrow)
            itercol+= 1
    except KeyError:pass
    iWardenObjective.writeFromDict(1, rowLightsOnFromBeginning+10, WardenObjective, "PowerCellsToDistribute")
    iWardenObjective.writeFromDict(1, rowLightsOnFromBeginning+12, WardenObjective, "Uplink_NumberOfVerificationRounds")
    iWardenObjective.writeFromDict(1, rowLightsOnFromBeginning+13, WardenObjective, "Uplink_NumberOfTerminals")
    iWardenObjective.writeFromDict(1, rowLightsOnFromBeginning+16, WardenObjective, "CentralPowerGenClustser_NumberOfGenerators")
    iWardenObjective.writeFromDict(1, rowLightsOnFromBeginning+17, WardenObjective, "CentralPowerGenClustser_NumberOfPowerCells")
    try:
        itercol,iterrow = 1,rowLightsOnFromBeginning+19
        for step in WardenObjective["CentralPowerGenClustser_FogDataSteps"]:
            GeneralFogDataStep(iWardenObjective, step, itercol, iterrow, horizontal=False)
            itercol+= 1
    except KeyError:pass

    writePublicNameFromDict(DATABLOCK_Item, iWardenObjective, 1, rowActivateHSU_ItemFromStart, WardenObjective, "ActivateHSU_ItemFromStart")
    writePublicNameFromDict(DATABLOCK_Item, iWardenObjective, 1, rowActivateHSU_ItemFromStart+1, WardenObjective, "ActivateHSU_ItemAfterActivation")
    iWardenObjective.writeFromDict(1, rowActivateHSU_ItemFromStart+4, WardenObjective, "ActivateHSU_StopEnemyWavesOnActivation")
    iWardenObjective.writeFromDict(1, rowActivateHSU_ItemFromStart+5, WardenObjective, "ActivateHSU_ObjectiveCompleteAfterInsertion")
    iWardenObjective.writeFromDict(1, rowActivateHSU_ItemFromStart+6, WardenObjective, "ActivateHSU_RequireItemAfterActivationInExitScan")
    try:
        itercol,iterrow = 2, rowActivateHSU_ItemFromStart+9
        for event in WardenObjective["ActivateHSU_Events"]:
            WardenObjectiveEventData(iWardenObjective, event, itercol, iterrow, horizontal=False)
    except KeyError:pass

    iWardenObjective.writeFromDict(1, rowname, WardenObjective, "name")
    iWardenObjective.writeFromDict(1, rowname+1, WardenObjective, "internalEnabled")
    iWardenObjective.writeFromDict(1, rowname+2, WardenObjective, "persistentID")


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
        return [],None,"",None

    # convert numerical tier to A-E
    try: levelTier = chr(65+int(levelTier))
    except: pass
    # make sure the tier letter is upper cased
    levelTier = levelTier.upper()

    # make sure the levelIndex is a number
    try:
        levelIndex = int(levelIndex)
    except ValueError:
        return [],None,"",None

    rundownIndex = rundownValueToIndex(RundownDataBlock,rundown)
    # if no such rundown exists, return a blank array
    if rundownIndex in [-1, None]:
        return [],None,"",None

    # get the persistentID of the rundown
    rundown = RundownDataBlock.data["Blocks"][rundownIndex]["persistentID"]

    try:
        return RundownDataBlock.data["Blocks"][rundownIndex]["Tier"+levelTier][levelIndex],rundown,"Tier"+levelTier,levelIndex
    except KeyError:
        return [],None,"",None
    except IndexError:
        return [],None,"",None

def UtilityJob(desiredReverse:str, RundownDataDataBlock:DatablockIO.datablock, LevelLayoutDataBlock:DatablockIO.datablock, WardenObjectiveDataBlock:DatablockIO.datablock, targetdir:os.PathLike, logger:logging.Logger=None):
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

    if logger == None:
        logger = logging.getLevelName("none")
        logger.addHandler(logging.NullHandler())
        logger.propagate = False

    logger.info("Starting reverse utilty job: \""+desiredReverse+"\"")

    # check the template version before doing any searching
    try:
        iKey = XlsxInterfacer.interface(pandas.read_excel(templatepath, "Key", header=None))
    except FileNotFoundError as e:
        raise FileNotFoundError("Missing template file: " + str(e))
    if iKey.read(str, 0, 5).split(".")[0:2] != Version.split(".")[0:2]:
        logger.critical("Version mismatch between utility and tempalte sheet, incompatible.")
        return False

    ExpeditionInTierData, rundown, levelTier, levelIndex = getExpeditionInTierData(desiredReverse, RundownDataDataBlock)

    if (rundown == None or levelTier == "" or levelIndex == None or ExpeditionInTierData==[]):
        # if no such level exists
        raise Exception("The search for \""+desiredReverse+"\" matched no level.")

    try:
        # get the name of the level if it exists (so the file name can be the name of the level)
        levelName = RundownDataDataBlock.data["Blocks"][RundownDataDataBlock.find(rundown)][levelTier][levelIndex]["Descriptive"]["PublicName"]
        logger.info("The search for \""+desiredReverse+"\" found:\t"+str(rundown)+" "+levelTier+" "+str(levelIndex)+" "+levelName)
    except KeyError:
        levelName = desiredReverse
        logger.info("The search for \""+desiredReverse+"\" found a nameless level:\t"+str(rundown)+" "+levelTier+" "+str(levelIndex))

    strippedLevelName = re.sub(winreserveregex, "", levelName)
    writepath = os.path.join(targetdir, strippedLevelName+".xlsx")
    try:
        shutil.copy(templatepath,writepath)
        fxlsx = open(writepath, 'rb+')
    except PermissionError:
        raise PermissionError("PermissionError opening \""+writepath+"\", is it open elsewhere?")

    iMeta = XlsxInterfacer.interface(pandas.read_excel(fxlsx, "Meta", header=None))
    iExpeditionInTier = XlsxInterfacer.interface(pandas.read_excel(fxlsx, "ExpeditionInTier", header=None))
    try:
        id_ = ExpeditionInTierData["LevelLayoutData"]
        if id_ == 0:raise KeyError
        LayerDataL1 = LevelLayoutDataBlock.data["Blocks"][LevelLayoutDataBlock.find(id_)]
        iExpeditionZoneDataL1 = XlsxInterfacer.interface(pandas.read_excel(fxlsx, "LX ExpeditionZoneData", header=None))
        iExpeditionZoneDataListsL1 = XlsxInterfacer.interface(pandas.read_excel(fxlsx, "LX ExpeditionZoneData Lists", header=None))
        logger.debug("Found L1 LevelLayout")
    except KeyError:
        logger.debug("No L1 LevelLayout")
    try:
        id_ = ExpeditionInTierData["SecondaryLayout"]
        if id_ == 0:raise KeyError
        LayerDataL2 = LevelLayoutDataBlock.data["Blocks"][LevelLayoutDataBlock.find(id_)]
        iExpeditionZoneDataL2 = XlsxInterfacer.interface(pandas.read_excel(fxlsx, "LX ExpeditionZoneData", header=None))
        iExpeditionZoneDataListsL2 = XlsxInterfacer.interface(pandas.read_excel(fxlsx, "LX ExpeditionZoneData Lists", header=None))
        logger.debug("Found L2 LevelLayout")
    except KeyError:
        logger.debug("No L2 LevelLayout")
    try:
        id_ = ExpeditionInTierData["ThirdLayout"]
        if id_ == 0:raise KeyError
        LayerDataL3 = LevelLayoutDataBlock.data["Blocks"][LevelLayoutDataBlock.find(id_)]
        iExpeditionZoneDataL3 = XlsxInterfacer.interface(pandas.read_excel(fxlsx, "LX ExpeditionZoneData", header=None))
        iExpeditionZoneDataListsL3 = XlsxInterfacer.interface(pandas.read_excel(fxlsx, "LX ExpeditionZoneData Lists", header=None))
        logger.debug("Found L3 LevelLayout")
    except KeyError:
        logger.debug("No L3 LevelLayout")
    try:
        id_ = ExpeditionInTierData["MainLayerData"]["ObjectiveData"]["DataBlockId"]
        if id_ == 0:raise KeyError
        WardenObjectiveL1 = WardenObjectiveDataBlock.data["Blocks"][WardenObjectiveDataBlock.find(id_)]
        iWardenObjectiveL1 = XlsxInterfacer.interface(pandas.read_excel(fxlsx, "LX WardenObjective", header=None))
        iWardenObjectiveReactorWavesL1 = XlsxInterfacer.interface(pandas.read_excel(fxlsx, "LX WardenObjective ReactorWaves", header=None))
        logger.debug("Found L1 WardenObjective")
    except KeyError:
        logger.debug("No L1 WardenObjective")
    try:
        id_ = ExpeditionInTierData["SecondaryLayerData"]["ObjectiveData"]["DataBlockId"]
        if id_ == 0:raise KeyError
        WardenObjectiveL2 = WardenObjectiveDataBlock.data["Blocks"][WardenObjectiveDataBlock.find(id_)]
        iWardenObjectiveL2 = XlsxInterfacer.interface(pandas.read_excel(fxlsx, "LX WardenObjective", header=None))
        iWardenObjectiveReactorWavesL2 = XlsxInterfacer.interface(pandas.read_excel(fxlsx, "LX WardenObjective ReactorWaves", header=None))
        logger.debug("Found L2 WardenObjective")
    except KeyError:
        logger.debug("No L2 WardenObjective")
    try:
        id_ = ExpeditionInTierData["ThirdLayerData"]["ObjectiveData"]["DataBlockId"]
        if id_ == 0:raise KeyError
        WardenObjectiveL3 = WardenObjectiveDataBlock.data["Blocks"][WardenObjectiveDataBlock.find(id_)]
        iWardenObjectiveL3 = XlsxInterfacer.interface(pandas.read_excel(fxlsx, "LX WardenObjective", header=None))
        iWardenObjectiveReactorWavesL3 = XlsxInterfacer.interface(pandas.read_excel(fxlsx, "LX WardenObjective ReactorWaves", header=None))
        logger.debug("Found L3 WardenObjective")
    except KeyError:
        logger.debug("No L3 WardenObjective")
    fxlsx.close()

    frameMeta(iMeta, rundown, levelTier, levelIndex)
    frameExpeditionInTier(iExpeditionInTier, ExpeditionInTierData)
    try:
        framesLevelLayoutBlock(iExpeditionZoneDataL1, iExpeditionZoneDataListsL1, LayerDataL1)
        logger.debug("Finished L1 LevelLayout")
    except NameError:pass
    except Exception as e:
        logger.error("Problem writing L1 LevelLayout (skipping layout): "+str(e))
        logger.debug(e, exc_info=True)
    try:
        framesLevelLayoutBlock(iExpeditionZoneDataL2, iExpeditionZoneDataListsL2, LayerDataL2)
        logger.debug("Finished L2 LevelLayout")
    except NameError:pass
    except Exception as e:
        logger.error("Problem writing L2 LevelLayout (skipping layout): "+str(e))
        logger.debug(e, exc_info=True)
    try:
        framesLevelLayoutBlock(iExpeditionZoneDataL3, iExpeditionZoneDataListsL3, LayerDataL3)
        logger.debug("Finished L2 LevelLayout")
    except NameError:pass
    except Exception as e:
        logger.error("Problem writing L3 LevelLayout (skipping layout): "+str(e))
        logger.debug(e, exc_info=True)

    try:
        framesWardenObjectiveBlock(iWardenObjectiveL1, iWardenObjectiveReactorWavesL1, WardenObjectiveL1)
        logger.debug("Finished L1 WardenObjective")
    except NameError:pass
    except Exception as e:
        logger.error("Problem writing L1 WardenObjective (skipping objective): "+str(e))
        logger.debug(e, exc_info=True)
    try:
        framesWardenObjectiveBlock(iWardenObjectiveL2, iWardenObjectiveReactorWavesL2, WardenObjectiveL2)
        logger.debug("Finished L2 WardenObjective")
    except NameError:pass
    except Exception as e:
        logger.error("Problem writing L2 WardenObjective (skipping objective): "+str(e))
        logger.debug(e, exc_info=True)
    try:
        framesWardenObjectiveBlock(iWardenObjectiveL3, iWardenObjectiveReactorWavesL3, WardenObjectiveL3)
        logger.debug("Finished L3 WardenObjective")
    except NameError:pass
    except Exception as e:
        logger.error("Problem writing L3 WardenObjective (skipping objective): "+str(e))
        logger.debug(e, exc_info=True)


    workbook = openpyxl.load_workbook(filename = writepath)

    # NOTE openpyxl does not copy data validation so all of the sheets that were copied lose their data validation, the bodge for this is just to have the template already have the layer sheets copied despite not being elegant
    # workbook["LX ExpeditionZoneData Lists"].title = "a" # this weird renaming is used to avoid UserWarnings by openpyxl because "LX ExpeditionZoneData Lists Copy" is too long
    # try:
    #     _ = LayerDataL1
    #     workbook.copy_worksheet(workbook["LX ExpeditionZoneData"]).title = "L1 ExpeditionZoneData"
    #     workbook.copy_worksheet(workbook["a"]).title = "L1 ExpeditionZoneData Lists"
    # except NameError:pass
    # try:
    #     _ = LayerDataL2
    #     workbook.copy_worksheet(workbook["LX ExpeditionZoneData"]).title = "L2 ExpeditionZoneData"
    #     workbook.copy_worksheet(workbook["a"]).title = "L2 ExpeditionZoneData Lists"
    # except NameError:pass
    # try:
    #     _ = LayerDataL3
    #     workbook.copy_worksheet(workbook["LX ExpeditionZoneData"]).title = "L3 ExpeditionZoneData"
    #     workbook.copy_worksheet(workbook["a"]).title = "L3 ExpeditionZoneData Lists"
    # except NameError:pass
    # workbook["a"].title = "LX ExpeditionZoneData Lists"
    # workbook["LX WardenObjective ReactorWaves"].title = "a" # this weird renaming is used to avoid UserWarnings by openpyxl because "LX ExpeditionZoneData Lists Copy" is too long
    # try:
    #     _ = WardenObjectiveL1
    #     workbook.copy_worksheet(workbook["LX WardenObjective"]).title = "L1 WardenObjective"
    #     workbook.copy_worksheet(workbook["a"]).title = "L1 WardenObjective ReactorWaves"
    # except NameError:pass
    # try:
    #     _ = WardenObjectiveL2
    #     workbook.copy_worksheet(workbook["LX WardenObjective"]).title = "L2 WardenObjective"
    #     workbook.copy_worksheet(workbook["a"]).title = "L2 WardenObjective ReactorWaves"
    # except NameError:pass
    # try:
    #     _ = WardenObjectiveL3
    #     workbook.copy_worksheet(workbook["LX WardenObjective"]).title = "L3 WardenObjective"
    #     workbook.copy_worksheet(workbook["a"]).title = "L3 WardenObjective ReactorWaves"
    # except NameError:pass
    # workbook["a"].title = "LX WardenObjective ReactorWaves"

    # NOTE openpyxl does not copy data validation, so lists that are longer than 20 cells could have the color formatting copied but would not copy the validation
    # as such, it is better the user see the cell has no formatting than think the cell does
    workbook.remove(workbook["LX ExpeditionZoneData"])
    workbook.remove(workbook["LX ExpeditionZoneData Lists"])
    workbook.remove(workbook["LX WardenObjective"])
    workbook.remove(workbook["LX WardenObjective ReactorWaves"])

    workbook.save(filename = writepath)

    logger.debug("Formatted template sheets copied")


    # NOTE using interface.save() can take a while (comparatively to other portions of the utility)
    iMeta.save(writepath, "Meta")
    iExpeditionInTier.save(writepath, "ExpeditionInTier")
    try:
        _ = LayerDataL1
        iExpeditionZoneDataL1.save(writepath, "L1 ExpeditionZoneData")
        iExpeditionZoneDataListsL1.save(writepath, "L1 ExpeditionZoneData Lists")
    except NameError:pass
    try:
        _ = LayerDataL2
        iExpeditionZoneDataL2.save(writepath, "L2 ExpeditionZoneData")
        iExpeditionZoneDataListsL2.save(writepath, "L2 ExpeditionZoneData Lists")
    except NameError:pass
    try:
        _ = LayerDataL3
        iExpeditionZoneDataL3.save(writepath, "L3 ExpeditionZoneData")
        iExpeditionZoneDataListsL3.save(writepath, "L3 ExpeditionZoneData Lists")
    except NameError:pass
    try:
        _ = WardenObjectiveL1
        iWardenObjectiveL1.save(writepath, "L1 WardenObjective")
        iWardenObjectiveReactorWavesL1.save(writepath, "L1 WardenObjective ReactorWaves")
    except NameError:pass
    try:
        _ = WardenObjectiveL2
        iWardenObjectiveL2.save(writepath, "L2 WardenObjective")
        iWardenObjectiveReactorWavesL2.save(writepath, "L2 WardenObjective ReactorWaves")
    except NameError:pass
    try:
        _ = WardenObjectiveL3
        iWardenObjectiveL3.save(writepath, "L3 WardenObjective")
        iWardenObjectiveReactorWavesL3.save(writepath, "L3 WardenObjective ReactorWaves")
    except NameError:pass

    logger.debug("Data written to sheets")

    logger.info("Finished reverse utilty job: \""+desiredReverse+"\"")
    return True

def SearchJob(desiredReverse:str, RundownDataBlock, LevelLayoutDataBlock, WardenObjectiveDataBlock, logger:logging.Logger=None):
    """
    Secondary job meant specifically to search for and display the search result for a level
    """
    if logger == None:
        logger = logging.getLevelName("none")
        logger.addHandler(logging.NullHandler())
        logger.propagate = False
    # TODO finish the SearchJob
    logger.critical("SearchJob is not written")

def main():
    parser = argparse.ArgumentParser(
        prog="DPK LevelReverseUtility",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        description=textwrap.dedent("""\
        This is a tool created by DPK.
        This tool can convert a level from the GTFO Datablocks into excel templated files.
        This tool searches the blocks for the levels you want to convert, search terms are formatted like below (including quotes for names with spaces):
            "Cuernos"
            "Contact,Cuernos"
            "Contact,C,2"
            "Rundown 004 - EA,C,2"
            "Rundown 004 - EA,2,2"
            "25,2,2"
        """)
    )

    parser.add_argument('terms', type=str, nargs='*', help='Search term(s) for which levels to convert.')
    parser.add_argument('-n', "--noinput", action='store_true', help='[N]o inputs (which could be annoying in CLI and scripts)')
    parser.add_argument('-v', "--verbosity", type=str.upper, help='Changes console [v]erbosity', choices=['DEBUG','INFO','WARNING','ERROR','CRITICAL'], default='INFO')
    parser.add_argument('-d', "--directory", type=str, help='Directory to write to')

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

    logger = logging.getLogger("LevelReverseUtilty")
    logger.setLevel(logging.DEBUG)

    logfilehandler = logging.FileHandler(os.path.join(logdir,time.strftime("%Y.%m.%d.%H.%M.%S",time.gmtime())+".LevelReverseUtility.log"))
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

    logger.debug("Running DPK's LevelReverseUtilty with the given arguments:\n\t"+str(args))

    # Open Datablocks to get level from
    try:
        RundownDataBlock = DatablockIO.datablock(open(os.path.join(blockpath,ConfigManager.config["Project"]["blockprefix"]+"RundownDataBlock"+ConfigManager.config["Project"]["blocksuffix"]), 'r', encoding="utf8"))
        LevelLayoutDataBlock = DatablockIO.datablock(open(os.path.join(blockpath,ConfigManager.config["Project"]["blockprefix"]+"LevelLayoutDataBlock"+ConfigManager.config["Project"]["blocksuffix"]), 'r', encoding="utf8"))
        WardenObjectiveDataBlock = DatablockIO.datablock(open(os.path.join(blockpath,ConfigManager.config["Project"]["blockprefix"]+"WardenObjectiveDataBlock"+ConfigManager.config["Project"]["blocksuffix"]), 'r', encoding="utf8"))
    except FileNotFoundError as e:
        if not args.noinput:
            print("Missing a DataBlock: " + str(e))
            input()
        raise FileNotFoundError("Missing a DataBlock: " + str(e))

    if not(args.directory):
        args.directory = os.path.dirname(__file__)

    anythingDone = False

    if args.terms == [] and not(args.noinput):
        print(parser.description)
        logger.info("No term arguments given, entering interactive mode.")
        print("Input search terms below and leave line blank to continue.")
        inputterm = input()
        while inputterm != "":
            # regex substitute to remove (double or single) quotes on the outside of input terms in interactive mode
            args.terms.append(re.sub("^(\"|\')(.*)\\1$","\\2",inputterm))
            inputterm = input()
        logger.debug("New arguments after interactive mode.:\n\t"+str(args))

    for desiredReverse in args.terms:
        logger.info("Working with: \""+desiredReverse+"\"")
        try:
            if UtilityJob(desiredReverse, RundownDataBlock, LevelLayoutDataBlock, WardenObjectiveDataBlock, args.directory, logger=joblogger):
                logger.info("Finished with: \""+desiredReverse+"\"")
                anythingDone = True
            else:
                logger.info("Failed with: \""+desiredReverse+"\"")
        except Exception as e:
            logger.error("Exception with: \""+desiredReverse+"\"\n\t"+str(e))
            logger.debug(e, exc_info=True)

    # TODO allow for a secondary job just to search for a level but not run the reverse tool on it

    if not anythingDone:logger.warning("Nothing happened... are you sure you didn't do anything wrong?")
    logger.info("Done.")
    if not(args.noinput):waitUser()

if __name__ == "__main__":
    main()
