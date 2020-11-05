"""
This is a tool created by DPK

This tool can convert xlsx to a bunch of GTFO Datablock pieces and also convert levels from the Datablocks back into the templated form
Template: https://docs.google.com/spreadsheets/d/1FLA-eHv9NhU3IdxcdQ29ueW8Qav9IlknaLdb4xbGZCI/edit?usp=sharing
"""

import DatablockIO
import EnumConverter
import XlsxInterfacer

import pandas
import xlrd
import numpy
import re
import json

# pandas:   used to read the excel data
# xlrd:     used to catch and throw excel errors when initially reading the sheets
# numpy:    used to manipulate the inconsistant numpy data read by pandas
# re:       used to preform regex searches
# json:     used to export the data to a json

# a regex to capture the newlines the devs put into their json
devnewlnregex = "(\\\\n|\\\\r){1,2}"

# Settings
#####
# Version number meaning:
# R.G.S
# R: Rundown
# G: Generator
# S: Sheet (minor changes to the sheet are insignificant to the generator)
Version = "4.1"
# relative path to location for datablocks, defaultly its folder should be on the same layer as this project's folder
blockpath = "../Datablocks/"
#####

def EnsureKeyInDictArray(dictionary, key):
    """this function will ensure that an array exists in a key if there is not already a value"""
    try:_ = dictionary[key]
    except KeyError:dictionary[key] = []


# load all datablock files
if True:
    DATABLOCK_Rundown = DatablockIO.datablock(open(blockpath+"RundownDataBlock.json", 'r', encoding="utf8"))
    DATABLOCK_LevelLayout = DatablockIO.datablock(open(blockpath+"LevelLayoutDataBlock.json", 'r', encoding="utf8"))
    DATABLOCK_WardenObjective = DatablockIO.datablock(open(blockpath+"WardenObjectiveDataBlock.json", 'r', encoding="utf8"))
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

def ZonePlacementData(interface, col, row, horizontal=True):
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

def BulkheadDoorPlacementData(interface, col, row, horizontal=False):
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

def FunctionPlacementData(interface, col, row, horizontal=True):
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

def ZonePlacementWeightsList(interface, col, row, horizontal=True):
    """
    return a list of lists containing ZonePlacementWeights \n
    col and row define the upper left value (not header) \n
    horizontal describes the direction to iterate to look for waves
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

def ZonePlacementWeights(interface, col, row, horizontal=True):
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

def GenericEnemyWaveData(interface, col, row, horizontal=False):
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
    interface.readIntoDict(bool, col+3*horizontal, row+3*(not horizontal), "TriggerAlarm")
    interface.readIntoDict(str, col+4*horizontal, row+4*(not horizontal), "IntelMessage")
    return data

def GenericEnemyWaveDataList(interface, col, row, horizontal=True):
    """
    return a GenericEnemyWaveData dict \n
    col and row define the upper left value (not header) \n
    horizontal describes the direction to iterate to look for waves
    """
    datalist = []
    while not(interface.isEmpty(col,row)):
        # the direction of the set of waves and values in the data are perpendicular
        datalist+= GenericEnemyWaveData(interface, col, row, horizontal=(not horizontal))
        col+= horizontal
        row+= not horizontal

    return datalist

def WardenObjectiveEventData(interface, col, row, horizontal=False):
    """
    return a WardenObjectiveEventData dict \n
    col and row define the upper left value (not header) \n
    horizontal is true if the values are in the same row
    """
    data = {}
    interface(str, col, row, data, "Trigger")
    EnumConverter.enumInDict(ENUMFILE_eWardenObjectiveEventTrigger, data, "Trigger")
    interface(str, col+horizontal, row+(not horizontal), data, "Type")
    EnumConverter.enumInDict(ENUMFILE_eWardenObjectiveEventType, data, "Type")
    interface(str, col+2*horizontal, row+2*(not horizontal), data, "Layer")
    EnumConverter.enumInDict(ENUMFILE_LG_LayerType, data, "Layer")
    interface(str, col+3*horizontal, row+3*(not horizontal), data, "LocalIndex")
    EnumConverter.enumInDict(ENUMFILE_eLocalZoneIndex, data, "LocalIndex")
    interface(float, col+4*horizontal, row+4*(not horizontal), data, "Delay")
    interface(str, col+5*horizontal, row+5*(not horizontal), data, "WardenIntel")
    return data

def LayerData(interface, col, row):
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
        data["ZonesWithBulkheadEntrance"].append(EnumConverter.enumToIndex(ENUMFILE_eLocalZoneIndex, interface.read(str, itercol, iterrow)))
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
    return data

def ExpeditionInTier(iExpeditionInTier):
    """returns the expedition in tier piece to be inserted into the rundown data block"""
    data = {}
    data["Enabled"] = iExpeditionInTier.read(bool, 0, 2)
    data["Accessibility"] = iExpeditionInTier.read(str, 1, 2)
    EnumConverter.enumInDict(ENUMFILE_eExpeditionAccessibility, data, "Accessibility")
    data["Descriptive"] = {}
    data["Descriptive"]["Prefix"] = iExpeditionInTier.read(str, 10, 2)
    data["Descriptive"]["PublicName"] = iExpeditionInTier.read(str, 10, 3)
    iExpeditionInTier.readIntoDict(int, 10, 4, data["Descriptive"], "ExpeditionDepth")
    data["Descriptive"]["EstimatedDuration"] = iExpeditionInTier.read(XlsxInterfacer.blankable, 10, 5)
    data["Descriptive"]["ExpeditionDescription"] = re.sub(devnewlnregex,"\r\n", iExpeditionInTier.read(XlsxInterfacer.blankable, 10, 6))
    data["Descriptive"]["RoleplayedWardenIntel"] = re.sub(devnewlnregex,"\r\n", iExpeditionInTier.read(XlsxInterfacer.blankable, 10, 7))
    data["Descriptive"]["DevInfo"] = re.sub(devnewlnregex,"\n", iExpeditionInTier.read(XlsxInterfacer.blankable, 10, 8))
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
    data["SecondaryLayerData"] = LayerData(iExpeditionInTier, 0, 42)
    iExpeditionInTier.readIntoDict(bool, 7, 13, data, "ThirdLayerEnabled")
    iExpeditionInTier.readIntoDict(int, 8, 13, data, "ThirdLayout")
    data["BuildThirdFrom"] = {}
    iExpeditionInTier.readIntoDict(str, 7, 17, data["BuildThirdFrom"], "LayerType")
    EnumConverter.enumInDict(ENUMFILE_LG_LayerType, data["BuildThirdFrom"], "LayerType")
    iExpeditionInTier.readIntoDict(str, 8, 17, data["BuildThirdFrom"], "Zone")
    EnumConverter.enumInDict(ENUMFILE_eLocalZoneIndex, data["BuildThirdFrom"], "Zone")
    data["ThirdLayerData"] = LayerData(iExpeditionInTier, 0, 64)
    data["SpecialOverrideData"] = {}
    iExpeditionInTier.readIntoDict(float, 5, 6, data["SpecialOverrideData"], "WeakResourceContainerWithPackChanceForLocked")
    return data

def LevelLayoutBlock(iExpeditionZoneData, iExpeditionZoneDataLists):
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

def ExpeditionZoneData(iExpeditionZoneData, listdata, row):
    """returns the ExpeditionZoneData for a particular row"""
    # set up some checkpoints so if some of the data gets reformatted, not the entire function needs to be altered,
    # just the headings and contents of the section will need edited column values
    colPuzzleType = XlsxInterfacer.ctn("Q")
    colHSUClustersInZone = XlsxInterfacer.ctn("AB")
    colHealthMulti = XlsxInterfacer.ctn("AR")

    data = {}

    zonestr = iExpeditionZoneData.read(str, 0, row)
    data["LocalIndex"] =  EnumConverter.enumToIndex(ENUMFILE_eLocalZoneIndex, zonestr)

    iExpeditionZoneData.readIntoDict(int, 1, row, data, "SubSeed")
    iExpeditionZoneData.readIntoDict(int, 2, row, data, "BulkheadDCScanSeed")
    iExpeditionZoneData.readIntoDict(str, 3, row, data, "SubComplex")
    EnumConverter.enumInDict(ENUMFILE_SubComplex, data, "SubComplex")
    iExpeditionZoneData.readIntoDict(str, 4, row, data, "CustomGeomorph")
    data["CoverageMinMax"] = {}
    iExpeditionZoneData.readIntoDict(int, 5, row, data["CoverageMinMax"], "x")
    iExpeditionZoneData.readIntoDict(int, 6, row, data["CoverageMinMax"], "y")
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
    iExpeditionZoneData.readIntoDict(int, colPuzzleType+8, row, data["ActiveEnemyWave"], "EnemyGroupsInArea")
    data["EnemySpawningInZone"] = listdata.EnemySpawningInZone(zonestr)

    iExpeditionZoneData.readIntoDict(int, colHSUClustersInZone, row, data, "HSUClustersInZone")
    iExpeditionZoneData.readIntoDict(int, colHSUClustersInZone+1, row, data, "CorpseClustersInZone")
    iExpeditionZoneData.readIntoDict(int, colHSUClustersInZone+2, row, data, "ResourceContainerClustersInZone")
    iExpeditionZoneData.readIntoDict(int, colHSUClustersInZone+3, row, data, "GeneratorClustersInZone")
    iExpeditionZoneData.readIntoDict(str, colHSUClustersInZone+4, row, data, "CorpsesInZone")
    EnumConverter.enumInDict(ENUMFILE_eZoneDistributionAmount, data, "CorpsesInZone")
    iExpeditionZoneData.readIntoDict(str, colHSUClustersInZone+5, row, data, "HSUsInZone")
    EnumConverter.enumInDict(ENUMFILE_eZoneDistributionAmount, data, "HSUsInZone")
    iExpeditionZoneData.readIntoDict(str, colHSUClustersInZone+6, row, data, "DeconUnitsInZone")
    EnumConverter.enumInDict(ENUMFILE_eZoneDistributionAmount, data, "DeconUnitsInZone")
    iExpeditionZoneData.readIntoDict(bool, colHSUClustersInZone+7, row, data, "AllowSmallPickupsAllocation")
    iExpeditionZoneData.readIntoDict(bool, colHSUClustersInZone+8, row, data, "AllowResourceContainerAllocation")
    iExpeditionZoneData.readIntoDict(bool, colHSUClustersInZone+9, row, data, "ForceBigPickupsAllocation")
    iExpeditionZoneData.readIntoDict(str, colHSUClustersInZone+10, row, data, "ConsumableDistributionInZone")
    DatablockIO.nameInDict(DATABLOCK_ConsumableDistribution, data, "ConsumableDistributionInZone")
    iExpeditionZoneData.readIntoDict(str, colHSUClustersInZone+11, row, data, "BigPickupDistributionInZone")
    DatablockIO.nameInDict(DATABLOCK_ConsumableDistribution, data, "BigPickupDistributionInZone")
    data["TerminalPlacements"] = listdata.TerminalPlacements(zonestr)
    iExpeditionZoneData.readIntoDict(bool, colHSUClustersInZone+12, row, data, "ForbidTerminalsInZone")
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

class ExpeditionZoneDataLists:
    """a class that contains a dictionaries for the ExpeditionZoneData (since the sheet cannot contain 2d-3d data)"""

    def __init__(self, iExpeditionZoneDataLists):
        """Generates numerous stubs that can have zone specific data request from the object getters"""
        startrow = 2

        startcolEventsOnEnter = XlsxInterfacer.ctn("A")
        startcolProgressionPuzzleToEnter = XlsxInterfacer.ctn("K")
        startcolEnemySpawningInZone = XlsxInterfacer.ctn("R")
        startcolTerminalPlacements = XlsxInterfacer.ctn("Y")
        startcolLocalLogFiles = XlsxInterfacer.ctn("AJ")
        startcolPowerGeneratorPlacements = XlsxInterfacer.ctn("AP")
        startcolDisinfectionStationPlacements = XlsxInterfacer.ctn("AW")
        startcolStaticSpawnDataContainers = XlsxInterfacer.ctn("BD")

        self.stubEventsOnEnter = {}
        self.stubProgressionPuzzleToEnter = {}
        self.stubEnemySpawningInZone = {}
        self.stubTerminalPlacements = {}
        self.stubLocalLogFiles = {}
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
            iExpeditionZoneDataLists.readIntoDict(int, startcolEventsOnEnter+7, row, Snippet["Sound"], "SoundEvent")
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
        # LocalLogFiles
        # TODO
        while not(iExpeditionZoneDataLists.isEmpty(startcolLocalLogFiles,row)):
            Snippet = {}
            iExpeditionZoneDataLists.readIntoDict(str, startcolLocalLogFiles+1, row, Snippet, "FileName")
            iExpeditionZoneDataLists.readIntoDict(str, startcolLocalLogFiles+2, row, Snippet, "FileContent")
            try:Snippet["FileContent"] = re.sub(devnewlnregex,"\r\n", Snippet["FileContent"])
            except KeyError:pass
            iExpeditionZoneDataLists.readIntoDict(int, startcolLocalLogFiles+3, row, Snippet, "AttachedAudioFile")
            # TODO convert sound placeholders
            iExpeditionZoneDataLists.readIntoDict(int, startcolLocalLogFiles+4, row, Snippet, "AttachedAudioByteSize")
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
            iExpeditionZoneDataLists.readIntoDict(str, startcolTerminalPlacements+7, row, Snippet, "StartingState")
            EnumConverter.enumInDict(ENUMFILE_TERM_State, Snippet, "StartingState")
            iExpeditionZoneDataLists.readIntoDict(int, startcolTerminalPlacements+8, row, Snippet, "AudioEventEnter")
            iExpeditionZoneDataLists.readIntoDict(int, startcolTerminalPlacements+9, row, Snippet, "AudioEventExit")
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


    def EventsOnEnter(self, identifier):
        """returns the EventsOnEnter array for a specific zone"""
        try:return self.stubEventsOnEnter[identifier]
        except KeyError:pass
        return []

    def ProgressionPuzzleToEnterZonePlacementData(self, identifier):
        """returns the ZonePlacementData for the ProgressionPuzzleToEnter for a specific zone"""
        try:return self.stubProgressionPuzzleToEnter[identifier]
        except KeyError:pass
        return []

    def EnemySpawningInZone(self, identifier):
        """returns the EnemySpawningInZone array for a specific zone"""
        try:return self.stubEnemySpawningInZone[identifier]
        except KeyError:pass
        return []

    def TerminalPlacements(self, identifier):
        """returns the TerminalPlacements array for a specific zone"""
        try:return self.stubTerminalPlacements[identifier]
        except KeyError:pass
        return []

    def LocalLogFiles(self, group):
        """returns the LocalLogFiles array for a specific grouping to be used in the TerminalPlacements"""
        try:return self.stubLocalLogFiles[group]
        except KeyError:pass
        return []

    def PowerGeneratorPlacements(self, identifier):
        """returns the ZonePlacementWeights for the PowerGeneratorPlacements for a specific zone"""
        try:return self.stubPowerGeneratorPlacements[identifier]
        except KeyError:pass
        return []
    
    def DisinfectionStationPlacements(self, identifier):
        """returns the ZonePlacementWeights for the DisinfectionStationPlacements for a specific zone"""
        try:return self.stubDisinfectionStationPlacements[identifier]
        except KeyError:pass
        return []

    def StaticSpawnDataContainers(self, identifier):
        """returns the StaticSpawnDataContainers array for a specific zone"""
        try:return self.stubStaticSpawnDataContainers[identifier]
        except KeyError:pass
        return []

def WardenObjectiveBlock(iWardenObjective, iWardenObjectiveReactorWaves):
    """returns the Warden Objective"""
    # set up some checkpoints so if some of the data gets reformatted, not the entire function needs to be altered,
    # just the headings and contents of the section will need edited column values
    rowWavesOnElevatorLand = 22-1
    rowChainedPuzzleToActive = 46-1
    rowLightsOnFromBeginning = 60-1
    rowActivateHSU_ItemFromStart = 81-1

    data = {}

    data["Type"] = iWardenObjective.read(str, 1, 1)
    EnumConverter.enumInDict(ENUMFILE_eWardenObjectiveType, data, "Type")
    data["Header"] = iWardenObjective.read(str, 1, 3)
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
    iWardenObjective.readIntoDict(int, 1, 19, data, "ShowHelpDelay")

    data["WavesOnElevatorLand"] = GenericEnemyWaveDataList(iWardenObjective, 2, rowWavesOnElevatorLand+1, horizontal=True)
    iWardenObjective.readIntoDict(str, 1, rowWavesOnElevatorLand+6, data, "WaveOnElevatorWardenIntel")
    data["WavesOnActivate"] = GenericEnemyWaveDataList(iWardenObjective, 2, rowWavesOnElevatorLand+9, horizontal=True)
    iWardenObjective.readIntoDict(bool, 1, rowWavesOnElevatorLand+14, data, "StopAllWavesBeforeGotoWin")
    data["WavesOnGotoWin"] = GenericEnemyWaveDataList(iWardenObjective, 2, rowWavesOnElevatorLand+17, horizontal=True)
    iWardenObjective.readIntoDict(str, 1, rowWavesOnElevatorLand+22, data, "WaveOnGotoWinTrigger")
    EnumConverter.enumInDict(ENUMFILE_eWardenObjectiveWinCondition, data, "WaveOnGotoWinTrigger")

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
        data["Retrieve_Items"].append(DatablockIO.nameToId(DATABLOCK_Item, iWardenObjective(str, col, row)))
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
        data["PostCommandOutput"].append(iWardenObjective(str, col, row))
        col+= 1
    iWardenObjective.readIntoDict(int, 1, rowLightsOnFromBeginning+8, data, "PowerCellsToDistribute")
    iWardenObjective.readIntoDict(int, 1, rowLightsOnFromBeginning+10, data, "Uplink_NumberOfVerificationRounds")
    iWardenObjective.readIntoDict(int, 1, rowLightsOnFromBeginning+11, data, "Uplink_NumberOfTerminals")
    iWardenObjective.readIntoDict(int, 1, rowLightsOnFromBeginning+13, data, "CentralPowerGenClustser_NumberOfGenerators")
    iWardenObjective.readIntoDict(int, 1, rowLightsOnFromBeginning+14, data, "CentralPowerGenClustser_NumberOfPowerCells")

    iWardenObjective.readIntoDict(str, 1, rowActivateHSU_ItemFromStart, data, "ActivateHSU_ItemFromStart")
    DatablockIO.nameInDict(DATABLOCK_Item, data, "ActivateHSU_ItemFromStart")
    iWardenObjective.readIntoDict(str, 1, rowActivateHSU_ItemFromStart+1, data, "ActivateHSU_ItemAfterActivation")
    DatablockIO.nameInDict(DATABLOCK_Item, data, "ActivateHSU_ItemAfterActivation")
    iWardenObjective.readIntoDict(bool, 1, rowActivateHSU_ItemFromStart+2, data, "ActivateHSU_StopEnemyWavesOnActivation")
    iWardenObjective.readIntoDict(bool, 1, rowActivateHSU_ItemFromStart+3, data, "ActivateHSU_ObjectiveCompleteAfterInsertion")
    data["ActivateHSU_Events"] = []
    col,row = 2,rowActivateHSU_ItemFromStart+6
    while not(iWardenObjective.isEmpty(col,row)):
        data["ActivateHSU_Events"] = WardenObjectiveEventData(iWardenObjective, col, row, horizontal=False)
        col+= 1
    
    # Set default values
    data["name"] = "DPK Utility Objective"
    data["internalEnabled"] = False
    data["persistentID"] = -1
    # Attempt to fill default values with those from the table
    iWardenObjective.readIntoDict(str,1, rowActivateHSU_ItemFromStart+13, data, "name")
    iWardenObjective.readIntoDict(bool,1, rowActivateHSU_ItemFromStart+14, data, "internalEnabled")
    iWardenObjective.readIntoDict(int,1, rowActivateHSU_ItemFromStart+15, data, "persistentID")
    return data

class ReactorWaveData:
    """
    a class that contains a dictionaries for the ReactorWaveData (since the sheet cannot contain 2d-3d data) \n
    access wave data by referencing ReactorWaveData.waves
    """

    def __init__(self, iWardenObjectiveReactorWaves):
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
            iWardenObjectiveReactorWaves.readIntoDict(str, startcolEnemyWaves, row, Snippet, "WavePopulation")
            DatablockIO.nameInDict(DATABLOCK_SurvivalWavePopulation, Snippet, "WavePopulation")
            iWardenObjectiveReactorWaves.readIntoDict(float, startcolEnemyWaves, row, Snippet, "SpawnTimeRel")
            iWardenObjectiveReactorWaves.readIntoDict(str, startcolEnemyWaves, row, Snippet, "SpawnType")
            EnumConverter.enumInDict(ENUMFILE_eReactorWaveSpawnType, Snippet, "SpawnType")
            EnsureKeyInDictArray(Snippet, waveNo)
            self.stubEnemyWaves[waveNo].append(Snippet)
            row+= 1

        # Events
        row = startrow
        while not(iWardenObjectiveReactorWaves.isEmpty(startcolEvents-1, row)):
            Snippet = {}
            waveNo = iWardenObjectiveReactorWaves.read(str, startcolEvents-1, row)
            Snippet["Events"] = WardenObjectiveEventData(iWardenObjectiveReactorWaves, startcolEvents, row, horizontal=True)
            EnsureKeyInDictArray(Snippet, waveNo)
            self.stubEvents[waveNo].append(Snippet)
            row+= 1

        # ReactorWaves
        row = startrow
        while not(iWardenObjectiveReactorWaves.isEmpty(startcolReactorWaves, row)):
            wave = {}
            waveNo = iWardenObjectiveReactorWaves.read(str, startcolReactorWaves, row)
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

    def EnemyWaves(self, identifier):
        """returns the EnemyWaves array for a specific zone"""
        try:return self.stubEnemyWaves[identifier]
        except KeyError:pass
        return []

    def Events(self, identifier):
        """returns the EnemyWaves array for a specific zone"""
        try:return self.stubEvents[identifier]
        except KeyError:pass
        return []

def finalizeData(dictExpeditionInTier, arrdictLevelLayoutBlock, arrdictWardenObjectiveBlock):
    """
    finalizeData takes the ExpeditionInTier, LevelLayoutBlocks, and WardenObjectiveBlocks \n
    for the LevelLayoutBlocks, it fills in the default block metadata \n
    for the WardenObjectiveBlocks, it overrides the block metadata
    """
    utilityname = dictExpeditionInTier["Descriptive"]["PublicName"] + " - " + dictExpeditionInTier["Descriptive"]["Prefix"]
    try:
        arrdictLevelLayoutBlock[0]["name"] = utilityname + "L1 Layout"
        arrdictLevelLayoutBlock[0]["internalEnabled"] = dictExpeditionInTier["Enabled"]
        arrdictLevelLayoutBlock[0]["persistentID"] = dictExpeditionInTier["LevelLayoutData"]
        arrdictLevelLayoutBlock[1]["name"] = utilityname + "L2 Layout"
        arrdictLevelLayoutBlock[1]["internalEnabled"] = dictExpeditionInTier["SecondaryLayerEnabled"]
        arrdictLevelLayoutBlock[1]["persistentID"] = dictExpeditionInTier["SecondaryLayout"]
        arrdictLevelLayoutBlock[2]["name"] = utilityname + "L3 Layout"
        arrdictLevelLayoutBlock[2]["internalEnabled"] = dictExpeditionInTier["ThirdLayerEnabled"]
        arrdictLevelLayoutBlock[2]["persistentID"] = dictExpeditionInTier["ThirdLayout"]
    except IndexError:pass
    except KeyError:pass
    try:
        arrdictWardenObjectiveBlock[0]["name"] = utilityname + "L1 Objective"
        arrdictWardenObjectiveBlock[0]["internalEnabled"] = dictExpeditionInTier["Enabled"]
        arrdictWardenObjectiveBlock[0]["persistentID"] = dictExpeditionInTier["MainLayerData"]["ObjectiveData"]["DataBlockId"]
        arrdictWardenObjectiveBlock[1]["name"] = utilityname + "L2 Objective"
        arrdictWardenObjectiveBlock[1]["internalEnabled"] = dictExpeditionInTier["SecondaryLayerEnabled"]
        arrdictWardenObjectiveBlock[1]["persistentID"] = dictExpeditionInTier["SecondaryLayerData"]["ObjectiveData"]["DataBlockId"]
        arrdictWardenObjectiveBlock[2]["name"] = utilityname + "L3 Objective"
        arrdictWardenObjectiveBlock[2]["internalEnabled"] = dictExpeditionInTier["ThirdLayerEnabled"]
        arrdictWardenObjectiveBlock[2]["persistentID"] = dictExpeditionInTier["ThirdLayerData"]["ObjectiveData"]["DataBlockId"]
    except IndexError:pass
    except KeyError:pass

if __name__ == "__main__":
    filename = "testgenerator.xlsx"
    # get all interfaces
    iKey = XlsxInterfacer.interfacer(pandas.read_excel(filename, "Key", header=None))
    if iKey.read(str, 0, 5)[:len(Version)] != Version:
        raise Exception("Incompatible utility and sheet version.")

    iExpeditionInTier = XlsxInterfacer.interfacer(pandas.read_excel(filename, "ExpeditionInTier", header=None))
    try:
        iL1ExpeditionZoneData = XlsxInterfacer.interfacer(pandas.read_excel(filename, "L1 ExpeditionZoneData", header=None))
        iL1ExpeditionZoneDataLists = XlsxInterfacer.interfacer(pandas.read_excel(filename, "L1 ExpeditionZoneData Lists", header=None))
        iL2ExpeditionZoneData = XlsxInterfacer.interfacer(pandas.read_excel(filename, "L2 ExpeditionZoneData", header=None))
        iL2ExpeditionZoneDataLists = XlsxInterfacer.interfacer(pandas.read_excel(filename, "L2 ExpeditionZoneData Lists", header=None))
        iL3ExpeditionZoneData = XlsxInterfacer.interfacer(pandas.read_excel(filename, "L3 ExpeditionZoneData", header=None))
        iL3ExpeditionZoneDataLists = XlsxInterfacer.interfacer(pandas.read_excel(filename, "L3 ExpeditionZoneData Lists", header=None))
    except xlrd.biffh.XLRDError:
        pass
    try: 
        iL1WardenObjective = XlsxInterfacer.interfacer(pandas.read_excel(filename, "L1 WardenObjective", header=None))
        iL1WardenObjectiveReactorWaves = XlsxInterfacer.interfacer(pandas.read_excel(filename, "L1 WardenObjective ReactorWaves", header=None))
        iL2WardenObjective = XlsxInterfacer.interfacer(pandas.read_excel(filename, "L2 WardenObjective", header=None))
        iL2WardenObjectiveReactorWaves = XlsxInterfacer.interfacer(pandas.read_excel(filename, "L2 WardenObjective ReactorWaves", header=None))
        iL3WardenObjective = XlsxInterfacer.interfacer(pandas.read_excel(filename, "L3 WardenObjective", header=None))
        iL3WardenObjectiveReactorWaves = XlsxInterfacer.interfacer(pandas.read_excel(filename, "L3 WardenObjective ReactorWaves", header=None))
    except xlrd.biffh.XLRDError:
        pass

    dictExpeditionInTier = ExpeditionInTier(iExpeditionInTier)
    arrdictLevelLayoutBlock = []
    try:
        arrdictLevelLayoutBlock.append(LevelLayoutBlock(iL1ExpeditionZoneData, iL1ExpeditionZoneDataLists))
        arrdictLevelLayoutBlock.append(LevelLayoutBlock(iL2ExpeditionZoneData, iL2ExpeditionZoneDataLists))
        arrdictLevelLayoutBlock.append(LevelLayoutBlock(iL3ExpeditionZoneData, iL3ExpeditionZoneDataLists))
    except NameError:pass
    arrdictWardenObjectiveBlock = []
    try:
        arrdictWardenObjectiveBlock.append(WardenObjectiveBlock(iL1WardenObjective, iL1WardenObjectiveReactorWaves))
        arrdictWardenObjectiveBlock.append(WardenObjectiveBlock(iL2WardenObjective, iL2WardenObjectiveReactorWaves))
        arrdictWardenObjectiveBlock.append(WardenObjectiveBlock(iL3WardenObjective, iL3WardenObjectiveReactorWaves))
    except NameError:pass

    finalizeData(dictExpeditionInTier, arrdictLevelLayoutBlock, arrdictWardenObjectiveBlock)

    # test output
    json.dump(dictExpeditionInTier,open("RundownPiece.json", 'w', encoding="utf8"),ensure_ascii=False,allow_nan=False,indent=2)
    try:
        json.dump(arrdictLevelLayoutBlock[0],open("L1LevelLayoutBlock.json", 'w', encoding="utf8"),ensure_ascii=False,allow_nan=False,indent=2)
        json.dump(arrdictLevelLayoutBlock[1],open("L2LevelLayoutBlock.json", 'w', encoding="utf8"),ensure_ascii=False,allow_nan=False,indent=2)
        json.dump(arrdictLevelLayoutBlock[2],open("L3LevelLayoutBlock.json", 'w', encoding="utf8"),ensure_ascii=False,allow_nan=False,indent=2)
    except IndexError:pass
    try:
        json.dump(arrdictWardenObjectiveBlock[0],open("L1WardenObjective.json", 'w', encoding="utf8"),ensure_ascii=False,allow_nan=False,indent=2)
        json.dump(arrdictWardenObjectiveBlock[1],open("L2WardenObjective.json", 'w', encoding="utf8"),ensure_ascii=False,allow_nan=False,indent=2)
        json.dump(arrdictWardenObjectiveBlock[2],open("L3WardenObjective.json", 'w', encoding="utf8"),ensure_ascii=False,allow_nan=False,indent=2)
    except IndexError:pass

    input("Done.")