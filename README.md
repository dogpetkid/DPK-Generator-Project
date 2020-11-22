# DPK Generator Project

<!-- Author: Dogpetkid -->
<!-- Date: 2020-11-06 -->
<!-- Revised: 2020-11-22 -->
<style>
  body {
    background-color: #1e1e1e
  }
  *:not(a) {
    color: #d4d4d4
  }
</style>

This is a project to generate levels for GTFO. For an explanation on how to use this project, see [How to use](#how-to-use).

[Project releases.][DPK Releases]

## Project dependencies

- [numpy][]
- [xlrd][]
- [pandas][]

## Project contents

- [XlsxInterfacer.py](#xlsxinterfacer)
- [DatablockIO.py](#datablockio)
- [EnumConverter.py](#enumconverter)
- [LevelUtility.py](#levelutility)
  - [Template](#template)

---

## XlsxInterfacer

A module created to interface with the pandas.DataFrame and read specific typed values.

## DatablockIO

A module created to read from and write to GTFO Datablocks including search blocks based on name or persistantID.

## EnumConverter

A module created to convert between GTFO Enumerator names and indexes.

---

## LevelUtility

This is a file that is able to take a templated xlsx file and convert it into GTFO ExpeditionInTier, LevelLayout, and WardenObjective data. It takes arguments of which files to convert, converts the xlsx into dictionaries, and then writes those into the Datablocks.

### Template

[Template is here.][DPK Template]

Notes:

- The sheet tries it's best to stop you from doing illegal or irrational things for data types (ie you can't put a word where a number should go) but it is not idiot proof and cannot detect logical issues (eg to go into a zone with a key, it requires that key or a level doesn't have enough cells).
- Lots of variables have notes on them, look at the notes for more specific information about variables.
- Most values can be left blank, if leaving a value blank could be problematic, the program will probably describe a cell as empty.
- Some values are already filled out in the template, these are just typical values and can be changed.

Below is a description of the sheets in the template:

- Key: A sheet describing datatypes and the version of the sheet.
- Meta: Data for which rundown, tier, and index the level should be in.
- ExpeditionInTier: Data about the expedition as a whole (eg population, fog, tileset) and general information about the sectors.

All sheets labed LX should be duplicated per layer and renamed accordingly (eg L1, L2, L3).

- LX ExpeditionZoneData: List of all zones and information about them.
- LX ExpeditionZoneData Lists: Different lists of items a zone can have more than one of (eg terminals, enemy spawning, etc.)

- LX WardenObjective: Data about the objective, the enemy waves, objective events, etc.
- LX WardenObjective ReactorWaves: Data about the Reactor objective

Note: The program requires the Zone and Objective sheets to come as a pair (even if the second is blank). If you have L1 ExpeditionZoneData, you must have L1 ExpeditionZoneData. If you have L1 WardenObjective, you must have L1 WardenObjective.

---

## How to use

The usage of this program can be broken into a couple of steps:

- [Setup](#setup)
  - File Archiver (7-Zip)
  - Modding folder
  - Python and dependencies
  - Storm's Hack (for testing blocks easily)
  - (Unity) Asset Bundle Extractor (UABE) (for releasing blocks) & Binary Encoder
- [Template Usage](#template-usage)
  - Level Design (crash course)
  - Meta (placement in rundown)
  - General information (ExpeditionInTier)
  - Zones (ExpeditionZoneData & Lists)
  - Objectives (WardenObjective & ReactorWaves)
- [Tool Usage](#tool-usage)

### Setup

#### 7-Zip

The first item needed for most of the modding is being able to extract archives. I suggest [7-Zip][]. Download the x64 (unless your computer x32).

![7-Zip image][]

After running the Installer, a context menu should be present when right clicking on .zip files (and other archives).

#### Modding Folder

Create a `GTFO-modding` folder (it's name does not particularly matter, for clarity I will refer to it by that name). In it place this project's folder inside of the `GTFO-modding` folder.

##### Original Datablocks

Then download the [Original Datablocks][], move it to `GTFO-modding`, and extract the blocks (by right clicking and doing 7-Zip>Extract Here). Make sure to rename the blocks folder to `Datablocks` (this is for the Level Utility to recognize it). After doing this, your `GTFO-modding` folder should look like this:

![GTFO-modding folder image][]

#### [Python][]

Download Python [here][python downloads] and install it (all the default settings). Open a terminal (`Windows+R` and open CMD) and run:

`pip install pandas xlrd numpy`

This will download all the needed dependencies for the DPK Generator Project.

#### GTFO Hack

Download this [R4 Build][] and extract it where your steam games are (normally here: `C:\Program Files (x86)\Steam\steamapps\common`). Download the [GTFO hack][] and extract it inside of the `common\GTFO r4`. To swap rundowns, switch which one is named `GTFO` and rename the old one.

#### [(Unity) Asset Bundle Extractor][uabe]

Download UABE [here][uabe releases] and extract it somewhere.

##### Binary Encoder

Save the [Binary Encoder][]. (The best place to put is probably is in your `GTFO-modding` folder. It will be used with UABE and is needed because the datablocks are serialized.)

---

### Template Usage

See [Template](#template) to see notes about the template.

#### Level Design (crash course)

This section **will not** teach you how to make good levels but what it will do is walk you through the very general steps for creating the idea for the level and some tips.

##### Create a theme

A level that is unique is interesting and a level that follows a single idea seems more coherent (versus being all over). A good way to do this is create something the level will focus on or a rule for the level (e.g. only use specific enemies, all alarms are the same kind, sectors unlock zones in other sectors).

##### Floorplan

The next step is then creating a model of the level to help visualize the way the level works. This can be done whichever way, I prefer drawing the level in MS Paint or coloring different squares in Google Sheets. A thing to keep in mind while placing zones is which secuirty doors and [progression puzzles and chained puzzles](#glossary) they take to open and make notes of where the key or cell will come from.

##### Tips

- Backtracking is not fun, try to keep the amount of walking through empty rooms to a minimum.
- A cell in one sector can open a door in another sector.
- Try to walk through the way the level should be completed and double check if "cheesing" by skipping required zones and accidental softlocks aren't a thing.

---

#### Meta

This describes where in the rundown the level is placed. This is for the tool to know where in the `RundownDataBlock.json` to put the ExpeditionInTier. E.g. Rundown ID 25, TierA, Index 1 describes (R4)A2 Foster

- etc., describe in cell notes

---

#### General information (ExpeditionInTier)

A level will only have 1 ExpeditionInTier sheet to be inserted into the `RundownDataBlock.json`. This describe most general information for a particular level (including some details about the layers that are not in the other sheets.)

- `Enabled` & `Accessiblity`: see notes
- `LevelLayoutData`, `SecondaryLayout`, & `ThirdLayout`: An id number (1-101 are already taken by the game)
- `SecondaryLayerEnabled` & `ThirdLayerEnabled`: `True` if that Layer is in the level

##### Descriptive

Name and stuff about the level.

- `Prefix`: e.g. R3A
- `PublicName`: e.g. Bolt
- `ExpeditionDepth`: e.g. 579
- `EstimatedDuration`: A guess at the length of the level
- `ExpeditionDescription`: Flavortext (supports xml)
- `RoleplayedWardenIntel`: Flavortext (supports xml)
- `DevInfo`: A brief phrase describing the level

##### Seeds

Seeds the level, self explanatory.

##### SpecialOverrideData

Unsure of what exactly this does. All levels set `WeakResourceContainerWithPackChanceForLocked` to -1.

##### Expedition

Picks settings for the entire level.

- `ComplexResourceData`: The level tileset. (Only `Complex_Mining` and `Complex_Tech` have bulkheads.)
- `LightSettings`: Lights
- `FogSettings`: Fog level (the available setting names don't help much `¯\_(ツ)_/¯`)
- `EnemyPopulation`: Determines the pool of enemies the groups are selected from.
- `ExpeditionBalance`: a setting with a bunch of subsettings (only `Normal` works)
- `ScoutWaveSettings`: [Wave Setting](#glossary), defaultly `Scout`
- `ScoutWavePopulation`: [Wave Population](#glossary), defaulty `Baseline`

##### LayerData

Only fill out for layers you have enabled.

This section applies for `MainLayerData`, `SecondaryLayerData`, and `ThirdLayerData`.

- `ZoneAliasStart`: Zone 0's Alias
- `ZonesWithBulkheadEntrance`: List of zones containing bulkheads
- *`BulkheadDoorControllerPlacements`*: [Placements](#glossary) for bulkhead controllers
- *`BulkheadKeyPlacements`*: [Placements](#glossary) for bulkhead keys
- *`ObjectiveData`*: General about the objective
  - `DataBlockId`: an id number (1-58 are already taken by the game)
  - `WinCondition`: determines whether the players extract from spawn or a forward extraction
  - *`ZonePlacementDatas`*: [Placements](#glossary) for objective items

---

#### Zones (ExpeditionZoneData & Lists)

For every layer, a pair of `LX ExpeditionZoneData` and `LX ExpeditionZoneData Lists` sheets are needed. The `X` should be changed to whichever layer number it is (e.g. `L1 ExpeditionZoneData` for high). Every layer needs the sheets to come in pairs (even if the lists are empty). This will be inserted into `LevelLayoutDataBlock.json`. (Side note: the reason for the Lists sheet is due to having to adjust 3D and 4D data to be 2D to work with a spreadsheet.)

General Zone Data:

- `LocalIndex`: Zone index
- `SubSeed`: Seed
- `BulkheadDCScanSeed`: Seed
- `SubComplex`: Narrows selected tiles
- `CustomGeomorph`: a tile that's garunteed to be placed a.k.a. setpiece, see [Reference](#reference)
- `x`: zone length (North South)
- `y`: zone width (East West)
- `BuildFromLocalIndex`: Previous zone
- `StartPosition`: Where in the previous zone the security door for said zone should be placed.
- `StartPosition_IndexWeight`: 0 to 1 (0 to 100%) to determine where in zone for `StartPosition` set to `From_IndexWeight`
- `StartExpansion`: direction for the security door entering the zone
- `ZoneExpansion`: determines which direction to continue building the zone
- `LightSettings`: Lights
- `AllowedZoneAltitude`: Height of zone
- `ChanceToChange`: 0 to 1 (0 to 100%) chance for zone to stray from set altitude (*not sure about this one)

Door information:

- `PuzzleType`: [Progression puzzle](#glossary)
- `CustomText`: Flavortext on door using `Locked_no_key` (supports xml) (default `<color=red>://ERROR: Door in emergency lockdown, unable to operate.</color>`)
- `PlacementCount`: number of solutions to place (default 1)
- `ChainedPuzzleToEnter`: [Chained puzzle](#glossary) (defaultly unlocked) (Common puzzles)
- `SecurityGateToEnter`: Security or Apex door
- `HasActiveEnemyWave`: Blood door?
- `EnemyGroupInfrontOfDoor`: Enemy group that spawns directly in front of door (typically `Hunter_easy_infront`)
- `EnemyGroupInArea`: Enemy group(s) that spawn in the back of the area (typically `Hunter_easy_inback`)
- `EnemyGroupsInArea`: number of "`EnemyGroupInArea`"s (defaultly 1)

Decorations (and functional):

- `HSUClustersInZone` & `CorpseClustersInZone`  `ResourceContainerClustersInZone` & `GeneratorClustersInZone` & `CorpsesInZone` & `HSUsInZone` & `DeconUnitsInZone`: Amount of X in zone

Item spawns (and terminals):

- `AllowSmallPickupsAllocation`: True if consumables can spawn
- `AllowResourceContainerAllocation`: True if lockers/boxes can spawn
- `ForceBigPickupsAllocation`: True if carry items / big pickups can spawn
- `ConsumableDistributionInZone`: Consumable setting
- `BigPickupDistributionInZone`: Carry item / big pickup setting
- `ForbidTerminalsInZone`: True if no terminals can spawn
- `HealthMulti` & `WeaponAmmoMulti` & `ToolAmmoMulti` & `DisinfectionMulti`: Resource amounts (as a %, .2 is 1 use)
- `HealthPlacement` & `WeaponAmmoPlacement` & `ToolAmmoPlacement` & `DisinfectionPlacement`: [Placement](#glossary) weights

##### Lists

- *`EventsOnEnter`*: Events that occur when a security door opens
  - `DPK Identifier`: Describes which zone's security door has an event
  - `Delay`: Number of seconds after the door is opened that the event triggers
  - `Enabled`: (*not sure about this one)
  - `RadiusMin`: (*not sure about this one)
  - `RadiusMax`: (*not sure about this one)
  - `Enabled`: True if the door has an intel message
  - `IntelMessage`: Flavortext (supports xml)
  - `Enabled`: True if the door has a sound event
  - `SoundEvent`: Sound event id, see [Reference](#reference)
- *`ProgressionPuzzleToEnter`*: Describes all [Progression puzzles](#glossary) solutions
  - `DPK Identifier`: Describes which zone's security door requires a puzzle
  - PuzzleType: not functional, shows puzzle type for zone, simply for ease of use [Progression puzzle](#glossary)
  - *`ZonePlacementData`*: [Placement](#glossary) of puzzle solution
- *`EnemySpawningInZone`*: Describes enemy spawns
  - `DPK Identifier`: Describes which zone the spawn data is for
  - DPK Reminder: A place to write a reminder what enemy is attempting to be selected
  - `GroupType`: Enemy group type, see [Reference](#reference)
  - `Difficulty`: Enemy group difficulty, see [Reference](#reference)
  - `Distribution`: Determines distribution type, `None` places none, `Force_One` places 1 group, `Rel_Value` uses the distribution value to determine group size
  - `DistributionValue`: group population = `DistributionValue` \* zone population (zone population is 25 on `Normal` for `ExpeditionBalance`) (as a %)
- *`TerminalPlacements`*: Defines all terminal spawns
  - `DPK Identifier`: Describes which zone the terminal is in
  - *`PlacementWeights`*: [Placement](#glossary) weights
  - `AreaSeedOffset`: Seed
  - `MarkerSeedOffset`: Seed
  - `DPK Group`: [Group](#glossary) of logs to pull from and place onto the terminal
  - `StartingState`: The state the computer spawns in (defaultly sleeping)
  - `AudioEventEnter`: Audio event the terminal begins in, see [Reference](#reference)
  - `AudioEventExit`: Audio event that occurs when the above audio event is exited, see [Reference](#reference)
- *`LocalLogFiles`*: A list of log files that can be put onto terminals
  - `DPK Group`: [Group](#glossary) of logs that can put on terminals
  - `FileName`: The name of the file
  - `FileContent`: The contents of the log file (log files contents in order to be readable)
  - `AttachedAudioFile`: The id of the audio log to be attached
  - `AttachedAudioByteSize`: The display size of the audio log
- *`PowerGeneratorPlacements`*: [Placement](#glossary)
- *`DisinfectionStationPlacements`*: [Placement](#glossary)
- *`StaticSpawnDataContainers`*: Defines spawns for static enemies (likes spitters and sacks)
  - `DPK Identifier`: Describes the zone for the [Spawn Container](#glossary)
  - `Count`: Number of containers
  - `DistributionWeightType`: Describes whether the containers should be spread across the zone or localized near one weight (*not sure about this one)
  - `DistributionWeight`: 0 to 1 (0 to 100%) to determine where in zone to localize the containers when `DistributionWeightType` is `Weight_is_exact_node_index` (*not sure about this one)
  - `DistributionRandomBlend`: (*not sure about this one)
  - `DistributionResultPow`: (*not sure about this one)
  - `StaticSpawnDataId`: Which static enemy to assign to the container
  - `FixedSeed`: Seed

---

#### Objectives (WardenObjective & ReactorWaves)

For every layer, a pair of `LX WardenObjective` and `LX ExpeditionZoneData` sheets are needed. The `X` should be changed to whichever layer number it is (e.g. `L1 WardenObjective` for high). Every layer needs the sheets to come in pairs (even if the Waves are empty). This will be inserted into `WardenObjectiveDataBlock.json`.

For objectives, fill out only information pertaining to the objective type you are using.

Generic Objective Data:

- `Type`: The type of objective, see [Glossary](#glossary)
- `Header`: Display name of objective
- `MainObjective`: Objective text for level (displayed the whole time), see [Reference](#reference)
- `FindLocationInfo`: Help to find where objective is located, see [Reference](#reference)
- `FindLocationInfoHelp`: Alternate message displayed after help delay
- `GoToZone`: Help to go to the objective, see [Reference](#reference)
- `GoToZoneHelp`: Alternate message displayed after help delay
- `InZoneFindItem`: Help to find the objective in its zone, see [Reference](#Reference)
- `InZoneFindItemHelp`: Alternate message displayed after help delay
- `SolveItem`: Help to do the objective after the item is found, see [Reference](#reference)
- `SolveItemHelp`: Alternate message displayed after help delay
- `GoToWinCondition_Elevator`: Help to do the objective after the item is found, see [Reference](#reference)
- `GoToWinConditionHelp_Elevator`: Alternate message displayed after help delay
- `GoToWinCondition_CustomGeo`: Objective that displays after objective is completed and the team needs to leave at a forward extract, see [Reference](#reference)
- `GoToWinConditionHelp_CustomGeo`: Alternate message displayed after help delay
- `GoToWinCondition_ToMainLayer`: Objective that displays after objective is completed and the objective is optional, see [Reference](#reference)
- `GoToWinConditionHelp_ToMainLayer`: Alternate message displayed after help delay
- `ShowHelpDelay`: Duration of time in seconds before help variant of objective displays (default 180)
- *`WavesOnElevatorLand`*: [Generic Enemy Wave](#glossary)
- *`WavesOnActivate`*: [Generic Enemy Wave](#glossary)
- *`WavesOnGotoWin`*: [Generic Enemy Wave](#glossary)
- *`EventsOnGotoWin`*: [Warden Objective Event](#glossary)
- `ChainedPuzzleToActive`: Chained puzzle at the start of the objective (defaultly none, typically `HSU_Scan`)
- `ChainedPuzzleMidObjective`: Second chained puzzle for objective (defaultly none)
- `ChainedPuzzleAtExit`: Chained puzzle spawned at extract (defaultly none, typically `ExpeditionExit_Scan`)
- `ChainedPuzzleAtExitScanSpeedMultiplier`: speed multiplier for extract (as a %)

Gather (`GatherSmallItems`):

- `Gather_RequiredCount`: Number of Gather items required
- `Gather_ItemId`: Item to gether (can be any item? not sure)
- `Gather_SpawnCount`: Amount of Gather items that spawn
- `Gather_MaxPerZone`: Maximum Gather items that can spawn per zone

Retrieve (`RetrieveBigItems`):

- `Retrieve_Items`: List of Retrieve/Carry items needed for extract (can be any item? not sure)

Reactor (`ReactorStartup` & `Reactor_Shutdown`):

- *`ReactorWaves`*: see [Reactor Waves](#reactor-waves)
- `LightsOnFromBeginning`: True if level lights are on before the reactor
- `LightsOnDuringIntro`: True if level lights are on duing the reactor startup sequences
- `LightsOnWhenStartupComplete`: True if level lights are on after the reactor

Command (`SpecialTerminalCommand`):

- `SpecialTerminalCommand`: Name of the special terminal command
- `SpecialTerminalCommandDesc`: Flavortext displayed when using the COMMANDS
- `PostCommandOutput`: List of outputs displayed after the command is completed

Distribute (`PowerCellDistribution`):

- `PowerCellsToDistribute`: Number of cells to distribute to generators

Uplink (`TerminalUplink`):

- `Uplink_NumberOfVerificationRounds`: Number of uplinks per terminal
- `Uplink_NumberOfTerminals`: Number of uplink terminals

Cluster (`CentralGeneratorCluster`):

- `CentralPowerGenClustser_NumberOfGenerators`: Number of generators in the cluster
- `CentralPowerGenClustser_NumberOfPowerCells`: Number of cells spawned by the objective

Activate (`ActivateSmallHSU`):

- `ActivateHSU_ItemFromStart`: HSU item at start (can be any item? not sure)
- `ActivateHSU_ItemAfterActivation`: HSU item that replaces the first after activation (can be any item? not sure)
- `ActivateHSU_StopEnemyWavesOnActivation`: Stops waves when HSU item is inserted
- `ActivateHSU_ObjectiveCompleteAfterInsertion`: True if inserting the HSU completes the objective, false if the objective is completed when the HSU is brought to extract
- *`ActivateHSU_Events`*: [Warden Objective Event](#glossary)

##### Reactor Waves

ReactorWaves:

- `Wave No.`: Wave numbers in order (1 and up)
- `Warmup`: Time during warmup (seconds)
- `WarmupFail`: Time during warmup after failing (seconds)
- `Wave`: Time during wave (seconds)
- `Verify`: Time during verify (seconds)
- `VerifyFail`: Time during verify after failing (seocnds)
- `VerifyInOtherZone`: True if code is from a terminal, false if code is given
- `ZoneForVerification`: Zone the code is located in (several codes can be located in the same zone)

EnemyWaves:

- `Wave No.`: Reactor Wave number the Enemy Wave is for, see [Identifier](#glossary)
- `WaveSettings`: [Wave Setting](#glossary)
- `WavePopulation`: [Wave Population](#glossary)
- `SpawnTimeRel`: 0 to 1 (0 to 100%) as the amount through the wave duration the enemy wave will spawn
- `SpawnType`: Determines where the enemies spawn

Events:

- `Wave No.`: Reactor Wave number the Enemy Wave is for, see [Identifier](#glossary)
- *`Events`*: [Warden Objective Event](#glossary)
  - `Trigger`: `None` = being of Warmup, `OnStart` = beginning of Wave, `OnMid` = beginning of Verify, `OnEnd` = after Verify
  - etc., see [Warden Objective Event](#Glossary)

---

### Tool Usage

The following are the steps to use the tool.

1. Follow the [Setup](#setup)
2. Design a level
3. Create a copy of the [Template][DPK Template] in either Google Sheets or download it as a .xlsx (to edit in Excel, Google Sheets is recommended)
4. Fill out the copy of the template.
5. Download the Google Sheets as an Excel file (.xlsx) if you didn't in step 3
6. Drag the .xlsx onto the LevelUtility (or run the tool in CLI). Note: the only Rundown, LevelLayout, and WardenObjective datablocks are edited by the tool

GTFO Hack (testing):

1. (To test with the GTFO Hack) copy the updated datablocks over to `common\GTFO\gtfohack\CustomDataBlocks` (or in a folder)
2. (To test with the GTFO Hack) after launching the game, hit \` (Grave) and type `LoadBlocks` (or if blocks are in a sub folder, `LoadBlocks <folder>`)

UABE:

1. Use the Binary Encoder to encode all files. (All modified blocks can be dragged onto it for it to encode them.)
2. In UABE, File > Open and open `common\GTFO\GTFO_Data\resources.assets`
3. Select U2019.2.0f1 (should be the last one)
4. Sort the blocks in alphabetical order by clicking Name
5. To find the blocks, do View > Search by name, and search "Gamedata.*" to find the blocks
6. For each modified datablock:
7. Select the block > Plugins > Import from .txt and select the bin encoded file
8. After all blocks are added, File > Save and save it as `resourcesModded.assets` (or whatever other name, it can not save to the same name). This .assets can then replace the original and be shared easier.

---

## Glossary

- Progression Puzzle: Key, Cell, or Locked status on doors
- Chained Puzzle: A list of scans during an alarm
- Wave Setting: Defines the size and frequency of waves
- Wave Population: Defines the exact enemies that can spawn in a wave
- Placement-type Data: has information about weighting and can have information about DPK Idenfier, Zone, and seeding
- DPK Identifier/Group: A value to differentiate items in "lists of lists." (Due to complications with storing 3D data as 2D data).
  - e.g. Bulkhead keys need to be placed in \[zones 0\] AND \[zone 1 or zone 2\], meaning zone 0 always has a key and the other is between zones 1 and 1. This would be represented in the sheet as "a" zone 0, "b" zone 1, "b" zone 2 so the tool knows the two "b" items are in the same sublist.
  - e.g. A Cluster requires 2 cells. These can spawn in will spawn in \[zones 12 or zone 14\] AND \[zone 9 or zone 13\], meaning one cell is in either 12 or 14 and the other is either 9 or 13. This would be represented in the sheet as "a" zone 12, "a" zone 14, "b" zone 9, "b" zone 13 so the tool knows the two "b" items are in the same sublist.
- Objective type: Non-exhaustive list of examples listed below:
  - HSU_FindTakeSample: R1A1, R3C1
  - Reactor_Startup: R2D2, R4D1L1
  - Reactor_Shutdown: R1D1
  - GatherSmallItems: R1B1, R2C2, R4C2L1
  - ClearAPath: R2B1, R4D2L1
  - SpecialTerminalCommand: R4A2L3, R4B2L2
  - RetrieveBigItems: R2B4, R2E1
  - PowerCellDistribution: R2B2, R4B1L1
  - TerminalUplink: R2C1, R3B1, R4C1L2
  - CentralGeneratorCluster: R2D1, R4D2L2
  - ActivateSmallHSU: R3A1, R3D1, R4C1L1
- Generic Enemy Wave: Describes an Enemy Wave (for Objectives), made of a Wave Setting, Wave Population, a spawn delay (seconds), a boolean to set the alarm noise on, and an intel message (flavortext)
- Warden Objective Event: Describes an objective event to unlock or open a door or show an intel message, made of a Trigger (None in most cases), a type of action to a door, the layer the event is located, the zone the event is located, a delay (seconds), and intel message (flavortext)

## Reference

A list of common assets, ids, and etc. which are not options provided in the list.

### Geomorphs

All tiles can be found in [`ComplexResourceSetDataBlock.json`][ComplexResourceSetDataBlock]. (Note: Only tiles within the selected ComplexResourceSet can be used. E.g. lab tiles can not be used in the Mining Complex.)

| Nickname | Tile Path |
| --- | --- |
| Mining exit | Assets/AssetPrefabs/Complex/Mining/Geomorphs/geo_64x64_mining_exit_01.prefab |
| Tech exit | Assets/AssetPrefabs/Complex/Tech/Geomorphs/geo_32x32_lab_exit_01.prefab |
| R1C1 rector | Assets/AssetPrefabs/Complex/Mining/Geomorphs/geo_64x64_mining_reactor_open_HA_01.prefab |
| R2D2 reactor | Assets/AssetPrefabs/Complex/Mining/Geomorphs/geo_64x64_mining_reactor_HA_02.prefab |
| R4B1L1 reactor | Assets/AssetPrefabs/Complex/Tech/Geomorphs/geo_64x64_lab_reactor_HA_01.prefab |
| R2D2 bridge | Assets/AssetPrefabs/Complex/Mining/Geomorphs/Refinery/geo_64x64_mining_refinery_I_HA_01.prefab |
| R4D1 large room | Assets/AssetPrefabs/Complex/Mining/Geomorphs/Refinery/geo_64x64_mining_refinery_X_HA_02.prefab |
| R3A1 activator room | Assets/AssetPrefabs/Complex/Tech/Geomorphs/geo_64x64_Lab_dead_end_room_01.prefab |
| R3B2 activator room | Assets/AssetPrefabs/Complex/Mining/Geomorphs/Refinery/geo_64x64_mining_refinery_dead_end_HA_01.prefab |
| R3D1 activator room | Assets/AssetPrefabs/Complex/Tech/Geomorphs/geo_64x64_Lab_dead_end_room_02.prefab |
| R2D1 cluster room | Assets/AssetPrefabs/Complex/Mining/Geomorphs/Digsite/geo_64x64_mining_dig_site_hub_HA_01.prefab |
| R4D2L2 cluster room | Assets/AssetPrefabs/Complex/Tech/Geomorphs/geo_64x64_tech_lab_hub_HA_02.prefab |
| R4B2L1 cluster room | Assets/AssetPrefabs/Complex/Mining/Geomorphs/Digsite/geo_64x64_mining_dig_site_hub_HA_02.prefab |

### Sound Events

All sounds can be found in [`Wwise-ASM/AK/Events.cs`][Events.cs].

| Nickname | Event Name | Event Id |
| --- | --- | --- |
| R3D1 Door Scream | BIRTHER_DISTANT_ROAR | 3184121378 |
| R4B1 Door Knocking | R4_KNOCKING | 2879789566 |
| R3A1 Terminal Looping | LOG0140_CLIP_LOOP | 2322593598 |
| R3A1 Terminal Stop Looping | LOG0140_CLIP_STOP | 3945188996 |

### EnemySpawningInZone

All enemy groups can be found in [`EnemyGroupDataBlock.json`][EnemyGroupDataBlock]. (Note: Populations can be checked using the [Population Printer][]. This tool requires [.Net][].)

| Group(s) description | Group Type | Difficulty |
| --- | --- | --- |
| Sleepers (no scout) | Hibernate | Easy |
| Sleepers (possible scout) | Hibernate | Medium |
| Sleepers (guaranteed scout) | Hibernate | Hard |
| Scouts | PureDetect | Medium |
| Chargers/Bullrush | Detect | Easy |
| Big Chargers | Detect | Hard |
| Hybrids | PureSneak | Medium |
| Shadows | Hibernate | Biss |
| Big Shadows† | NA | NA |
| Shadow Scouts | PureDetect | Boss |
| Birthers | Hibernate | MiniBoss |
| Tank | Hibernate | MegaBoss |

†Removed from `EnemyPopulationDataBlock`

## Q&A

A list of common questions and answers:

- What if I want to have an enemy population that has two different kinds of things, do I need to create my own Enemy Population in `EnemyPopulationDataBlock.json`? For example shadow scouts and hybrids.
  - No. You can have several groups per zone and so you pick pre-existing groups with what you want and throw them all in the zone.

If you have questions please message me (and may end up on this list).

<!-- links -->
[DPK Releases]: <https://drive.google.com/drive/folders/1i37s30b550z84D27A6i4AsH7p4zlRGdS>
[DPK Template]: <https://docs.google.com/spreadsheets/d/1FLA-eHv9NhU3IdxcdQ29ueW8Qav9IlknaLdb4xbGZCI>
[Original DataBlocks]: <https://cdn.discordapp.com/attachments/736410752210042950/769656075880103936/OriginalDataBlocks.rar>
[Binary Encoder]: <https://gist.github.com/dogpetkid/e08716ceafd8e734ccde09228b0cd0f0>
[GTFO hack]: <https://discord.com/channels/765711589868568596/771076197525618758/771077099620270080>
[Population Printer]: <https://cdn.discordapp.com/attachments/771076197525618758/771468912814063616/population_printer.rar>
[R4 Build]: <https://drive.google.com/uc?id=1qVGNYboEMa27i0N8rR9wGO9wS5gZddkK>

[python]: <https://www.python.org/>
[python downloads]: <https://www.python.org/downloads/>
[numpy]: <https://numpy.org/doc/stable/>
[xlrd]: <https://xlrd.readthedocs.io/en/latest/>
[pandas]: <https://pandas.pydata.org/docs/>

[7-zip]: <https://www.7-zip.org/>

[uabe]: <https://github.com/DerPopo/UABE>
[uabe releases]: <https://github.com/DerPopo/UABE/releases>

[.net]: <https://dotnet.microsoft.com/>

[7-Zip image]: <resources/images/7-zip.png>
[GTFO-modding folder image]: <resources/images/gtfo-modding.png>
[ComplexResourceSetDataBlock]: <../Datablocks/ComplexResourceSetDataBlock.json>
[Events.cs]: <resources/EVENTS.cs>
[EnemyGroupDataBlock]: <../Datablocks/EnemyGroupDataBlock.json>
