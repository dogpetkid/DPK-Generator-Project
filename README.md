# DPK Generator Project

<!-- Author: Dogpetkid -->
<!-- Date: 2020-11-06 -->
<!-- Revised: 2021-06-04 -->

This is a project used assist in the creation and understanding of levels in GTFO. This is done by converting the levels from the JSON Datablocks to a custom Excel sheet. For an explanation on how to use this project, see [How to use](#how-to-use).

[Project releases.][Releases]

## Project dependencies

- [numpy][]
- [openpyxl][]
- [pandas][]
- [xlrd][]

## Project contents

- [XlsxInterfacer.py](#xlsxinterfacer)
- [DatablockIO.py](#datablockio)
- [EnumConverter.py](#enumconverter)
- [LevelUtility.py](#levelutility)
- [LevelReverseUtility.py](#levelreverseutility)
- [Template](#template)

---

## XlsxInterfacer

A module created to interface with the pandas.DataFrame to read typed values from cells and to write and save a pandas.DataFrame back to an Excel file.

## DatablockIO

A module created to read from and write to GTFO Datablocks including search blocks based on name or persistantID.

## EnumConverter

A module created to convert between GTFO Enumerator names and indexes. This module is obsolete as of the release of [Rundown 005 Rebirth][Rundown005] because enumerators in the DataBlocks are represented as strings instead of integers.

---

## LevelUtility

This is a file that is able to take a templated .xlsx file and convert it into GTFO ExpeditionInTier, LevelLayout, and WardenObjective data. It takes arguments of which files to convert, converts the .xlsx into dictionaries, and then writes those into the DataBlocks. It is also a Windows Drop Target so .xlsx files can be dragged and dropped to run the program.

## LevelReverseUtility

This is a file that is able to search and convert a Level in the JSON DataBlocks a the templated .xlsx. It takes arguments of search terms for the levels to convert and writes the dictionaries to the .xlsx.

### Template

[Template is here.][DPK Template]

Notes:

- The sheet tries it's best to stop you from doing illegal or irrational things for data types (ie you can't put a word where a number should go) but it is not idiot proof and cannot detect logical issues (eg to go into a zone with a key, it requires that key or a level doesn't have enough cells).
- Most values can be left blank, if leaving a value blank could be problematic, the LevelUtility will probably not run and explain which cell is blank.

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
  - Modding folder
  - MTFO
  - Python and dependencies
  - (Unity) Asset Bundle Extractor (UABE) (for releasing blocks) & Binary Encoder
- [Template Usage](#template-usage)
  - Level Design (crash course)
  - Meta (placement in rundown)
- [Tool Usage](#tool-usage)
  - [JSON to Excel](#json-to-excel)
  - [Excel to JSON](#excel-to-json)

### Setup

#### Modding Folder

Create a `GTFO-modding` folder (or use an existing folder, it's name does not particularly matter, for clarity I will refer to it by that name). In `GTFO-modding` place this project's folder `GTFO-modding`.

#### MTFO

Install [MTFO][] to your GTFO (follow its instructions). MTFO source [here][MTFO Github].

##### Original Datablocks

Obtain the Original Datablocks from MTFO. After obtaining the blocks, create a `GTFO-modding` folder and create a `Datablocks` folder inside of it. Move just the DataBlocks (the JSONs) into the `Datablocks` folder.

#### [Python][]

Download Python [here][python downloads] and install it (all the default settings are fine). Open a terminal (`Windows+R` and open CMD) and run:

`pip install pandas openpyxl numpy xlrd`

If you get `'pip' is not recognized as an internal or external command, operable program or batch file.`, try replacing `pip` with `py -m pip` or `python -m pip`.

This will download all the needed [dependencies listed above](#project-dependencies) for the DPK Generator Project.

---

### Template Usage

See [Template](#template) to see notes about the template. To learn what values mean, I suggest using the LevelReverseUtility to get a couple of examples of the dev's levels.

#### Level Design (crash course)

This section **will not** teach you how to make good levels but what it will do is walk you through the very general steps for creating the idea for the level and some tips.

##### Create a theme

A level that is unique is interesting and a level that follows a single idea seems more coherent (versus being all over). A good way to do this is create something the level will focus on or a rule for the level (e.g. only use specific enemies, all alarms are the same kind, sectors unlock zones in other sectors).

##### Floorplan

The next step is then creating a model of the level to help visualize the way the level works. This can be done whichever way, I prefer drawing the level in MS Paint or coloring different squares in Google Sheets. A thing to keep in mind while placing zones is which secuirty doors and progression puzzles and chained puzzles they take to open and make notes of where the key or cell will come from.

##### Tips

- Backtracking is not fun, try to keep the amount of walking through empty rooms to a minimum.
- A cell in one sector can open a door in another sector.
- Try to walk through the way the level should be completed and double check if "cheesing" by skipping required zones and accidental softlocks aren't a thing.

---

#### Meta

This describes where in the rundown the level is placed. This is for the tool to know where in the `RundownDataBlock.json` to put the ExpeditionInTier. E.g. Rundown ID 25, TierA, Index 1 describes (R4)A2 Foster

---

### Tool Usage

The following are the steps to use the tools.

For CLI instructions, run either Utility with the `-h` or `--help` option.

#### JSON to Excel

1. Follow the [Setup](#setup)
2. Run the LevelReverseUtility (or run the tool in CLI)
3. Specify level(s) (in the format specified by the tool) to spit out Excel files for 

**Note: the formatting may not cover all the cells that have data, you may need to copy the formatting with the Paint Format tool in Excel or Google sheets.**
(This is due to a limitation with openpyxl where it cannot read data validation, so instead of coloring cells but leaving them without data validation, I left all cell formatting unchanged.)

#### Excel to JSON

1. Follow the [Setup](#setup)
2. Design a level or edit an existing level (after using the LevelReverseUtility and skip to step 6)
3. Create a copy of the [Template][DPK Template] in either Google Sheets or download it as a .xlsx (to edit in Excel, Google Sheets is recommended)
4. Fill out the copy of the template.
5. Download the Google Sheets as an Excel file (.xlsx) if you didn't in step 3
6. Drag the .xlsx onto the LevelUtility (or run the tool in CLI).

**Note: the only Rundown, LevelLayout, and WardenObjective datablocks are edited by the LevelUtility**

---

## Feedback

If you find a bug, please report it [here][Issues].

If you have questions please message me.

<!-- links -->
[Releases]: <https://github.com/dogpetkid/DPK-Generator-Project/releases>
[Issues]: <https://github.com/dogpetkid/DPK-Generator-Project/issues>

[DPK Template]: <https://docs.google.com/spreadsheets/d/1ENa6McJnomHa5ugB-VBFjMF62nslj4VwdgM_5ERVRqw>

[python]: <https://www.python.org/>
[python downloads]: <https://www.python.org/downloads/>
[numpy]: <https://numpy.org/doc/stable/>
[openpyxl]: <https://openpyxl.readthedocs.io/en/stable/>
[pandas]: <https://pandas.pydata.org/docs/>
[xlrd]: <https://xlrd.readthedocs.io/en/latest/>

[Rundown005]: <https://gtfo.fandom.com/wiki/Rundown_005>
[MTFO]: <https://gtfo.thunderstore.io/package/dakkhuza/MTFO/>
[MTFO Github]: <https://github.com/GTFO-Modding/MTFO>
