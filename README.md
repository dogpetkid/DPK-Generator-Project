# DPK Generator Project

<!-- Author: Dogpetkid -->
<!-- Date: 2020-11-06 -->
<!-- Revised: 2020-11-06 -->

This is a project to generate levels for GTFO.

[Project releases.](https://drive.google.com/drive/folders/1i37s30b550z84D27A6i4AsH7p4zlRGdS)

## Project dependencies

- [numpy](https://numpy.org/doc/stable/)
- [xlrd](https://xlrd.readthedocs.io/en/latest/)
- [pandas](https://pandas.pydata.org/docs/)

## Project contents

- [XlsxInterfacer.py](#XlsxInterfacer)
- [DatablockIO.py](#DatablockIO)
- [EnumConverter.py](#EnumConverter)
- [LevelUtility.py](#LevelUtility)
  - [Template](#Template)

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

[Template is here](https://docs.google.com/spreadsheets/d/1FLA-eHv9NhU3IdxcdQ29ueW8Qav9IlknaLdb4xbGZCI).

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
