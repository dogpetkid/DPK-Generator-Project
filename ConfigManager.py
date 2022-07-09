"""
This is a tool created by DPK

This tool is used to manage the config of this project
"""

import json
from json.decoder import JSONDecodeError
import os

# os:   used to join the file path
# json: used to parse the config file

global config
try:
    cfg = open(os.path.join(os.path.dirname(__file__), "config.json"), "r+")
    config = json.load(cfg)
except FileNotFoundError:
    cfg = open(os.path.join(os.path.dirname(__file__), "config.json"), "w")
    config = {}
except JSONDecodeError:
    config = {}

try:
    if type(config["Project"]) != dict:raise TypeError
except (KeyError, TypeError):
    config["Project"] = {}
try:
    if type(config["Project"]["Version"]) != str:raise TypeError
except (KeyError, TypeError):
    config["Project"]["Version"] = "7.1"
try:
    if type(config["Project"]["blockpath"]) != str:raise TypeError
except (KeyError, TypeError):
    config["Project"]["blockpath"] = "..\\Datablocks"
try:
    if type(config["Project"]["blockprefix"]) != str:raise TypeError
except (KeyError, TypeError):
    config["Project"]["blockprefix"] = "GameData_"
try:
    if type(config["Project"]["blocksuffix"]) != str:raise TypeError
except (KeyError, TypeError):
    config["Project"]["blocksuffix"] = "_bin.json"

try:
    if type(config["LevelUtility"]) != dict:raise TypeError
except (KeyError, TypeError):
    config["LevelUtility"] = {}
try:
    if type(config["LevelUtility"]["defaultpaths"]) != list:raise TypeError
    for path in config["LevelUtility"]["defaultpaths"]:
        if type(path) != str: raise TypeError
except (KeyError, TypeError):
    config["LevelUtility"]["defaultpaths"] = ["in.xlsx"]

try:
    if type(config["LevelReverseUtility"]) != dict:raise TypeError
except (KeyError, TypeError):
    config["LevelReverseUtility"] = {}
try:
    if type(config["LevelReverseUtility"]["templatepath"]) != str:raise TypeError
except (KeyError, TypeError):
    config["LevelReverseUtility"]["templatepath"] = ".\\Template for Generator R7.xlsx"

cfg.truncate(0)
cfg.seek(0)
json.dump(config, cfg, ensure_ascii=False, allow_nan=False, indent=4)
cfg.close()
