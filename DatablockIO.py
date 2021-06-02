"""
This is a tool created by DPK

This tool can to read and write to datablocks.
"""

import json
import io
import typing

# json:     used to import data from the json or piece to a json
# io:       used to read from and write to files
# typing:   used to give types to function parameters

class datablock:
    """
    A class to interface with the "blocks" part of the datablocks \n
    Use datablock.data["Blocks"] to interact with the data within the datablock
    """
    def __init__(self, blockfile:io.FileIO):
        self.blockfile = blockfile
        self.data = json.load(blockfile)

    # TODO fix this function so it does not return the index when a block is not present (should return negative 1)
    def find(self, find: typing.Union[int, str]):
        """
        Finds the index of a block in the blocks array using find \n
        'find' can be the persistantID (int) or name (str) of the datablock \n
        Returns None if no block found
        """
        search = ["name","persistentID"][isinstance(find,int)]
        i = 0
        for block in self.data["Blocks"]:
            if (block[search] == find): return i
            i+= 1
        return None

    def writeblock(self, block:dict):
        """
        Writes a block \n
        It will add the block if the persistentID does not already exist and override the existing block \n
        (Note: this uses persistentID instead of name because no two blocks should have the same id) \n
        """
        blockindex = self.find(block["persistentID"])
        if blockindex == None:
            self.data["Blocks"].append(block)
        else:
            self.data["Blocks"][blockindex] = block

    def writedatablock(self):
        """Writes the datablocks data back into its file"""
        self.blockfile.truncate(0)
        self.blockfile.seek(0)
        json.dump(self.data,self.blockfile,ensure_ascii=False,allow_nan=False,indent=2)

def nameToId(block:datablock, name:str):
    """Convert a name into an id"""
    try:return block.data["Blocks"][block.find(name)]["persistentID"]
    except IndexError:raise IndexError("No such block exists with name \""+name+"\" within "+block.blockfile.name)
    except TypeError:raise TypeError("No such block exists with name \""+name+"\" within "+block.blockfile.name)

def idToName(block:datablock, persistentId:int):
    """Convert an id into a name"""
    try:return block.data["Blocks"][block.find(persistentId)]["name"]
    except IndexError:raise IndexError("No such block exists with id "+str(persistentId)+" within "+block.blockfile.name)
    except TypeError:raise TypeError("No such block exists with id "+str(persistentId)+" within "+block.blockfile.name)

def nameInDict(block:datablock, dictionary:dict, key:str):
    """Convert a name into an id from inside of a dictionary"""
    try:dictionary[key] = nameToId(block, dictionary[key])
    except KeyError:pass

def idInDict(block:datablock, dictionary:dict, key:str):
    """Convert an id into a name from inside of a dictionary"""
    try:dictionary[key] = idToName(block, dictionary[key])
    except KeyError:pass

if __name__ == "__main__":
    d = datablock(open("../Datablocks/ChainedPuzzleDataBlock.json", "r+"))
    print(d.find("Single x1"))
    print(d.data["Blocks"][d.find(4)])
    d = datablock(open("../Workspace/WardenObjectiveDataBlock_DPK.json", "r+"))
    print(d.data["Blocks"][d.find(6)])
    d.data["Blocks"][d.find(6)]["DPKTestCounter"]+= 1
    print(d.data["Blocks"][d.find(6)])
    d.writedatablock()
    input("Done.")
