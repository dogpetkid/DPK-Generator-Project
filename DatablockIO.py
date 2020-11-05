"""
This is a tool created by DPK

This tool can to read and write to datablocks.
"""

import json

# json: used to import data from the json or piece to a json

class datablock:
    """
    A class to interface with the "blocks" part of the datablocks
    Use datablock.data["Blocks"] to interact with the data within the datablock
    """
    def __init__(self, blockfile):
        self.blockfile = blockfile
        self.data = json.load(blockfile)
    
    def find(self, find):
        """
        Finds the index of a block in the blocks array using find
        'find' can be the persistantID (int) or name (str) of the datablock
        Returns -1 if no block found
        """
        search = ["name","persistentID"][isinstance(find,int)]
        i = 0
        for block in self.data["Blocks"]:
            if (block[search] == find): return i
            i+= 1
        return -1

    def writeblock(self):
        """Writes the datablocks data back into its file"""
        self.blockfile.seek(0)
        json.dump(self.data,self.blockfile,ensure_ascii=False,allow_nan=False,indent=2)

def nameToId(datablock, name):
    """Convert a name into an id"""
    try:return datablock.data["Blocks"][datablock.find(name)]["persistentID"]
    except IndexError:raise IndexError("No such block exists with name \""+name+"\" within "+datablock.blockfile.name)

def idToName(datablock, persistantId):
    """Convert an id into a name"""
    try:return datablock.data["Blocks"][datablock.find(persistantId)]["name"]
    except IndexError:raise IndexError("No such block exists with id "+persistantId+" within "+datablock.blockfile.name)

def nameInDict(datablock, dict, key):
    """Convert a name into an id from inside of a dictionary"""
    try:dict[key] = nameToId(datablock, dict[key])
    except KeyError:pass

def idInDict(datablock, dict, key):
    """Convert an id into a name from inside of a dictionary"""
    try:dict[key] = idToName(datablock, dict[key])
    except KeyError:pass

if __name__ == "__main__":
    d = datablock(open("../Datablocks/ChainedPuzzleDataBlock.json", "r+"))
    print(d.find("Single x1"))
    print(d.data["Blocks"][d.find(4)])
    d = datablock(open("../Workspace/WardenObjectiveDataBlock_DPK.json", "r+"))
    print(d.data["Blocks"][d.find(6)])
    d.data["Blocks"][d.find(6)]["DPKTestCounter"]+= 1
    print(d.data["Blocks"][d.find(6)])
    d.writeblock()
