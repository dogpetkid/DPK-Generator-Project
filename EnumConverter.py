"""
This is a tool created by DPK

This tool can convert GTFO enums and indexes.
"""

import io

# io:   used to read from files

def indexToEnum(enum:io.FileIO, index:int, force:bool=True):
    """
    Takes an enum file and index
    It will return the name of the enum value with said index
    When force is true, the function will not check if the value is not actually a number
    """
    enum.seek(0)
    lines = enum.readlines()
    try:
        # remove the \n value from the line since the name should not contain the newline
        return lines[index].replace("\n","")
    except IndexError:
        raise IndexError("No enum value exists with that index "+index+" within "+enum.name)
    except TypeError:
        if(force):raise TypeError
        else:return index

def enumToIndex(enum:io.FileIO, name:str, textmode:bool=False):
    """
    Takes an enum file and enum name
    It will return the index of the enum value with said name
    When textmode is true, the enum is to be represented as text (and therefore the function should return)
    """
    if(textmode):return name
    enum.seek(0)
    index = 0
    # iterate through all enum names without the \n (since the file will read \n)
    for line in [l.replace("\n","") for l in enum.readlines()]:
        if name==line:
            return index
        index+= 1
    raise IndexError("No enum value exists with the name \""+name+"\" within "+enum.name)

def enumInDict(enum:io.FileIO, dictionary:dict, key:str, textmode:bool=True):
    """
    Convert an enum into an index from inside of a dictionary
    When textmode is true, the enum is to be represented as text (and therefore the function should return)
    """
    try:dictionary[key] = enumToIndex(enum, dictionary[key], textmode=textmode)
    except KeyError:pass

def indexInDict(enum:io.FileIO, dictionary:dict, key:str, force:bool=False):
    """
    Convert an index into an enum from inside of a dictionary
    When force is true, the function will not check if the value is not actually a number
    """
    try:dictionary[key] = indexToEnum(enum, dictionary[key], force=force)
    except KeyError:pass

if __name__ == "__main__":
    ENUM_eWantedZoneHeighs = open("../Datablocks/TypeList/Enums/eWantedZoneHeighs.txt")
    print("0",indexToEnum(ENUM_eWantedZoneHeighs,0))
    print("LowHigh",enumToIndex(ENUM_eWantedZoneHeighs,"LowHigh"))
    input("Done.")
