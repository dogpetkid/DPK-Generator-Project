"""
This is a tool created by DPK

This tool can convert GTFO enums and indexes.
"""

def indexToEnum(enum, index):
    """
    Takes an enum file and index
    It will return the name of the enum value with said index
    """
    enum.seek(0)
    lines = enum.readlines()
    try:
        # remove the \n value from the line since the name should not contain the newline
        return lines[index].replace("\n","")
    except IndexError:
        raise IndexError("No enum value exists with that index "+index+" within "+enum.name)

def enumToIndex(enum, name):
    """
    Takes an enum file and enum name
    It will return the index of the enum value with said name
    """
    enum.seek(0)
    index = 0
    # iterate through all enum names without the \n (since the file will read \n)
    for line in [l.replace("\n","") for l in enum.readlines()]:
        if name==line:
            return index
        index+= 1
    raise IndexError("No enum value exists with the name \""+name+"\" within "+enum.name)

if __name__ == "__main__":
    ENUM_eWantedZoneHeighs = open("../Datablocks/TypeList/Enums/eWantedZoneHeighs.txt")
    print("0",indexToEnum(ENUM_eWantedZoneHeighs,0))
    print("LowHigh",enumToIndex(ENUM_eWantedZoneHeighs,"LowHigh"))
    input("Done.")
