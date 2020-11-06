"""
This is a tool created by DPK

This tool can interface xlsx files and give out consistant datatypes.
"""

import pandas
import xlrd
import numpy

# pandas:   used to read the excel data
# xlrd:     used to catch and throw excel errors when initially reading the sheets
# numpy:    used to manipulate the inconsistant numpy data read by pandas

def ctn(col:str):
    """"Column To Number", converts a column to its index, e.g. A = 0, Z = 25, AA = 26, AAA = 702"""
    col = col.upper()
    index = 0
    i = 0
    first = True
    for letter in col[::-1]:
        index+= 26**i * (ord(letter)-ord("A") + (not first))
        i+= 1
        first = False
    return index

class blankable(str):
    """
    A signalling class for a possible blank string
    This is because pandas will fail to read a cell if it is blank but some blank cells still need to be read
    """
    pass

class EmptyCell(Exception):
    """EmptyCell is an exception that is raised when a cell lacks a value"""
    pass

class interface:
    """A tool to read and write to xlsx files to garuntee datatypes, sits on top of pandas DataFrame"""

    def __init__(self, frame:pandas.DataFrame):
        if not isinstance(frame, pandas.DataFrame):
            raise TypeError("Interfacer must be passed a pandas.DataFrame")
        self.frame = frame
    
    def isEmpty(self, x:int, y:int):
        """isEmpty will return false if a value is present in the cell"""
        try:
            # if the value isnan, there is no value present
            if (numpy.isnan(self.frame.at[y,x])): return True
        except KeyError:
            # if a key error occurs, there is no value in that cell
            return True
        except TypeError:
            # if a type error occurs, it is because isnan threw an error because the value was not nan, meaning a value exists
            return False
        return False

    def read(self, readtype:object, x:int, y:int):
        """
        Read will attempt to read a value at x,y and return said type
        """

        if readtype == blankable:
            if self.isEmpty(x,y): return ""
            return str(self.frame.at[y,x])
        
        if self.isEmpty(x,y): raise EmptyCell("Expected cell at y,x: "+str(y)+","+str(x))

        rawvalue = self.frame.at[y,x]

        try:
            if readtype == str:
                return str(rawvalue)
            elif readtype == int:
                return int(rawvalue)
            elif readtype == float:
                return float(rawvalue)
            elif readtype == bool:
                if type(rawvalue) == bool:
                    return rawvalue
                elif type(rawvalue) == str:
                    return rawvalue.upper()=="TRUE"
                elif type(rawvalue) in [int,float]:
                    return rawvalue>0
                elif type(rawvalue) == numpy.float64:
                    # this must be seaparate from the other numbers because the comparison will create a
                    # numpy.bool_ which will crash the json dumps
                    return bool(rawvalue>0)
        except ValueError:
            raise ValueError("Interface failed to read an item of type \""+readtype.__name__+"\" at y,x "+str(y)+","+str(x)+" due to an incorrect read type")

        raise Exception("Interface failed to read an item of type \""+readtype.__name__+"\" at y,x "+str(y)+","+str(x))

    def readIntoDict(self, readtype:object, x:int, y:int, dictionary:dict, key:str):
        """
        ReadIntoDict will attempt to read a value at x,y
        If such a value exists, it will set dict[key] to the value
        If no such value exists, it will return None
        """

        try:
            v = self.read(readtype, x, y)
            dictionary[key] = v
            return v
        except EmptyCell:
            return None

if __name__ == "__main__":
    """Test code to read the first 4 cells of a test sheet and put the values into a dict"""
    i = interface(pandas.read_excel("test.xlsx", "Sheet1", header=None))
    print(i.read(str, 0, 0))
    print(i.read(int, 1, 0))
    print(i.read(float, 2, 0))
    print(i.read(bool, 3, 0))
    # print(i.read(int, 4, 0))
    a={}
    i.readIntoDict(str, 0, 0, a, "String1")
    i.readIntoDict(int, 1, 0, a, "Int1")
    i.readIntoDict(float, 2, 0, a, "Float1")
    i.readIntoDict(int, 4, 0, a, "Int2") # should not end up in the dict
    print(a)
    input("Done.")
