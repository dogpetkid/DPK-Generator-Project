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

def ctn(col):
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

class interfacer:
    """A tool to read and write to xlsx files to garuntee datatypes, sits on top of pandas DataFrame"""

    def __init__(self, frame):
        if not isinstance(frame, pandas.DataFrame):
            raise TypeError("Interfacer must be passed a pandas.DataFrame")
        self.frame = frame
    
    def isEmpty(self, y, x):
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

    def read(self, readtype, y, x):
        """
        Read will attempt to read a value at x,y and return said type
        """
        if readtype == blankable:
            if self.isEmpty(y,x): return ""
            return str(self.frame.at[y,x])
        
        if self.isEmpty(y,x): raise EmptyCell("Expected cell at x,y: "+str(x)+","+str(y))

        value = self.frame.at[y,x]

        if readtype == str:
            return str(value)
        elif readtype == int:
            return int(value)
        elif readtype == float:
            return float(value)
        elif readtype == bool:
            if type(value) == bool:
                return value
            elif type(value) == str:
                return value!=""
            elif type(value) in [int,float]:
                return value>0
            elif type(value) == numpy.float64:
                # this must be seaparate from the other numbers because the comparison will create a
                # numpy.bool_ which will crash the json dumps
                return bool(value>0)

        raise Exception("Interface failed to read an item of type: "+readtype.__name__)

if __name__ == "__main__":
    i = interfacer(pandas.read_excel("test.xlsx", "Sheet1", header=None))
    print(i.read(str, 0, 0))
    print(i.read(int, 1, 0))
    print(i.read(float, 2, 0))
    print(i.read(bool, 3, 0))
    print(i.read(int, 4, 0))
    input("Done.")
