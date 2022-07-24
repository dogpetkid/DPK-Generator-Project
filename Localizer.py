"""
This is a tool created by DPK

This tool can handle the data used for GTFO's Localization system from TextDataBlock.
"""

import re
import typing

import DatablockIO

languages={
    "en": 0,
    "english": 0,
    "fr": 1,
    "french": 1,
    "it": 2,
    "italian": 2,
    "de": 3,
    "german": 3,
    "es": 4,
    "spanish": 4,
    "ru": 5,
    "russian": 5,
    "pt": 6,
    "portuguese": 6,
    "pl": 7,
    "polish": 7,
    "ja": 8,
    "japanese": 8,
    "ko": 9,
    "korean": 9,
    "zh_t": 10,
    "chinese_traditional": 10,
    "zh_s": 11,
    "chinese_simplified": 11,
}

# GH-1 finish the localization system to support creating or filling out localizations

class LanguageError(Exception):
    """
    An exception used by this class to specify language based errors.
    """
    def notSupported(language:typing.Union[str,int]):
        if type(language) == int:
            try: language = getLanguageName(language) # attempt converting the language id to a name to provide a more informative error message
            except LanguageError: # if a language happens from getLanguageName, the language id does not exist
                raise LanguageError("Language of id '" + str(language) + "' is not currently supported.")
        raise LanguageError("Language '" + str(language) + "' is not currently supported.")

def getLanguageName(id:int):
    """
    Used to get the name of a language.
    """
    languagesflipped = {v: k for k, v in languages.items()}
    try:
        return languagesflipped[id]
    except KeyError:
        raise LanguageError("No language has id '" + str(id) + "'.")

def getLanguageId(name:str):
    """
    Used to get the id of a language.
    """
    try: return languages[name.lower()]
    except KeyError as e: raise LanguageError("No languge '" + str(name) + "'.")

def compareLanguageById(a:typing.Union[str,int], b:typing.Union[str,int]):
    """
    Used to compare if two languages are the same.
    """
    try:
        if type(a) == str: a = getLanguageId(a)
        if type(b) == str: b = getLanguageId(b)
        return a == b
    except LanguageError:
        return False

def sanitize(string:str):
    """
    Sanitizes strings of problematic characters the GTFO devs use.
    The characters removed cause IllegalCharacterError and/or UnicodeError.
    @see https://openpyxl.readthedocs.io/en/stable/_modules/openpyxl/utils/exceptions.html?highlight=IllegalCharacterError
    @see https://docs.python.org/3/library/exceptions.html#UnicodeError
    """
    # GH-1 there has a way to escape the character rather than just replacing it
    string = re.sub(chr(0x202c), "?", string)   # used in         16 @   37
    string = re.sub(chr(0x0420), "?", string)   # used in        926 @    0
    string = re.sub(chr(0x0443), "?", string)   # used in        926 @    1
    string = re.sub(chr(0x0441), "?", string)   # used in        926 @    2 &        926 @    3
    string = re.sub(chr(0x043a), "?", string)   # used in        926 @    4
    string = re.sub(chr(0x0438), "?", string)   # used in        926 @    5
    string = re.sub(chr(0x0439), "?", string)   # used in        926 @    6
    string = re.sub(chr(0x65e5), "?", string)   # used in       1383 @    0
    string = re.sub(chr(0x672c), "?", string)   # used in       1383 @    1
    string = re.sub(chr(0x8a9e), "?", string)   # used in       1383 @    2
    string = re.sub(chr(0xd55c), "?", string)   # used in       1384 @    0
    string = re.sub(chr(0xad6d), "?", string)   # used in       1384 @    1
    string = re.sub(chr(0xc5b4), "?", string)   # used in       1384 @    2
    string = re.sub(chr(0x7e41), "?", string)   # used in       1385 @    0
    string = re.sub(chr(0x9ad4), "?", string)   # used in       1385 @    1
    string = re.sub(chr(0x4e2d), "?", string)   # used in       1385 @    2 &       1386 @    2
    string = re.sub(chr(0x6587), "?", string)   # used in       1385 @    3 &       1386 @    3
    string = re.sub(chr(0x7b80), "?", string)   # used in       1386 @    0
    string = re.sub(chr(0x4f53), "?", string)   # used in       1386 @    1
    string = re.sub(chr(0x009d), "?", string)   # used in  420980446 @   71 &  420980446 @  250
    string = re.sub(chr(0x008d), "?", string)   # used in  420980446 @  162
    string = re.sub(chr(0x0103), "?", string)   # used in 2180421988 @  367
    string = re.sub(chr(0x2032), "?", string)   # used in 2180421988 @  369
    string = re.sub(chr(0x0259), "?", string)   # used in 2180421988 @  371
    string = re.sub(chr(0x2033), "?", string)   # used in 2180421988 @  376
    string = re.sub(chr(0x0113), "?", string)   # used in 2180421988 @  377
    string = re.sub(chr(0x2009), "?", string)   # used in 3295463044 @ 1247 & 3295463044 @ 1249
    return string

def mangle(string:str):
    """
    Mangle any characters that could be replaced to be the same.
    E.g. linebreaks could be represented as \\n or \\r\\n
    """
    # GH-1 there has to be a way to more intelligent way to compare strings rather than mangling strings to be the same before compraing them
    string = re.sub("\r?\n", "\r\n", string)
    return string

# GH-1 there has to be a better way than passing the TextDataBlock to every call of the below functions, same goes for the language setting
def idToLocalizedText(textdatablock:DatablockIO.datablock, persistentId:int, language:typing.Union[str, int]="English"):
    """
    Return the specific localization of a datablock using its id. Returns "" when there is no corresponding block.
    """
    if type(persistentId) != int: return "" # GH-1 should this raise an error instead?
    if persistentId == 0: return "" # save time since the devs use id 0 as a placeholder to mean nothing
    index = textdatablock.find(persistentId)
    if index == None:
        return ""
    if compareLanguageById("English", language):
        return sanitize(textdatablock.data["Blocks"][index]["English"])
    else:
        LanguageError.notSupported(language)

def localizedtextToId(textdatablock:DatablockIO.datablock, text:str, language:typing.Union[str, int]="English"):
    """
    Return the id of a block that contains the specific localization. Returns 0 when there is no corresponding block.
    """
    if type(text) != str: return 0 # GH-1 should this raise an error instead?
    if text == "": return 0 # save time since an empty string shouldn't have a localization
    if compareLanguageById("English", language):
        getlocalization = lambda data: data["English"]
    else:
        LanguageError.notSupported(language)

    for block in textdatablock.data["Blocks"]:
        try:
            # return the block with the text matching the localization we are looking for
            if mangle(sanitize(text)) == mangle(sanitize(getlocalization(block))):
                return block["persistentID"]
               # GH-1 handle missing persistentID in a block in a separate try catch to actually raise the error
        except KeyError: pass
    return 0
    # despite None being the better fail return value,
    # the devs give the id 0 to things that are unused so that should be used for the function

def localizeFromIdInDict(textdatablock:DatablockIO.datablock, dictionary:dict, key:str, passthrough:bool=False, language:typing.Union[str, int]="English"):
    """
    Convert an id into a localized text from inside a dictionary.
    @param passthrough will make the function do nothing.
    """
    if passthrough:return
    try: _ = dictionary[key]
    except KeyError:return
    if type(dictionary[key]) == int:
        dictionary[key] = idToLocalizedText(textdatablock, dictionary[key], language=language)

def localizeToIdInDict(textdatablock:DatablockIO.datablock, dictionary:dict, key:str, passthrough:bool=False, force:bool=False, language:typing.Union[str, int]="English"):
    """
    Convert a localized text into an id from inside a dictionary.
    @param passthrough will make the function do nothing.
    @param force will write the id 0 into the dict when the localization isn't found. (Otherwise it will let text without a localiztion pass through. This excludes "" which always localizes to 0.)
    """
    if passthrough: return
    try: _ = dictionary[key]
    except KeyError:return
    if dictionary[key] == "":
        dictionary[key] = 0
        return
    if type(dictionary[key]) == str:
        id = localizedtextToId(textdatablock, dictionary[key], language=language)
        if force or id != 0: dictionary[key] = id

if __name__ == "__main__":
    textdatablock = DatablockIO.datablock(open("../OriginalDatablocks/TextDataBlock.json", "r", encoding="UTF-8"))

    # I'm unsure why, but the debug messages of the test get messed up, it doesn't matter because it properly asserts when incorrect.

    ids = [2, 966, 4233525490, 0, 2992419498]
    texts = ["CONNECT TO RUNDOWN", "[Voice chat active]", "Bat", "",
        "D-Lock Block Cipher\r\nalias:int_server.1024_ciph.tier5.phys_ops/DLockwoodA074.flagged\r\n\nMr. Lockwood,\r\n\nYou requested an update. Here it is.\r\n\nThrough my experiments, I have discovered the virus has a multipronged attack strategy, which not only\r\nincreases the R0, but also allows it to be highly contagious in multiple environments. Infected individual\r\ncan release create pathogens and fomites through any bodily fluid. The virus has been observed in fluids\r\nextracted from anywhere on the patient - sweat, blood, saliva, bile, urine, even in the aqueous and\r\nvitreous humors. The virus can also survive for several weeks while airborne or on surfaces. In my control\r\ntests, subjects exposed to locations that had been contaminated up to 4 weeks prior to exposure still\r\ncontracted the virus. This was also true of subjects immersed in various solutions (I tried fresh water, salt\r\nwater, and several synthetic oils and polymers). 85% of the subjects contracted the virus within a few\r\nhours. The remaining 15% who survived other exposure methods contracted the virus when exposed to a\r\nvaporous environment heavily dosed with infected fomites.\r\n\nThe parasite that we had previously attributed as the primary carrier of the virus is, in my opinion, merely\r\nanother victim of this remarkable lifeform. The only difference is the virus does not mutate the parasites,\r\nrather it prolongs the life cycle of the parasite indefinitely. The relationship is symbiotic. The parasite\r\ncarries the virus to new hosts, and the virus helps the parasite live for an extended (and currently\r\nunknown) period. Perhaps indefinitely.\r\n\nThe virus is remarkable. It appears to have been perfectly designed to infect regardless of the\r\ncircumstance. When it finds a host, it induces coughing, sweating, sneezing, and ultimately violent attacks\r\nto draw blood and saliva. Without a host, it can transfer itself through any medium. The only test I have\r\nnot been able to perform yet is in a vacuum, but I am planning such experiments next week as soon as I\r\nreceive new subjects from Mr. Piros.\r\n\nResearch continues.\r\n\nDr Abeo Dauda A153"
    ]

    print("Testing idToLocalizedText")
    for i in range(len(ids)):
        # check functionality
        result = idToLocalizedText(textdatablock, ids[i], language="English")
        print("Found text '" + result.split("\n")[0] + "' for " + str(ids[i]) + ", '" + texts[i].split("\n")[0] + "'")
        assert result==texts[i], "Result does not match expected text"

    texts = ["CONNECT TO RUNDOWN", "[Voice chat active]", "Bat", "unga bunga",
        "D-Lock Block Cipher\r\nalias:int_server.1024_ciph.tier5.phys_ops/DLockwoodA074.flagged\r\n\nMr. Lockwood,\r\n\nYou requested an update. Here it is.\r\n\nThrough my experiments, I have discovered the virus has a multipronged attack strategy, which not only\r\nincreases the R0, but also allows it to be highly contagious in multiple environments. Infected individual\r\ncan release create pathogens and fomites through any bodily fluid. The virus has been observed in fluids\r\nextracted from anywhere on the patient - sweat, blood, saliva, bile, urine, even in the aqueous and\r\nvitreous humors. The virus can also survive for several weeks while airborne or on surfaces. In my control\r\ntests, subjects exposed to locations that had been contaminated up to 4 weeks prior to exposure still\r\ncontracted the virus. This was also true of subjects immersed in various solutions (I tried fresh water, salt\r\nwater, and several synthetic oils and polymers). 85% of the subjects contracted the virus within a few\r\nhours. The remaining 15% who survived other exposure methods contracted the virus when exposed to a\r\nvaporous environment heavily dosed with infected fomites.\r\n\nThe parasite that we had previously attributed as the primary carrier of the virus is, in my opinion, merely\r\nanother victim of this remarkable lifeform. The only difference is the virus does not mutate the parasites,\r\nrather it prolongs the life cycle of the parasite indefinitely. The relationship is symbiotic. The parasite\r\ncarries the virus to new hosts, and the virus helps the parasite live for an extended (and currently\r\nunknown) period. Perhaps indefinitely.\r\n\nThe virus is remarkable. It appears to have been perfectly designed to infect regardless of the\r\ncircumstance. When it finds a host, it induces coughing, sweating, sneezing, and ultimately violent attacks\r\nto draw blood and saliva. Without a host, it can transfer itself through any medium. The only test I have\r\nnot been able to perform yet is in a vacuum, but I am planning such experiments next week as soon as I\r\nreceive new subjects from Mr. Piros.\r\n\nResearch continues.\r\n\nDr Abeo Dauda A153"
    ]
    ids = [2, 966, 4233525490, 0, 2992419498]

    print("Testing localizedtextToId")
    for i in range(len(texts)):
        # check functionality
        result = localizedtextToId(textdatablock, texts[i], language="English")
        print("Found id '" + str(result) + "' for " + str(ids[i]) + ", '" + texts[i].split("\n")[0] + "'")
        assert result==ids[i], "Result does not match expected id"

    dictids = {
        "a": 2,
        "b": 966,
        "c": 4233525490,
        "d": 0,
        "e": 2992419498
    }
    texts = {
        "a": "CONNECT TO RUNDOWN",
        "b": "[Voice chat active]",
        "c": "Bat",
        "d": "",
        "e": "D-Lock Block Cipher\r\nalias:int_server.1024_ciph.tier5.phys_ops/DLockwoodA074.flagged\r\n\nMr. Lockwood,\r\n\nYou requested an update. Here it is.\r\n\nThrough my experiments, I have discovered the virus has a multipronged attack strategy, which not only\r\nincreases the R0, but also allows it to be highly contagious in multiple environments. Infected individual\r\ncan release create pathogens and fomites through any bodily fluid. The virus has been observed in fluids\r\nextracted from anywhere on the patient - sweat, blood, saliva, bile, urine, even in the aqueous and\r\nvitreous humors. The virus can also survive for several weeks while airborne or on surfaces. In my control\r\ntests, subjects exposed to locations that had been contaminated up to 4 weeks prior to exposure still\r\ncontracted the virus. This was also true of subjects immersed in various solutions (I tried fresh water, salt\r\nwater, and several synthetic oils and polymers). 85% of the subjects contracted the virus within a few\r\nhours. The remaining 15% who survived other exposure methods contracted the virus when exposed to a\r\nvaporous environment heavily dosed with infected fomites.\r\n\nThe parasite that we had previously attributed as the primary carrier of the virus is, in my opinion, merely\r\nanother victim of this remarkable lifeform. The only difference is the virus does not mutate the parasites,\r\nrather it prolongs the life cycle of the parasite indefinitely. The relationship is symbiotic. The parasite\r\ncarries the virus to new hosts, and the virus helps the parasite live for an extended (and currently\r\nunknown) period. Perhaps indefinitely.\r\n\nThe virus is remarkable. It appears to have been perfectly designed to infect regardless of the\r\ncircumstance. When it finds a host, it induces coughing, sweating, sneezing, and ultimately violent attacks\r\nto draw blood and saliva. Without a host, it can transfer itself through any medium. The only test I have\r\nnot been able to perform yet is in a vacuum, but I am planning such experiments next week as soon as I\r\nreceive new subjects from Mr. Piros.\r\n\nResearch continues.\r\n\nDr Abeo Dauda A153"
    }

    print("Testing localizeFromIdInDict")
    for key, value in dictids.items():
        # check functionality
        localizeFromIdInDict(textdatablock, dictids, key, passthrough=False, language="English")
        assert dictids[key]==texts[key], "Result does not match expected text."
    for key, _ in dictids.items():
        # check crash proof
        localizeFromIdInDict(textdatablock, dictids, key+key, language="English")

    texts = {
        "a": "CONNECT TO RUNDOWN",
        "b": "[Voice chat active]",
        "c": "Bat",
        "d": "unga bunga",
        "e": "D-Lock Block Cipher\r\nalias:int_server.1024_ciph.tier5.phys_ops/DLockwoodA074.flagged\r\n\nMr. Lockwood,\r\n\nYou requested an update. Here it is.\r\n\nThrough my experiments, I have discovered the virus has a multipronged attack strategy, which not only\r\nincreases the R0, but also allows it to be highly contagious in multiple environments. Infected individual\r\ncan release create pathogens and fomites through any bodily fluid. The virus has been observed in fluids\r\nextracted from anywhere on the patient - sweat, blood, saliva, bile, urine, even in the aqueous and\r\nvitreous humors. The virus can also survive for several weeks while airborne or on surfaces. In my control\r\ntests, subjects exposed to locations that had been contaminated up to 4 weeks prior to exposure still\r\ncontracted the virus. This was also true of subjects immersed in various solutions (I tried fresh water, salt\r\nwater, and several synthetic oils and polymers). 85% of the subjects contracted the virus within a few\r\nhours. The remaining 15% who survived other exposure methods contracted the virus when exposed to a\r\nvaporous environment heavily dosed with infected fomites.\r\n\nThe parasite that we had previously attributed as the primary carrier of the virus is, in my opinion, merely\r\nanother victim of this remarkable lifeform. The only difference is the virus does not mutate the parasites,\r\nrather it prolongs the life cycle of the parasite indefinitely. The relationship is symbiotic. The parasite\r\ncarries the virus to new hosts, and the virus helps the parasite live for an extended (and currently\r\nunknown) period. Perhaps indefinitely.\r\n\nThe virus is remarkable. It appears to have been perfectly designed to infect regardless of the\r\ncircumstance. When it finds a host, it induces coughing, sweating, sneezing, and ultimately violent attacks\r\nto draw blood and saliva. Without a host, it can transfer itself through any medium. The only test I have\r\nnot been able to perform yet is in a vacuum, but I am planning such experiments next week as soon as I\r\nreceive new subjects from Mr. Piros.\r\n\nResearch continues.\r\n\nDr Abeo Dauda A153"
    }
    dictids = {
        "a": 2,
        "b": 966,
        "c": 4233525490,
        "d": 0,
        "e": 2992419498
    }

    print("Testing localizeToIdInDict")
    for key, value in dictids.items():
        # check functionality
        localizeToIdInDict(textdatablock, texts, key, passthrough=False, force=True, language="English")
        assert texts[key]==dictids[key], "Result does not match expected ids."
    for key, _ in dictids.items():
        # check crash proof
        localizeToIdInDict(textdatablock, texts, key+key, language="English")

    # some extra code to debug Monster because it has an invalid character according to openpyxl
    # here is the error produced by openpyxl: https://openpyxl.readthedocs.io/en/stable/_modules/openpyxl/utils/exceptions.html?highlight=IllegalCharacterError
    print("Extra debug for Monster")
    ids = [
        # all unique text ids used throughout the first layer
        0,
        1460368036,
        1561677308,
        1562516213,
        1681134989,
        1704474804,
        1775975616,
        1775975616,
        2374160068,
        2532265839,
        2565679904,
        2565679904,
        2668534995,
        2782882580,
        3171715834,
        3993990468,
        4025272896,
        4028253467,
        420980446, # <-- this string is the issue and even crashes python's print
        # here is the error it produces: https://docs.python.org/3/library/exceptions.html#UnicodeError
        972,
        972465867,
    ]
    for id in ids:
        print("id " + str(id))
        text = idToLocalizedText(textdatablock, id)
        print("Text:\n", text)
        print("--------------------------------------------------")
    # the invalid characters in question is \x9d and \x8d

    # of course there are other strings that cause issues, so instead it may be more useful to just print them all and see who crashes
    print("\n==================================================\n")
    print("Dumping ALL localizations (this may take a while)...")

    def problempositions(errorstring):
        """
        Take a UnicodeError string and parse out the indexes of characters causing the issue
        """
        # find the hex of the problem characters
        strings = re.split("position |: ", errorstring)
        if (len(strings) < 2): raise Exception("error string lacks positional information" + errorstring) # if the positions in the error string cannot be found, continue
        positions = strings[1].split("-")
        start = int(positions[0])
        end = int(positions[1]) if len(positions) > 1 else int(positions[0])
        return start,end

    problemcharacterlist=[]
    problems=[]
    for localization in textdatablock.data["Blocks"]:
    # if True:
        # localization = textdatablock.data["Blocks"][textdatablock.find(3295463044)]
        print("--------------------------------------------------")
        text = localization["English"]
        id = localization["persistentID"]

        offset = 0
        while offset < len(text):
            remainingtext = text[offset:]
            errorstring = ""
            try:
                # all breaks need to be replaced to normal characters
                remainingtext = re.sub("\r|\n", "N", remainingtext)
                # this is because UnicodeError treats \r\n as taking 1 position where as indexing the string treats \r\n as 2 positions
                # using a normal character will make both treat \r and \n as individual characters
                print(remainingtext)
            except UnicodeError as e:
                errorstring = str(e)

            if len(errorstring) == 0: break # if there was no current or caught issue, continue

            start,end = problempositions(errorstring)
            tmp = []
            for i in range(start, end+1):
                problemcharacter = ord(remainingtext[i])
                try:
                    usedindex = problemcharacterlist.index(problemcharacter)
                    problems[usedindex]+= " & %10u @%5u" % (id, offset+i)
                except ValueError:
                    problemcharacterlist.append(problemcharacter)
                    problems.append("    string = re.sub(chr(0x%04x), \"?\", string)   # used in %10u @%5u" % (ord(remainingtext[i]), id, offset+i))

            offset+= end+1

    print("--------------------------------------------------")
    print("This code should be used as the code in the sanitize() function:")
    print("\n".join(problems))
