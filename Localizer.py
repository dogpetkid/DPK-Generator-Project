"""
This is a tool created by DPK

This tool can handle the data used for GTFO's Localization system from TextDataBlock.
"""

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

# GH-1 there has to be a better way than passing the TextDataBlock to every call of the below functions, same goes for the language setting
def idToLocalizedText(textdatablock:DatablockIO.datablock, persistentId:int, language:typing.Union[str, int]="English"):
    """
    Return the specific localization of a datablock using its id.
    """
    index = textdatablock.find(persistentId)
    if index == None:
        return ""
    if compareLanguageById("English", language):
        return textdatablock.data["Blocks"][index]["English"]
    else:
        LanguageError.notSupported(language)

def localizedtextToId(textdatablock:DatablockIO.datablock, text:str, language:typing.Union[str, int]="English"):
    """
    Return the id of a block that contains the specific localization.
    """
    if compareLanguageById("English", language):
        getlocalization = lambda data: data["English"]
    else:
        LanguageError.notSupported(language)

    for block in textdatablock.data["Blocks"]:
        try:
            # return the block with the text matching the localization we are looking for
            if text == getlocalization(block):
                return block["persistentID"]
               # GH-1 handle missing persistentID in a block in a separate try catch to actually raise the error
        except KeyError: pass
    return None

def localizeFromIdInDict(textdatablock:DatablockIO.datablock, dictionary:dict, key:str, language:typing.Union[str, int]="English"):
    """
    Convert an id into a localized text from inside a dictionary
    """
    try: _ = dictionary[key]
    except KeyError:return
    dictionary[key] = idToLocalizedText(textdatablock, dictionary[key], language=language)

def localizeToIdInDict(textdatablock:DatablockIO.datablock, dictionary:dict, key:str, language:typing.Union[str, int]="English"):
    """
    Convert a localized text into an id from inside a dictionary
    """
    try: _ = dictionary[key]
    except KeyError:return
    dictionary[key] = localizedtextToId(textdatablock, dictionary[key], language=language)

if __name__ == "__main__":
    textdatablock = DatablockIO.datablock(open("../OriginalDatablocks/TextDataBlock.json", "r", encoding="UTF-8"))

    # I'm unsure why, but the debug messages of the test get messed up, it doesn't matter because it properly asserts when incorrect.

    ids = [2, 966, 4233525490, 99999999999999999999999999, 2992419498]
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
    ids = [2, 966, 4233525490, None, 2992419498]

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
        "d": 99999999999999999999999999,
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
        localizeFromIdInDict(textdatablock, dictids, key, language="English")
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
        "d": None,
        "e": 2992419498
    }

    print("Testing localizeToIdInDict")
    for key, value in dictids.items():
        # check functionality
        localizeToIdInDict(textdatablock, texts, key, language="English")
        assert texts[key]==dictids[key], "Result does not match expected ids."
    for key, _ in dictids.items():
        # check crash proof
        localizeToIdInDict(textdatablock, texts, key+key, language="English")
