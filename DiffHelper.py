"""
This is a tool created by DPK

This tool can is meant to take in a git diff from a file and remove all insignificant changes.
"""

patternstoignore = [
    # 0ed ids
    '\\S\\s*"ChainPuzzle": 0',
    '\\S\\s*"DialogueID": 0',
    '\\S\\s*"FogSetting": 0',
    '\\S\\s*"WaveSettings": 0',
    '\\S\\s*"WavePopulation": 0',
    '\\S\\s*"EnemyID": 0',
    '\\S\\s*"IntelMessage": 0',
    '\\S\\s*"WardenIntel": 0',
    '\\S\\s*"SoundSubtitle": 0',
    '\\S\\s*"ChainedPuzzleToEnter": 0',
    '\\S\\s*"EnemyGroupInfrontOfDoor": 0',
    '\\S\\s*"EnemyGroupInArea": 0',
    '\\S\\s*"BigPickupDistributionInZone": 0',
    '\\S\\s*"CustomSubObjectiveHeader": 0',
    '\\S\\s*"CustomSubObjective": 0',
    '\\S\\s*"ConsumableDistributionInZone": 0',
    '\\S\\s*"CustomInfoText": 0',
    '\\S\\s*"Output": 0',
    '\\S\\s*"CommandDesc": 0',
    '\\S\\s*"CustomText": 0',
    '\\S\\s*"LightSettings": 0',
    # empty arrays
    '\\S\\s*"EventsOnBossDeath": \\[\\]',
    '\\S\\s*"WorldEventChainedPuzzleDatas": \\[\\]',
    '\\S\\s*"EnemySpawningInZone": \\[\\]',
    '\\S\\s*"EnemyRespawnExcludeList": \\[\\]',
    '\\S\\s*"SpecificPickupSpawningDatas": \\[\\]',
    '\\S\\s*"TerminalZoneSelectionDatas": \\[\\]',
    '\\S\\s*"EventsOnTrigger": \\[\\]',
    '\\S\\s*"TerminalPlacements": \\[\\]',
    '\\S\\s*"EventsOnEnter": \\[\\]',
    '\\S\\s*"EventsOnPortalWarp": \\[\\]',
    # empty strings
    '\\S\\s*"AliasPrefixShortOverride": ""',
    '\\S\\s*"AliasPrefixOverride": ""',
    '\\S\\s*"CustomGeomorph": ""',
    '\\S\\s*"PasswordHintText": ""',
    '\\S\\s*"Command": ""',
    '\\S\\s*"WardenIntel": ""',
]

if __name__ == "__main__":
    import argparse
    import re
    # import Localizer

    parser = argparse.ArgumentParser()
    parser.add_argument('path', type=str)
    args = parser.parse_args()

    with open(args.path, "r", encoding='utf8') as file:
        content = file.read()

    content = content.splitlines()
    linenum = 0
    changecount = 0
    remainingchangecount = 0
    insignificantcount = 0
    output = []
    while linenum < len(content):
        if (content[linenum][:10] == "diff --git"):
            # diff headings should be copied
            output+= content[linenum:linenum+4]
            linenum+= 4
            continue
        if (content[linenum][:2]+content[linenum][-2:] == "@@@@"):
            # the start of a chunk will start a search to find the end of the chunk
            # chunkstart and chunkend will be the line numbers of the first line in the chunk and the line after the chunk
            chunkstart = linenum+1
            chunkend = linenum+1
            while chunkend < len(content):
                if (content[chunkend][:10] == "diff --git"):break
                if (content[chunkend][:2]+content[chunkend][-2:] == "@@@@"):break
                chunkend+=1
            chunkcontent = content[chunkstart:chunkend]

            chunklinenum = 0
            # check every line in the chunk to see if they should be filtered by pattern
            while chunklinenum < len(chunkcontent):
                if (chunkcontent[chunklinenum][0] in ['+', '-']): changecount+= 1
                # filter matching lines
                for pattern in patternstoignore:
                    if (matches:=re.match(pattern, chunkcontent[chunklinenum])):
                        chunkcontent = chunkcontent[:chunklinenum] + chunkcontent[chunklinenum+1:]
                        insignificantcount+= 1
                        break
                if not matches: chunklinenum+= 1

            chunklinenum = 0
            # ignore changes that are only different by a comma
            while chunklinenum < len(chunkcontent)-1: # -1 because the last line doesn't have a following line
                line1 = chunkcontent[chunklinenum]
                line2 = chunkcontent[chunklinenum+1]
                try:
                    # attempt to remove the + or - from diffs
                    line1 = line1[1:]
                    line2 = line2[1:]
                except IndexError:
                    chunklinenum+= 1
                    continue
                if line1 == line2+',':
                    # remove both lines and count both lines
                    chunkcontent = chunkcontent[:chunklinenum] + chunkcontent[chunklinenum+2:]
                    insignificantcount+= 2
                    continue
                chunklinenum+= 1

            chunksignificant = False
            for line in chunkcontent:
                if (line[0] in ['+', '-']):
                    chunksignificant = True
                    remainingchangecount+= 1

            if chunksignificant:
                output.append(content[linenum]) # add the chunk header
                output+= chunkcontent # add the edited content
            linenum = chunkend
            continue
        assert False, "Failed to find diff header or start of a hunk"

    with open(args.path, "w", encoding='utf8') as file:
        for line in output:
            # print(Localizer.sanitizeUnicodeError(line))
            file.write(line+"\n")

    print("Original diff size %u" % len(content))
    print("New diff size %u" % len(output))
    print("Original change count %u" % changecount)
    print("Remaining change count %u" % remainingchangecount)
    print("Insignificant changes filtered %u" % insignificantcount)
    print("Done.")
