import sys
import xlrd
import string
import unicodedata
import warnings
import getopt
import argparse
from os import listdir
from os import walk
from easygui import *

# C:\Python27\python.exe CheckBookAudio.py "Sheet1 A1"
# CheckBookAudio.exe "Sheet1 A1"

# args:
# path to audios
# path to excel lut
# name of tabs

# To Run: 
# python CheckBookAudio.py "O
# :/Projects/IL - Activities/[Shared Media]/Level 2/Shared Word Rec and Reading/Au
# dio/Narration/English - Alex" "C:/Users/adam.christiansen.IMAGINELEARNING/Downlo
# ads/Start Reading English - All 3 Modes (decodables) LUT.xlsx" "C:/Users/adam.ch
# ristiansen.IMAGINELEARNING/Desktop/" "58 - The Sun and the North Wind" "59 - Sto
# ne Soup" "60 - The Story of Watermelon"

# "O:/Projects/IL - Activities/[Shared Media]/Level 2/Shared Word Rec and Reading/Audio/Narration/English - Alex" "C:/Users/adam.christiansen.IMAGINELEARNING/Downloads/Start Reading English - All 3 Modes (decodables) LUT.xlsx" "C:/Users/adam.christiansen.IMAGINELEARNING/Desktop/" "58 - The Sun and the North Wind" "59 - Stone Soup" "60 - The Story of Watermelon"
# http://il2tfs/sites/DefaultCollection/ILE/ilactivitydocs/Specs/Look Up Tables/Start Reading English - All 3 Modes (decodables) LUT.xlsx


# def ParseArguments():
#     parser = argparse.ArgumentParser()
#     parser.add_argument("AudioFile", help="Path to existing audio files to check against")
#     parser.add_argument("ExcelLUT", help="Path to LUT where book data is stored")
#     parser.add_argument("OutputPath", help="Path to location for output files")
#     # parser.add_argument("ExcelStartLocation", help="The location in the LUT to begin reading from. Ex: B14")
#     parser.add_argument("Tabs", nargs=argparse.REMAINDER, help="Name of tabs in LUT to parse")

#     args = parser.parse_args()
#     pathToAudioWeHave = args.AudioFile.replace("\\", "/")
#     pathToBookText = args.ExcelLUT.replace("\\", "/")
#     pathToOutputFile = args.OutputPath.replace("\\", "/")
#     # startingLocation = args.ExcelStartLocation
#     sheetNamesToParse = args.Tabs

#     return pathToBookText, pathToAudioWeHave, pathToOutputFile, sheetNamesToParse

def ParseArguments():
    parser = argparse.ArgumentParser()
    parser.add_argument("Tabs", nargs=argparse.REMAINDER, help="Name of tabs in LUT to parse")

    args = parser.parse_args()
    sheetNamesToParse = args.Tabs

    return sheetNamesToParse

def CompareAudioAndText(sheetName):
    #check if we're already using that sheet name
    #this is the case if they want to parse multiple columns in the same LUT
    i = 0
    while sheetName in sheetNames:
        if i != 0:
            sheetName = sheetName[:-1]
        i += 1
        sheetName = sheetName + str(i)
    sheetNames.append(sheetName)

    out = open(pathToOutputFile + sheetName + ".txt", "w")
    bookText.sort()

    warnings.filterwarnings('error')
    warningcount = 0
    output = []
    missingAudios = []
    humanCheck = {}
    for word in bookText:
        try:
            if word in wordToAudioName.keys():
                out.write("%s %16s\n" % (word, wordToAudioName[word]))
                output.append(wordToAudioName[word])
            else:
                #check word without '
                if "'" in word:
                    tempWord = word.replace("'", "")
                    if tempWord in wordToAudioName.keys():
                        out.write(word + "\t\t" + wordToAudioName[tempWord] + "**\n")
                        output.append(wordToAudioName[tempWord])
                        humanCheck[tempWord] = wordToAudioName[tempWord]
                        continue
                out.write(word + "\t \n")
                # output.append(narratorName + "_" + word)
                output.append(word + "_" + narratorName)
                missingAudios.append(word)

                if word not in combinedMissingWords:
                    combinedMissingWords.append(word)

        except Warning:
            print "warning thrown on word: " + word
            warningcount += 1
            print warningcount

    out.write("\nWords: \n")
    for word in bookText:
        out.write(word + "\n")
    out.write("\nAudio File Names: \n")
    for word in output:
        out.write(word + "\n")
    out.write("\nMissing Words: \n")
    for word in missingAudios:
        out.write(word + "\n")
    out.write("\nNeed Human Verification: \n")
    for word in humanCheck:
        out.write(word + "\n")

    out.close()

def GetCellValue(sheet, row, col):
    cellType = sheet.cell_type(row, col)
    print "cellType: ", cellType

    if cellType == 2:                               #"XL_CELL_NUMBER": float
        print "cell value: ", sheet.cell_value(row, col)
        return str(int(sheet.cell_value(row, col)))
    elif cellType == 4:                             #"XL_CELL_BOOLEAN": int 1 = true, 0 = false
        value = sheet.cell_value(row, col)
        print "value: ", value
        if value == 1:
            return "true"
        else:
            return "false"
    else:
        return sheet.cell_value(row, col)

def Normalize(word):
    try:
        replaceWord = unicodedata.normalize('NFKD', word).encode('ascii', 'replace') #replaces the unicode chars with ?
        return replaceWord
    except Exception as e:
        # word has to be unicode to normalize. If string it throws an error
        return word

def ParseBookText(sheetName, startRow, startCol):
    # xlrd examples for parsing Excel sheets:
    # 
    # print "The number of worksheets is", book.nsheets
    # print "Worksheet name(s):", book.sheet_names()
    # sh = book.sheet_by_index(0)
    # sh = book.sheet_by_name("58 - The Sun and the North Wind")    
    # print sh.name, sh.nrows, sh.ncols
    # print "Cell D30 is", sh.cell_value(rowx=29, colx=3)
    # print "Cell B14 is", sh.cell_value(rowx=16, colx=1)
    # for rx in range(sh.nrows):
    #     print sh.row(rx)

    book = xlrd.open_workbook(pathToBookText)
    sh = book.sheet_by_name(sheetName)
    pageText = GetCellValue(sh, startRow, startCol)
    exclude = set(string.punctuation)
    exclude.remove("'") #remove punctuation from book text except "'"

    while(pageText != ""):
        pageText = ''.join(char.lower() for char in pageText if char not in exclude)
        # print "pageText: ", pageText

        for word in pageText.split():
            # replaceWord = unicodedata.normalize('NFKD', word).encode('ascii', 'replace') #replaces the unicode chars with ?
            # replaceWord = word
            replaceWord = Normalize(word)

            if u'\u2019' in word: #checks for unicode char for "'"
                if replaceWord[0] == '?':
                    replaceWord = replaceWord[1:]
                replaceWord = replaceWord.replace('?', '\'')

            if '?' in replaceWord: #if ? is still in replaceWord it's for a unicode char I don't want
                replaceWord = replaceWord.replace('?', '')

            if replaceWord not in bookText:
                bookText.append(replaceWord)
        startRow += 1
        if startRow < sh.nrows:
            pageText = GetCellValue(sh, startRow, startCol)
        else:
            #no rows left so will throw an index out of range exception
            break

def LoadAudiosWeHave():
    narratorName = ""

    for (dirpath, dirnames, filenames) in walk(pathToAudioWeHave):
        for i in range(0, len(filenames)):
            if filenames[i].endswith('.wav') or filenames[i].endswith('.wma'):
                if narratorName == "":
                    #is the narrator name first or second? "Alex_dad" or "dad_Alex"
                    narratorName = filenames[i].split('_')[-1]
                # key = "_".join(filenames[i].split('_')[1:]) #split on '_' to remove 'Alex' then join the rest (narrator name first)
                key = "_".join(filenames[i].split('_')[:-1]) #get all but the last piece for when narrator name is last
                # key = key[:-4] #remove .wav or .wma
                key = ''.join(char.lower() for char in key)
                value = filenames[i][:-4]
                wordToAudioName[key] = value
        print "wordToAudioName: ", wordToAudioName
        return narratorName
        #break

def ParseStartLocationFromSheetName(sheetName):
    coord = sheetName.split()[-1] # "B14" is the last item in space seperated list
    sheetName = " ".join(sheetName.split()[:-1]) #recombining sheet name without coordinate
    asciiValForCapitalA = 65
    startCol = ord(coord[0]) - asciiValForCapitalA #ascii val of letter they pass in minus val of A
    startRow = int(coord[1:])-1

    print "sheetName: " + sheetName
    print "startRow: " + str(startRow)
    print "startCol: " + str(startCol)

    return sheetName, startRow, startCol

if __name__ == "__main__":
    pathToBookText = "" #"C:/Users/adam.christiansen.IMAGINELEARNING/Downloads/Start Reading English - All 3 Modes (decodables) LUT.xlsx"
    pathToAudioWeHave = "" #"O:/Projects/IL - Activities/[Shared Media]/Level 2/Shared Word Rec and Reading/Audio/Narration/English - Alex"
    pathToOutputFile = "" #"C:/Users/adam.christiansen.IMAGINELEARNING/Desktop/"#CheckBookAudioOut.txt"
    sheetNamesToParse = []
    combinedMissingWords = []
    sheetNames = []
    wordToAudioName = {}
    startRow = 0
    startCol = 0

    pathToAudioWeHave = diropenbox("msg", "Choose the folder with your existing audio", 
        default="O:/Projects/IL - Activities/[Shared Media]/Level 2/Shared Word Rec and Reading/Audio/Narration/English - Alex")
    pathToBookText = fileopenbox("msg", "Select your LUT file")
    pathToOutputFile = diropenbox("msg", "Select the folder where you want to save your output")
    pathToOutputFile = pathToOutputFile + "\\"

    sheetNamesToParse = ParseArguments()
    # pathToBookText, pathToAudioWeHave, pathToOutputFile, sheetNamesToParse = ParseArguments()
    # pathToBookText = "//il2tfs/sites/DefaultCollection/ILE/ilactivitydocs/Specs/Look Up Tables/Start Reading English - All 3 Modes (decodables) LUT.xlsx"
    narratorName = LoadAudiosWeHave()

    for sheetName in sheetNamesToParse:
        bookText = []
        sheetName, startRow, startCol = ParseStartLocationFromSheetName(sheetName)
        ParseBookText(sheetName, startRow, startCol)
        CompareAudioAndText(sheetName)

    combinedMissingWords.sort()
    combinedOutFile = open(pathToOutputFile + "CombinedMissingWords.txt", "w")
    combinedOutFile.write("Combined list of missing words: \n")
    for word in combinedMissingWords:
        combinedOutFile.write(word + '\n')