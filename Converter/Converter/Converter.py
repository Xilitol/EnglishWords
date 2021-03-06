import docx, xlsxwriter
import os, shutil, sys
from datetime import datetime
import re

maxWordsInFile = 300

class Paragraph:
    def __init__(self, nname, wwords) -> None:
        self.name = nname
        self.words = wwords
    pass

class Word:
    def __init__(self, enVal, ruVal, examp) -> None:
        self.enValue = enVal
        self.ruValue = ruVal
        self.example = examp
    pass

def updOneVal(val):
    if val[-1] == ' ':
        val = val[:-1]
    if val[0] == ' ':
        val = val[1:]

    slash = '/'
    goodslash = '/ '
    testVal = val.split()
    if len(testVal) > 1:
       for i in range(0,len(testVal)):
           if len(testVal[i]) > 15 and slash in testVal[i]:
              testVal[i] = testVal[i].replace(slash, goodslash)

    outVal = ' '.join([str(item) for item in testVal])
    return outVal

def updateRuAndEnVal(ruVal, enVal):
    ruVal = updOneVal(ruVal)
    enVal = updOneVal(enVal)
    ruVal = ruVal.capitalize()

    return ruVal, enVal
    
def getText(filename):
    doc = docx.Document(filename) 
    data = []
    wordsList = []

    usePar = False

    for para in doc.paragraphs:
        if len(para.text) == 0:
            continue
        cleanLine = re.sub('[^A-Za-z0-9]+', '', para.text)
        if len(cleanLine) == 0:
           continue
        if para.style.name == 'Heading 1':
            if usePar:
                wordsList.append(Word(enVal, ruVal, examp));
                data.append(Paragraph(parName, wordsList.copy()))
                wordsList.clear()
            
            enVal=ruVal=""    
            
            if para.text[0] != '#':
                parName = para.text
                usePar = True
            else:
                parName = ""
                usePar = False
            continue
        if para.style.name == 'Heading 2':
            if not usePar:
                continue
            if enVal and ruVal:
                wordsList.append(Word(enVal, ruVal, examp));
            if para.text == '\n':
                continue
            try:
                enVal, ruVal = para.text.split('\u2013')
            except Exception:
                print("Error splitting line:")
                print(para.text)
                exit()
            ruVal, enVal = updateRuAndEnVal(ruVal, enVal)
            examp = ""
            examp_number = 1
            continue
        if para.style.name == 'Normal':
            if not usePar:
                continue
            if examp_number != 1:
                examp += "\n"
            examp += str(examp_number) + ". " + para.text
            examp_number += 1
        else:
            continue

    if usePar:
         wordsList.append(Word(enVal, ruVal, examp));
         data.append(Paragraph(parName, wordsList.copy()))
         wordsList.clear()

    return data

# ru | en | part of speech | hint (example)
def writeText(data, outputDir, inputFileName, inputFile):

    current_datetime = datetime.now()
    current_datetime_text = "_" + str(current_datetime.year) +"_" + str(current_datetime.month) + "_" + str(current_datetime.day) + "_" + str(current_datetime.hour) + "_" + str(current_datetime.minute) + "_" + str(current_datetime.second)
    orig_name, end = inputFileName.split('.');

    outputDirName = outputDir + orig_name +  "_[Output]" + current_datetime_text;

    os.mkdir(outputDirName)

    for paragraph in data:
        n = int(len(paragraph.words) / maxWordsInFile)
        for i in range(0, n + 1):
            if n == 0:
                outputFile = outputDirName + "/" + paragraph.name + ".xlsx";
            else:
                outputFile = outputDirName + "/" + paragraph.name + "Part" + str(i + 1) + ".xlsx";
            workbook = xlsxwriter.Workbook(outputFile)
            worksheet = workbook.add_worksheet()

            row = 0
            for j in range(i * maxWordsInFile, (i + 1) * maxWordsInFile):
                if j == len(paragraph.words):
                    break
                worksheet.write(row, 0, paragraph.words[j].ruValue)
                worksheet.write(row, 1, paragraph.words[j].enValue)
                worksheet.write(row, 3, paragraph.words[j].example)
 
                row += 1

            workbook.close()
    shutil.copy(inputFile, outputDirName)
  
if len(sys.argv) < 2:
    print("Arguments error!")
elif len(sys.argv) == 2:
    inputDir = sys.path[0] + "\\"
    outputDir = sys.path[0] + "\\"
    inputFileName = sys.argv[1]
else:
   inputDir = sys.argv[1]
   inputFileName = sys.argv[2]
   if len(sys.argv) > 3:
       outputDir = sys.argv[3]
   else:
       outputDir = inputDir

inputFile = inputDir + inputFileName
data = getText(inputFile)

writeText(data, outputDir, inputFileName, inputFile)

print("Success!")