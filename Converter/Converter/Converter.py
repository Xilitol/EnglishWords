import docx
import xlsxwriter
from datetime import datetime

inputFile = "../../../../English/English words.docx"
outputDir = "../../../../English/MemoWordApp/"

class Word:
    def __init__(self, enVal, ruVal, examp) -> None:
        self.enValue = enVal
        self.ruValue = ruVal
        self.example = examp
    pass
    
def getText(filename):
    doc = docx.Document(filename) 
    data = []

    for para in doc.paragraphs:
        if para.style.name == 'Heading 1':
            if para.text == "Vocabulary":
                enVal=ruVal = ""
                continue
            else:
                data.append(Word(enVal, ruVal, examp));
                break

        if para.style.name == 'Normal':
            if enVal and ruVal:
                data.append(Word(enVal, ruVal, examp));
            if para.text == '\n':
                break
            enVal, ruVal = para.text.split('\u2013')
            if enVal[-1] == ' ':
                enVal = enVal[:-1]
            if ruVal[0] == ' ':
                ruVal = ruVal[1:]
            ruVal = ruVal.capitalize()
            examp = ""
            continue

        if para.style.name == 'No Spacing':
            examp += para.text + '\n'

        else:
            break
        
    return data

# ru | en | part of speech | hint (example)
def writeText(data, outputDir):

    current_datetime = datetime.now()

    outputFile = outputDir + "Output"+"_" + str(current_datetime.year) +"_" + str(current_datetime.month) + "_" + str(current_datetime.day) + "_" + str(current_datetime.hour) + "_" + str(current_datetime.minute) + "_" + str(current_datetime.second) + ".xlsx";
    workbook = xlsxwriter.Workbook(outputFile)
    worksheet = workbook.add_worksheet()
 
    row = 0 

    # here is a magic constant: Memo Word doesn't allow files more than 300 lines
    # must be updated
    for i in range(301, 600) :
        worksheet.write(row, 0, data[i].ruValue)
        worksheet.write(row, 1, data[i].enValue)
        worksheet.write(row, 3, data[i].example)
 
        row += 1
     
    workbook.close()
    

data = getText(inputFile)

writeText(data, outputDir)

print("Success!")
#for word in data:
#    print(word.enValue + ":" + word.ruValue + "\n" + word.example + "\n\n\n")