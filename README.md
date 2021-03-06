# EnglishWords
A small app converting the DB from.docx file to xlsx file for MemoWord App

Memoword App: https://memoword.online/en/

This script converts .docx file to .xlsx file formatted for Memoword App

## Input file structure
You can find the example file (TheMakingOfTheMob.docx) in the repository

### Name of section
Style = Heading 1. First symbol = '#' to ignore this section

In one section you can have any amount of words

If you use more than 300 words in one section, the section will be splitted to several .xlsx files by 300 words, because one set in Memoword doesn't allow more than 300 words

### Words
Style = Heading 2

Symbol '–' (U+2013) splits the word in original language and the translation

### Hints (I add here examples of using)
Style = default

Any amount of paragraphs. 

## Output file:
The directory is created, its name contains "Output", input file name, date and time

In this directory it will be created a list of .xlsx files, named: section name and "PartX", if needed

Also this script copies the input file to the output directory

### Each .xlsx file
1st colomn: original words

2nd colomn: translation

3rd colomn: -

4th colomn: the hint (examples)

## Input arguments:
1st way: 1 argument (filename). Output directory will be created in script's directory

2nd way: 2 arguments: inputDirectory filename. Output directory will be created in inputDirectory

3rd way: 3 arguments: inputDirectory filename outputDirectory. Output directory will be created in outputDirectory

## Example

```
Converter.py "C:\Data\English\WordsLists\" "TheMakingOfTheMob.docx" "C:\Data\English\MemoWordApp\Movies\"
```

To use you need Python3 and install packages: 
```
pip install python-docx xlsxwriter
```
