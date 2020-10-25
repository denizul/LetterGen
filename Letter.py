#!/usr/bin/python
#
# Requires python3!
# A script that will automate the personalisation of cover letters
# This script only works with ODT files, so you need LibreOffice!
# This script will unzip, edit and zip back ODT documents in python
# Source ODT file "in.odt" shall exist in "/tmp"
# If ODT file contains string token, it will be replaced with string replacement
#

import os
import zipfile
import fileinput
import sys
import datetime

# This is the directory of the Cover letter!
rootDirURL="/home/ziad/Documents/Work/Study"

tmpDir="~odt_contents"
tmpDirURL=rootDirURL+"/"+tmpDir

zipSourceFile="Cover Letter.odt" # Put your cover letter file here!
zipSourceFileURL=rootDirURL+"/"+zipSourceFile

zipOutFile="NEWout.odt" 
zipOutFileURL=rootDirURL+"/"+zipOutFile

xmlFile="content.xml"
xmlFileURL=tmpDirURL+"/"+xmlFile

token01="COMPANY01"
replacement01=(str(sys.argv[1]))

token02="POSITION02"
replacement02=(str(sys.argv[2]) + " " + str(sys.argv[3]))

token03="DATE03"
myDateVer=datetime.datetime.now()
replacement03=str(str(myDateVer.year) + "/"+ str(myDateVer.month) +"/"+ str(myDateVer.day))

command="libreoffice  --headless --convert-to pdf NEWout.odt"
command02="cp NEWout.pdf " + replacement01 + "-CoverLetter.pdf"

#
# Unzip ODT
#
print (" -- Extracting ---------------------")
print ("%s -> %s" % (zipSourceFileURL, tmpDirURL))

zipdata = zipfile.ZipFile(zipSourceFileURL)
zipdata.extractall(tmpDirURL)

#
# Find and replace tokens
#
print (" -- Replacing -------------")
print (xmlFileURL)

for line in fileinput.input(xmlFileURL, inplace=1):
    print (line.replace(token01,replacement01))

for line in fileinput.input(xmlFileURL, inplace=1):
    print (line.replace(token02,replacement02))
    

for line in fileinput.input(xmlFileURL, inplace=1):
    print (line.replace(token03,replacement03))

# Zip contents of the temporary directory to ODT
# Use file list from the original archive
# This preserves the file structure in the new Zip file
# The most important is that the "mimetype" is the first file in archive

print (" -- Compressing --------------------")
print ("%s -> %s" % (tmpDirURL , zipOutFileURL))


with zipfile.ZipFile(zipOutFileURL, 'w') as outzip:
    zipinfos = zipdata.infolist()
    for zipinfo in zipinfos:
        fileName=zipinfo.filename # The name and path as stored in the archive
        fileURL=tmpDirURL+"/"+fileName # The actual name and path
        outzip.write(fileURL,fileName)



os.system(command)
os.system(command02)