import re
import os

def getExcelFileName():
    """Search *.xlsx (*.XLSX) files in current folder and return its filename."""

    pattern = re.compile(".*?xlsx$", re.I)

    for fileName in os.listdir():
        match = re.match(pattern, fileName)
        if match != None:
            return match.group(0)
            
    return "None"
