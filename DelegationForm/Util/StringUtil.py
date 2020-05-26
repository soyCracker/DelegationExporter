import os

def ClearSpace(str):
    if str!=None:
        return str.replace('\u200b', "")
    else:
        return str
