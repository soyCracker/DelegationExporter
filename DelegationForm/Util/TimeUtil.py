import datetime

def GetCurrentTime():
    return datetime.datetime.now().strftime("%Y%m%d%H%M%S")
