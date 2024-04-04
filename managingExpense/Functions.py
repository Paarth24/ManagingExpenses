def AddingFileExt(fileName):
    fileNameExt = fileName + ".xlsx"
    return(fileNameExt)

def GetExcelFileName(path):
    count = 0
    for i in range(0, len(path)):
        if(path[i] == "/"):
           count = count + 1

    file = ""
    for i in range(0, len(path)):
        if(path[i] == "/"):
           count = count - 1


        if(count == -1):
            file = file + path[i]

        if(count == 0):
            count = -1
            
    return(file)


