from datetime import datetime

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

def IfValue(row):
                        
    for i in range (0, len(row)):
        if(type(row[i]) != str):
            check = 1
            break
        else:
            check = 2
            
    if(check == 1):
        return True

    else:
        return False    

def DATETIME(date):

    date = date.replace("(", "")
    date = date.replace(")", "")
    date = date + " 00:00:00"
    
    month = date[3] + date[4] + date[5]
    monthInt = ""
    dateFormat = "%d-%m-%Y %H:%M:%S"                            
    
    if(month == "Jan"):
        monthInt = "01"
        date = date.replace(month,monthInt)
    
    elif(month == "Feb"):
        monthInt = "02"
        date = date.replace(month,monthInt)
        
    elif(month == "Mar"):
        monthInt = "03"
        date = date.replace(month,monthInt)
        
    elif(month == "Apr"):
        monthInt = "04"
        date = date.replace(month,monthInt)
        
    elif(month == "May"):
        monthInt = "05"
        date = date.replace(month,monthInt)
        
    elif(month == "Jun"):
        monthInt = "06"
        date = date.replace(month,monthInt)
        
    elif(month == "Jul"):
        monthInt = "07"
        date = date.replace(month,monthInt)
        
    elif(month == "Aug"):
        monthInt = "08"
        date = date.replace(month,monthInt)
        
    elif(month == "Sep"):
        monthInt = "09"
        date = date.replace(month,monthInt)
        
    elif(month == "Oct"):
        monthInt = "10"
        date = date.replace(month,monthInt)
        
    elif(month == "Nov"):
        monthInt = "11"
        date = date.replace(month,monthInt)
        
    elif(month == "Dec"):
        monthInt = "12"
        date = date.replace(month,monthInt)
    
    date = datetime.strptime(date, dateFormat)
    return(date) 