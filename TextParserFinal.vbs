Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile("debug.log", 1)
Set objCSV = objFSO.CreateTextFile("AutoDeskLog.csv", True)

Dim Products(), Users(), FileArr()

count = 0
TimeFound = False 
Do Until objFile.AtEndofStream
    strA = Trim(objFile.Readline)
    If InStr(strA,"IN:") <> 0 or InStr(strA,"IN:") <> Null or InStr(strA,"OUT:") <> 0 or InStr(strA,"OUT:") <> Null or InStr(stra,"TIMESTAMP") <> 0 or InStr(strA,"TIMESTAMP") <> Null Then
        Redim PRESERVE FileArr(count)
        FileArr(Count) = strA  
        count = count + 1
    End If
    If (InStr(stra,"TIMESTAMP") <> 0 or InStr(strA,"TIMESTAMP") <> Null) AND Timefound = False Then
        strA = Split(strA)
        StartTime = strA(3)
        TimeFound = True
    End If
Loop

Call BuildList(Products,3)
Call BuildList(Users,4)

'Write column names for CSV file
objCSV.Write "From Date" &","& "To Date" &","& "Product" &","& "User" &","& "Time in seconds" & vbNewLine


For i = 0 to UBound(Products)
    For x = 0 to UBound(Users)
        LastOut = 0
        TotalTime = 0
        TimeStampOUT = StartTime
        TimeStampIN = StartTime
        For y = 0 to UBound(FileArr)
            strA = Trim(FileArr(y))
            LogData = Split(strA)
            If (InStr(strA,"OUT:") <> 0 or InStr(strA,"OUT:") <> Null) AND (InStr(strA,Users(x)) <> 0 or InStr(strA,Users(x)) <> Null) AND (InStr(strA,Products(i)) <> 0 or InStr(strA,Products(i)) <> Null) Then
                fromDate = TimeStampOUT + " " + Logdata(0)
                For z = LastOut to UBound(FileArr)
                    strA = Trim(FileArr(z))
                    LogData = Split(strA)
                    If (InStr(strA,"IN:") <> 0 or InStr(strA,"IN:") <> Null) AND (InStr(strA,Users(x)) <> 0 or InStr(strA,Users(x)) <> Null) AND (InStr(strA,Products(i)) <> 0 or InStr(strA,Products(i)) <> Null) Then
                        toDate=TimeStampIN + " " + Logdata(0)
                        'wscript.echo "From Date:",fromDate,"To Date:", toDate, "Product:",Products(i),"User:", Users(x), DateDiff("s",fromDate,toDate) , " seconds"
                        If DateDiff("s",fromDate, toDate) < 0 Then
                            TimeStampIN = CStr(DateAdd("d",1,TimeStampIN))
                            toDate=TimeStampIN + " " + Logdata(0)
                        End If
                        objCSV.Write fromDate &","& toDate &","& ProductName(Products(i)) &","& Users(x) &","& DateDiff("s",fromDate, toDate) &","& vbNewLine
                        LastOut = z + 1
                        TotalTime = TotalTime + DateDiff("s",fromDate,toDate)
                        Exit For
                    End If
                    If InStr(strA,"TIMESTAMP") <> 0 or InStr(strA,"TIMESTAMP") <> Null Then
                        LogData = Split(strA)
                        TimeStampIN = LogData(3)
                    End If
                Next 
            End If
            If InStr(strA,"TIMESTAMP") <> 0 or InStr(strA,"TIMESTAMP") <> Null Then
                LogData = Split(strA)
                TimeStampOUT = LogData(3)
                TimeStampIN = LogData(3)
            End If
        Next
        'wscript.echo "Product:",Products(i),"User:", Users(x), "Total Time Usage:", Int(TotalTime/3600), "hours",Int((TotalTime-(Int(TotalTime/3600)* 3600))/60), "minutes"
    Next
Next
objFile.Close
objCSV.Close

'Build a list by searching for unique values through array 
'The x value represents Product if it is = 3, and User if it is 4
Function BuildList(ByRef Arr(), ByVal x)
    count = 0 
    Redim PRESERVE Arr(count)
    For i = 0 to UBound(FileArr)
        strA = Trim(FileArr(i))'remove leading whitespaces

        'If the string contains an 'IN:' or 'OUT:' value
        If InStr(strA,"IN:") <> 0 or InStr(strA,"IN:") <> Null or InStr(strA,"OUT:") <> 0 or InStr(strA,"OUT:") <> Null Then
            LogData = Split(strA) 'split the string

            '
            If FindinArray(LogData(x), Arr) = False Then
                Redim PRESERVE Arr(count)
                Arr(count) = LogData(x)
                count = count + 1
            End If
        End If
    Next
    
End Function

'See if the item already exsists in the array 
Function FindinArray(ByRef Name, ByRef Arr())
    For i = 0 to UBound(Arr)
        If Name = Arr(i) Then 
            FindinArray = True
            Exit Function
        End If
    Next
    FindinArray = False
End Function

'Get Product Name from product id 
Function ProductName(ByVal Name)
    Name = Replace(Name, """","")
    If Name = "87091MAP_2019_0F" Then
        ProductName = "Autodesk AutoCAD Map 3D 2019"
    End If
    
    If Name = "86893CIV3D_2018_0F" Then
        ProductName = "Autodesk AutoCAD Civil 3D 2018"
    End If
    
    If Name = "86815AECCOL_T_F" Then
        ProductName = "AEC Collection Package"
    End If

    If Name = "86718CIV3D_2017_0F" Then
        ProductName = "Autodesk AutoCAD Civil 3D 2017"
    End If

    If Name = "64300ACD_F" Then
        ProductName = "AutoCAD Package"
    End If

    If Name = "86604ACD_2017_0F" Then
        ProductName = "Autodesk AutoCAD 2017"
    End If

    If Name = "87140CIV3D_2019_0F" Then
        ProductName = "Autodesk Civil 3D 2019"
    End If

    If Name = "86719MAP_2017_0F" Then
        ProductName = "Autodesk AutoCAD Map 3D 2017"
    End If

    If Name = "866333DSMAX_2017_0F" Then
        ProductName = "Autodesk 3ds Max 2017"
    End If

    If Name = "86606ARDES_2017_0F" Then
        ProductName = "Autodesk AutoCAD Raster Design 2017"
    End If
End Function