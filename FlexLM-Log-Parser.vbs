'FlexLM-Log-Parser
'Author: Sukhjit Singh
'Github: https://github.com/ImSukh/FlexLM-Log-Parser

'___________________
'Functions
'-------------------

'Build an array list by searching for unique values through array 
'   Passing in x = 3 generates a list of unique products
'   Passing in x = 4 generates a list of unique users
'   This is because when the data from relevant lines in the debug file are split
'   the fourth value in the array is the product and the fifth is the user id
Function BuildList(ByRef Arr(), ByVal x)
    count = 0 
    Redim PRESERVE Arr(count)
    For i = 0 to UBound(FileArr)
        strA = Trim(FileArr(i))'remove leading whitespaces

        'Only check lines which contain information about when the license was checked in or out
        If InStr(strA,"IN:") <> 0 or InStr(strA,"IN:") <> Null or InStr(strA,"OUT:") <> 0 or InStr(strA,"OUT:") <> Null Then
            LogData = Split(strA) 'split the string

            'Check if the item exists in the array
            If FindinArray(LogData(x), Arr) = False Then
                'Increase the size of the array by 1 and add the new value
                Redim PRESERVE Arr(count)
                Arr(count) = LogData(x)
                count = count + 1
            End If
        End If
    Next
End Function

'See if the item already exists in an array 
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

'declare file steam objects
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile("debug.log", 1) 'debug file to be used
Set objCSV = objFSO.CreateTextFile("AutoDeskLog.csv", True) 'name of csv file to be generated

'Initialize array objects and other variables
Dim Products(), Users(), FileArr()

count = 0
TimeFound = False 'boolean value to determine start time of debug file. This only needed to establish the start time for the first license that is checked out

'Read all relevant lines from debug file into an array 
'This will cut down on execution time by only looking at line data that relates to 
'when a license was checked in or out
Do Until objFile.AtEndofStream
    strA = Trim(objFile.Readline)
    If InStr(strA,"IN:") <> 0 or InStr(strA,"IN:") <> Null or InStr(strA,"OUT:") <> 0 or InStr(strA,"OUT:") <> Null or InStr(stra,"TIMESTAMP") <> 0 or InStr(strA,"TIMESTAMP") <> Null Then
        'increase size of array to accomdate for relevant data
        Redim PRESERVE FileArr(count)
        FileArr(Count) = strA  
        count = count + 1
    End If

    'Set the start time
    If (InStr(stra,"TIMESTAMP") <> 0 or InStr(strA,"TIMESTAMP") <> Null) AND Timefound = False Then
        strA = Split(strA)
        StartTime = strA(3)
        TimeFound = True
    End If
Loop

Call BuildList(Products,3) 'Create a list of Unique Products
Call BuildList(Users,4) 'Create a list of Unique Users

'Write column names for CSV file
objCSV.Write "From Date" &","& "To Date" &","& "Product" &","& "User" &","& "Time in seconds" & vbNewLine

'Iterate over every product found
For i = 0 to UBound(Products)
    'Iterate for every user found
    For x = 0 to UBound(Users)
        LastOut = 0 'the point we should start looking from when trying to create License Out and In pairs
        TotalTime = 0 'Total usage time for said product by said user
        
        'Set initial time stamp
        TimeStampOUT = StartTime 
        TimeStampIN = StartTime 

        'Iterate over every item in FileArr
        For y = 0 to UBound(FileArr) 
            
            'Trim leading whitespace from FileArr item and the split it on whitespace
            strA = Trim(FileArr(y))
            LogData = Split(strA)

            'Search for a license out that has matching Product and User info
            If (InStr(strA,"OUT:") <> 0 or InStr(strA,"OUT:") <> Null) AND (InStr(strA,Users(x)) <> 0 or InStr(strA,Users(x)) <> Null) AND (InStr(strA,Products(i)) <> 0 or InStr(strA,Products(i)) <> Null) Then
                
                fromDate = TimeStampOUT + " " + Logdata(0)

                'search for a license in that matches the Product and User info and occurs after the License Out
                For z = LastOut to UBound(FileArr)
                    strA = Trim(FileArr(z))
                    LogData = Split(strA)
                    If (InStr(strA,"IN:") <> 0 or InStr(strA,"IN:") <> Null) AND (InStr(strA,Users(x)) <> 0 or InStr(strA,Users(x)) <> Null) AND (InStr(strA,Products(i)) <> 0 or InStr(strA,Products(i)) <> Null) Then
                        
                        toDate=TimeStampIN + " " + Logdata(0)

                        'Uncomment line below for cscript or wscript output. Do not uncomment if running regularly otherwise you will have alot of textboxes to close
                        'wscript.echo "From Date:",fromDate,"To Date:", toDate, "Product:",Products(i),"User:", Users(x), DateDiff("s",fromDate,toDate) , " seconds"
                        
                        'Calculate license usage duration and write to csv file
                        objCSV.Write fromDate &","& toDate &","& ProductName(Products(i)) &","& Users(x) &","& DateDiff("s",fromDate, toDate) &","& vbNewLine
                        LastOut = z + 1
                        TotalTime = TotalTime + DateDiff("s",fromDate,toDate)
                        Exit For
                    End If

                    'Update Timestamp if found
                    If InStr(strA,"TIMESTAMP") <> 0 or InStr(strA,"TIMESTAMP") <> Null Then
                        LogData = Split(strA)
                        TimeStampIN = LogData(3)
                    End If
                Next 
            End If

            'Update TimeStamp if found
            If InStr(strA,"TIMESTAMP") <> 0 or InStr(strA,"TIMESTAMP") <> Null Then
                LogData = Split(strA)
                TimeStampOUT = LogData(3)
                TimeStampIN = LogData(3)
            End If
        Next
        'wscript.echo "Product:",Products(i),"User:", Users(x), "Total Time Usage:", Int(TotalTime/3600), "hours",Int((TotalTime-(Int(TotalTime/3600)* 3600))/60), "minutes"
        If TotalTime > 0 Then
            objCSV.Write ",,,Total Usage Time:," & Int(TotalTime/3600) & " hrs " & Int((TotalTime-(Int(TotalTime/3600)* 3600))/60) & " mins," & vbNewLine
            objCSV.Write vbNewLine
        End If
    Next
Next
objFile.Close
objCSV.Close