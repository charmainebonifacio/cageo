Attribute VB_Name = "CCDST_4_DataProcessAndMerge"
'---------------------------------------------------------------------
' Date Created : March 17, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : April 20, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : ProcessMergeData
' Description  : This function cleans each .CSV file within a user
'                specified folder directory. Function will delete
'                unnecesary information and then it will replace
'                existing column headers (for efficiency purposes).
'                Any trailing white space will be removed. All text
'                will be converted to uppercase format. All changes
'                will be saved. Function will return how long the
'                procedure took.
' Parameters   : String, String, String, Integer
' Returns      : Long
'---------------------------------------------------------------------
Function ProcessMergeData(ByVal ClimateStation As String, ByVal DestFolder As String, _
ByVal MasterDestinationFolder As String, ByVal UserDataInterval As Integer) As Boolean
    
    Dim wbDest As Workbook, wbSource As Workbook
    Dim DestSheet As Worksheet, SourceSheet As Worksheet
    Dim sThisFilePath As String, sFile As String
    Dim FileCount As Integer
    Dim start_time As Date, end_time As Date
    Dim CleanTime As Long
    Dim TimeElapsed As String
    Dim SubFolderPath As String

    ' Disable all the pop-up menus
    Application.ScreenUpdating = False
    
    ' Initialize Variables
    FileCount = 0
    
    '-------------------------------------------------------------
    ' Add New WorkBook with one worksheet.
    ' Set it as the destination worksheet.
    '-------------------------------------------------------------
    Set wbDest = Workbooks.Add(1)
    Set DestSheet = wbDest.Worksheets(1)
    ActiveSheet.Name = "Data"
    
    '-------------------------------------------------------------
    ' Select .csv file only. This will contain
    ' the filepath and filename in one string.
    '-------------------------------------------------------------
    sThisFilePath = DestFolder
    If (Right(sThisFilePath, 1) <> "\") Then sThisFilePath = sThisFilePath & "\"
    sFile = Dir(sThisFilePath & "*.csv*") ' Only .csv files
    ChDir (sThisFilePath)

    '-------------------------------------------------------------
    ' Loop all the .csv files in the current path.
    '-------------------------------------------------------------
    Debug.Print "First, delete unnecessary headings found in all Environment Canada RAW file."
    Debug.Print "Replace column header for each file in the directory."
    Debug.Print "Last, trim any white space and convert all column headers to uppercase."

    Do While sFile <> vbNullString
        ' File count of current folder
        FileCount = FileCount + 1
    
        ' Open file and set it as source worksheet
        Set wbSource = Workbooks.Open(sFile)
        Set SourceSheet = wbSource.Worksheets(1)
        SourceSheet.Activate
    
        '-------------------------------------------------------------
        ' Remove unnecessary headings found in all RAW file.
        ' Remove any white space and convert all text to UPPERCASE
        ' Assuming the ReplaceColumnHeader Function worked!
        '-------------------------------------------------------------
        Call RemoveTopRows(SourceSheet)
        Call FindAndDeleteColumns(SourceSheet)
        Call ColumnHeaderProcessing(SourceSheet, UserDataInterval)
        wbSource.Save
        
        '-------------------------------------------------------------
        ' Copy / Paste Data
        ' For First File - use header
        ' For the rest - keep header but only copy data
        '-------------------------------------------------------------
        Call AppendProcessedData(FileCount, SourceSheet, DestSheet)

        ' Ignore Clipboard Alerts
        Application.CutCopyMode = True
        
        ' Save Changes to the Processed Files
        wbSource.Close SaveChanges:=False
    
        sFile = Dir
    Loop
    
    ' Save Master File
    Call SaveCSV(wbDest, MasterDestinationFolder, ClimateStation)
    
    ProcessMergeData = True
    Application.StatusBar = False
    
EndHandler:
    '-------------------------------------------------------------
    ' Clean up memory.
    '-------------------------------------------------------------
    Set wbDest = Nothing
    Set wbSource = Nothing
    Set DestSheet = Nothing
    Set SourceSheet = Nothing
End Function
'---------------------------------------------------------------------
' Date Created : June 1, 2012
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : April 20, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : AppendProcessedData
' Description  : This function copies data from the source worksheet
'                and appends the data onto the destination worksheet.
' Parameters   : Integer, Worksheet, Worksheet
' Returns      : -
'---------------------------------------------------------------------
Function AppendProcessedData(ByVal Index As Integer, SourceSht As Worksheet, DestSht As Worksheet)
    
    Dim RngSelect
    Dim PasteSelect

    ' Activate Source Worksheet.
    SourceSht.Activate
    
    '-------------------------------------------------------------
    ' Process data. Retain the header information for the first
    ' file opened. Only copy data for the remaining files.
    '-------------------------------------------------------------
    If Index <> 1 Then
        ActiveSheet.Rows(1).Select
        Selection.Delete
    End If
    
    '-------------------------------------------------------------
    ' Call FindRange function to select the current used data
    ' within the Source Worksheet. Only copy the selected data.
    '-------------------------------------------------------------
    Call FindRange(SourceSht)
    RngSelect = Selection.Address
    Range(RngSelect).Copy

    ' Activate Destination Worksheet.
    DestSht.Activate
    
    '-------------------------------------------------------------
    ' Call RowCheck function to check the last row.
    ' Then append the copied data into the Destination Worksheet.
    '-------------------------------------------------------------
    Call RowCheck(DestSht)
    PasteSelect = Selection.Address
    Range(PasteSelect).Select
    DestSht.Paste
    Debug.Print "Paste_Successful. Index at " & Index
    
    ' Clear Clipboard of any copied data.
    Application.CutCopyMode = False
    
End Function
'---------------------------------------------------------------------
' Date Created : March 17, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : March 17, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : RemoveTopRows
' Description  : This function cleans data by removing the top
'                unnecessary rows.
' Parameters   : Worksheet
' Returns      : -
'---------------------------------------------------------------------
Function RemoveTopRows(ByVal SourceSht As Worksheet)
    
    Dim rowIndex As Integer
    Dim Found As Range
    Dim LastCell
    Dim StringToFind As String

    ' Activate Source Worksheet.
    SourceSht.Activate
    StringToFind = "Date/Time"
    '-------------------------------------------------------------
    ' Process data. Retain the header information for the first
    ' file opened. Only copy data for the remaining files.
    '-------------------------------------------------------------
    Range("A1").Activate
    Set Found = Cells.Find(What:=StringToFind, After:=ActiveCell)
    LastCell = Found.Address
    Range(LastCell).Select
    rowIndex = Selection.Row
    
    '-------------------------------------------------------------
    ' Process metadata. The first retains the header information.
    '-------------------------------------------------------------
    rowIndex = rowIndex - 1
        
    '-------------------------------------------------------------
    ' Delete metadata according to RowIndex.
    '-------------------------------------------------------------
    ActiveSheet.Rows("1:" & rowIndex).Select
    Selection.Delete
    
End Function
'---------------------------------------------------------------------
' Date Created : June 14, 2012
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : March 25, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : FindAndDeleteColumns
' Description  : This function will delete columns that contains
'                specific strings.
' Parameters   : Worksheet
' Returns      : -
'---------------------------------------------------------------------
Function FindAndDeleteColumns(WorksheetToEdit As Worksheet)

    Dim rng As Range, A As Range
    Dim r As Integer, c As Integer, IndexDel As Integer
    Dim StringToFind As String
    Dim Found As Range
    Dim Index As Integer
    Dim str As String
    Dim FoundCell

    WorksheetToEdit.Activate
    
    'Set Variable
    IndexDel = 1
    StringToFind = "Date/Time"
    
    Range("A1").Select
    Set Found = Range("A1")

    Set Found = Rows(1).Find(What:=StringToFind, After:=ActiveCell, LookIn:=xlValues, _
            LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False)
    With Found
        FoundCell = Found.Address
        Range(FoundCell).Select
        
        ' If header information is activated then exit the Do Loop
        str = ActiveCell.Value
        ActiveCell.Select ' Activate Cell
        ActiveCell.EntireColumn.Delete
    End With
End Function
'---------------------------------------------------------------------
' Date Created : June 14, 2012
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : March 25, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : ReplaceColumnHeader_Hourly
' Description  : This function will replace existing string to a new
'                header on an hourly climate data .CSV file.
' Parameters   : -
' Returns      : -
'---------------------------------------------------------------------
Function ReplaceColumnHeader_Hourly()

    Dim rng As Range, A As Range
    Dim r As Long, c As Long
    Dim StringToFind(), stringToReplace()
    Dim Found As Range
    Dim Index As Integer, IndexCount As Integer
    Dim str As String
    Dim FoundCell

    'Set Variable
    Set Found = Range("A1")
    IndexCount = 15
    
    ReDim StringToFind(1 To IndexCount)
    ReDim stringToReplace(1 To IndexCount)
    
    StringToFind(1) = "Temp"
    StringToFind(2) = "Dew Point"
    StringToFind(3) = "Wind "
    StringToFind(4) = "Rel "
    StringToFind(5) = "Stn Press"
    StringToFind(6) = "Visibility"
    StringToFind(7) = " "
    StringToFind(8) = "(°C)"
    StringToFind(9) = "(%)"
    StringToFind(10) = "(km)"
    StringToFind(11) = "(km/h)"
    StringToFind(12) = "(kPa)"
    StringToFind(13) = "(10s deg)"
    StringToFind(14) = "Flag"
    StringToFind(15) = "Quality"
    
    stringToReplace(1) = "TMP"
    stringToReplace(2) = "DEWP"
    stringToReplace(3) = "W"
    stringToReplace(4) = "R"
    stringToReplace(5) = "STNPRES"
    stringToReplace(6) = "VIS"
    stringToReplace(7) = "_"
    stringToReplace(8) = "C"
    stringToReplace(9) = "PER"
    stringToReplace(10) = "KM"
    stringToReplace(11) = "KMH"
    stringToReplace(12) = "KPA"
    stringToReplace(13) = "10SD"
    stringToReplace(14) = "F"
    stringToReplace(15) = "Q"
    
    For Index = 1 To IndexCount

        Set Found = Rows(1).Find(What:=StringToFind(Index), _
            After:=ActiveCell, LookIn:=xlValues, _
            LookAt:=xlPart, SearchOrder:=xlByColumns, _
            SearchDirection:=xlNext, MatchCase:=False)
    
        Do Until Found Is Nothing
            If Found Is Nothing Then
            Else
                With Found
                    FoundCell = Found.Address
                    Range(FoundCell).Select
                    str = ActiveCell.Value
                    ActiveCell.Value = Replace(str, StringToFind(Index), stringToReplace(Index))
                End With
            End If
            
            'Reintialize value for Do While Loop
             Set Found = Rows(1).Find(What:=StringToFind(Index), _
                After:=ActiveCell, LookIn:=xlValues, _
                LookAt:=xlPart, SearchOrder:=xlByColumns, _
                SearchDirection:=xlNext, MatchCase:=False)
        Loop
    Next Index
    
End Function
'---------------------------------------------------------------------
' Date Created : June 14, 2012
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : March 25, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : ReplaceColumnHeader_Daily
' Description  : This function will replace existing string to a new
'                header on an daily climate data .CSV file.
' Parameters   : -
' Returns      : -
'---------------------------------------------------------------------
Function ReplaceColumnHeader_Daily()

    Dim rng As Range, A As Range
    Dim r As Long, c As Long
    Dim StringToFind(), stringToReplace()
    Dim Found As Range
    Dim Index As Integer, IndexCount As Integer
    Dim str As String
    Dim FoundCell

    'Set Variable
    Set Found = Range("A1")
    IndexCount = 19
    
    ReDim StringToFind(1 To IndexCount)
    ReDim stringToReplace(1 To IndexCount)
    
    StringToFind(1) = "Max Temp"
    StringToFind(2) = "Min Temp"
    StringToFind(3) = "Mean Temp"
    StringToFind(4) = "Deg Days"
    StringToFind(5) = "Heat"
    StringToFind(6) = "Cool"
    StringToFind(7) = "Total"
    StringToFind(8) = "Snow"
    StringToFind(9) = "on Grnd"
    StringToFind(10) = "Dir of Max Gust"
    StringToFind(11) = "Spd of Max Gust"
    StringToFind(12) = " "
    StringToFind(13) = "(°C)"
    StringToFind(14) = "(mm)"
    StringToFind(15) = "(cm)"
    StringToFind(16) = "(km/h)"
    StringToFind(17) = "(10s deg)"
    StringToFind(18) = "Flag"
    StringToFind(19) = "Quality"
    
    stringToReplace(1) = "MAXTMP"
    stringToReplace(2) = "MINTMP"
    stringToReplace(3) = "MEANTMP"
    stringToReplace(4) = "DEGDAY"
    stringToReplace(5) = "HT"
    stringToReplace(6) = "CL"
    stringToReplace(7) = "TOT"
    stringToReplace(8) = "SNW"
    stringToReplace(9) = "GRND"
    stringToReplace(10) = "DIRMAXGUST"
    stringToReplace(11) = "SPDMAXGUST"
    stringToReplace(12) = "_"
    stringToReplace(13) = "C"
    stringToReplace(14) = "MM"
    stringToReplace(15) = "CM"
    stringToReplace(16) = "KMH"
    stringToReplace(17) = "10SD"
    stringToReplace(18) = "F"
    stringToReplace(19) = "Q"

    For Index = 1 To IndexCount

        Set Found = Rows(1).Find(What:=StringToFind(Index), _
            After:=ActiveCell, LookIn:=xlValues, _
            LookAt:=xlPart, SearchOrder:=xlByColumns, _
            SearchDirection:=xlNext, MatchCase:=False)
    
        Do Until Found Is Nothing
            If Found Is Nothing Then
            Else
                With Found
                    FoundCell = Found.Address
                    Range(FoundCell).Select
                    str = ActiveCell.Value
                    Debug.Print str
                    Debug.Print StringToFind(Index) & " " & stringToReplace(Index)
                    ActiveCell.Value = Replace(str, StringToFind(Index), stringToReplace(Index))
                End With
            End If
            
            'Reintialize value for Do While Loop
             Set Found = Rows(1).Find(What:=StringToFind(Index), _
                After:=ActiveCell, LookIn:=xlValues, _
                LookAt:=xlPart, SearchOrder:=xlByColumns, _
                SearchDirection:=xlNext, MatchCase:=False)
        Loop
    Next Index
    
End Function
'---------------------------------------------------------------------
' Date Created : March 17, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : March 25, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : ReplaceColumnHeader_Monthly
' Description  : This function will replace existing string to a new
'                header on an monthly climate data .CSV file.
' Parameters   : -
' Returns      : -
'---------------------------------------------------------------------
Function ReplaceColumnHeader_Monthly()

    Dim rng As Range, A As Range
    Dim r As Long, c As Long
    Dim StringToFind(), stringToReplace()
    Dim Found As Range
    Dim Index As Integer, IndexCount As Integer
    Dim str As String
    Dim FoundCell

    'Set Variable
    Set Found = Range("A1")
    IndexCount = 18
    
    ReDim StringToFind(1 To IndexCount)
    ReDim stringToReplace(1 To IndexCount)
    
    StringToFind(1) = "Max Temp"
    StringToFind(2) = "Min Temp"
    StringToFind(3) = "Mean Temp"
    StringToFind(4) = "Last Day"
    StringToFind(5) = "Total"
    StringToFind(6) = "Snow"
    StringToFind(7) = "on Grnd"
    StringToFind(8) = "Dir of"
    StringToFind(9) = "Spd of"
    StringToFind(10) = "Max Gust"
    StringToFind(11) = " "
    StringToFind(12) = "(°C)"
    StringToFind(13) = "(mm)"
    StringToFind(14) = "(cm)"
    StringToFind(15) = "(km/h)"
    StringToFind(16) = "(10s deg)"
    StringToFind(17) = "Flag"
    StringToFind(18) = "Quality"
    
    stringToReplace(1) = "MAXTMP"
    stringToReplace(2) = "MINTMP"
    stringToReplace(3) = "MEANTMP"
    stringToReplace(4) = "LASTDAY"
    stringToReplace(5) = "TOT"
    stringToReplace(6) = "SNW"
    stringToReplace(7) = "GRND"
    stringToReplace(8) = "DIR"
    stringToReplace(9) = "SPD"
    stringToReplace(10) = "MAXGUST"
    stringToReplace(11) = "_"
    stringToReplace(12) = "C"
    stringToReplace(13) = "MM"
    stringToReplace(14) = "CM"
    stringToReplace(15) = "KMH"
    stringToReplace(16) = "10SD"
    stringToReplace(17) = "F"
    stringToReplace(18) = "QLY"

    For Index = 1 To IndexCount

        Set Found = Rows(1).Find(What:=StringToFind(Index), _
            After:=ActiveCell, LookIn:=xlValues, _
            LookAt:=xlPart, SearchOrder:=xlByColumns, _
            SearchDirection:=xlNext, MatchCase:=False)
    
        Do Until Found Is Nothing
            If Found Is Nothing Then
            Else
                With Found
                    FoundCell = Found.Address
                    Range(FoundCell).Select
                    str = ActiveCell.Value
                    Debug.Print str
                    Debug.Print StringToFind(Index) & " " & stringToReplace(Index)
                    ActiveCell.Value = Replace(str, StringToFind(Index), stringToReplace(Index))
                End With
            End If
            
            'Reintialize value for Do While Loop
             Set Found = Rows(1).Find(What:=StringToFind(Index), _
                After:=ActiveCell, LookIn:=xlValues, _
                LookAt:=xlPart, SearchOrder:=xlByColumns, _
                SearchDirection:=xlNext, MatchCase:=False)
        Loop
    Next Index
    
End Function
'---------------------------------------------------------------------
' Date Created : June 25, 2012
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : March 25, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : ColumnHeaderProcessing
' Description  : Assumes the unnecessary rows have been deleted.
'                Only replaces thencolumn header's text, not data.
'                This function also converts all lowercase text to
'                UPPERCASE format and trims blank space.
' Parameters   : Worksheet, Integer
' Returns      : -
'---------------------------------------------------------------------
Function ColumnHeaderProcessing(WorksheetToEdit As Worksheet, ByVal UserDataInterval As Integer)

    Dim rACells As Range, rLoopCells As Range
    Dim lReply As Long    'Set variable to needed cells
    Dim rCellValue As String
    Dim LastColumnHeader As Long
    
    ' Activate Worksheet
    WorksheetToEdit.Activate
    
    Range("A1").Select
    LastColumnHeader = LastCol
    
    If WorksheetFunction.CountA(Cells) > 0 Then
        Range(Cells(1, 1), Cells(1, LastColumnHeader)).Select
    End If
    Set rACells = Selection
    
    On Error Resume Next 'In case of NO text constants.
    
    ' Set variable to all text constants
    Set rACells = rACells.SpecialCells(xlCellTypeConstants, xlTextValues)
    MsgBox rACells
        
    ' If could not find any text
    If rACells Is Nothing Then
        MsgBox "Could not find any text."
        On Error GoTo 0
        Exit Function
    End If
    
    '---------------------------------------------------------------------
    ' This section deals with the user data interval selection to
    ' replacing the appropriate data column headers."
    '   Integer 49 = Hourly Data
    '   Integer 50 = Daily Data
    '   Integer 51 = Monthly Data
    '---------------------------------------------------------------------
    Debug.Print "Then replace column headers."
    Select Case UserDataInterval
        Case 49
            Call ReplaceColumnHeader_Hourly
        Case 50
            Call ReplaceColumnHeader_Daily
        Case 51
            Call ReplaceColumnHeader_Monthly
    End Select
    
    For Each rLoopCells In rACells
        rCellValue = rLoopCells.Value
        rLoopCells = Trim(rCellValue)
        rLoopCells.Value = UCase(rLoopCells.Value)
    Next rLoopCells

End Function
'---------------------------------------------------------------------------------------
' Date Acquired: August 21, 2012
' Source: http://www.rondebruin.nl/win/s3/win026.htm
'---------------------------------------------------------------------------------------
' Date Edited  April 20, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : OriginalFolder
' Description  : This function check if the Root Folder exists. If not it creates it.
'                It also creates a sub-folder which is renamed after the Station Name.
' Parameters   : String, Integer
' Returns      : String
'---------------------------------------------------------------------------------------
Function CopyOriginalFolder(ByVal OFolder As String, ByVal OFolderInt As Integer) As String

    Dim FSO As Object
    Dim CSPath As String, OrigPath As String, ProPath As String
    Dim FolderPath As String, TempPath As String
    Dim ProcessedFolder As String, ProcessedFolderPath As String
    Dim OriginalFolderInt As Integer
    
    ' Disable all the pop-up menus
    Application.ScreenUpdating = False
       
    Set FSO = CreateObject("scripting.filesystemobject")

    OrigPath = OFolder
    OriginalFolderInt = OFolderInt
    If Right(OrigPath, 1) = "\" Then
        OrigPath = Left(OrigPath, Len(OrigPath) - 1)
    End If
    If FSO.FolderExists(OrigPath) = False Then
        Debug.Print "Original Path Folder doesn't exist"
        MkDir (OrigPath)
    Else: Debug.Print "Original Path Folder exist"
    End If
    
    If Right(OrigPath, 1) <> "\" Then
        FolderPath = Right(OrigPath, OriginalFolderInt)
        Debug.Print FolderPath
        CSPath = Replace(OrigPath, FolderPath, "")
        Debug.Print CSPath
    Else 'Right(OrigPath, 1) = "\"
        OriginalFolderInt = OriginalFolderInt + 1
        FolderPath = Right(OrigPath, OriginalFolderInt)
        Debug.Print FolderPath
        CSPath = Replace(OrigPath, FolderPath, "")
        Debug.Print CSPath
    End If
    
    ProcessedFolder = "_PROCESSED"
    ProPath = CSPath & ProcessedFolder
    If Right(ProPath, 1) = "\" Then
        ProPath = Left(ProPath, Len(ProPath) - 1)
    End If
    If FSO.FolderExists(ProPath) = False Then
        Debug.Print "Processed Path Folder doesn't exist"
        MkDir (ProPath)
    Else: Debug.Print "Processed Path Folder exist"
    End If
    
    '---------------------------------------------------------------------
    ' Copy Source Folder. Make sure that the strings do not end in \.
    ' Otherwise, error will occur. "Error 76 - Path not found."
    '---------------------------------------------------------------------
    FSO.CopyFolder Source:=OrigPath, Destination:=ProPath
    Debug.Print "The original copy of the files are located " & OrigPath
    Debug.Print "The processed copy of the files will be located " & ProPath
    
    CopyOriginalFolder = ProPath
    Debug.Print CopyOriginalFolder

ErrHandler:
    If Err.Number <> 0 Then
        Debug.Print Err.Number
        Err.Clear
        Resume Next
    End If
End Function
