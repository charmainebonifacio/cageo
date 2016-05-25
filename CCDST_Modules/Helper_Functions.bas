Attribute VB_Name = "Helper_Functions_DoNotChange"
'---------------------------------------------------------------------------------------
' Date Acquired : April 16, 2013
' http://www.mrexcel.com/forum/excel-questions/294728-browse-folder-visual-basic-applications.html
'---------------------------------------------------------------------------------------
' Date Edited  : June 5, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : UserSelectsFolder
' Description  : This function returns the string of the folder location selected.
' Parameters   : -
' Returns      : String
'---------------------------------------------------------------------------------------
Function UserSelectsFolder() As String

    Dim UserFolderSelect As FileDialog
    Dim FolderPath As String
    
    Set UserFolderSelect = Application.FileDialog(msoFileDialogFolderPicker)
    With UserFolderSelect
        .title = "Select a Folder."
        .AllowMultiSelect = False
        .InitialFileName = FolderPath
        If .Show <> -1 Then GoTo NextCode
        FolderPath = .SelectedItems(1)
    End With
    
NextCode:
    UserSelectsFolder = FolderPath
    Set UserFolderSelect = Nothing
End Function
'---------------------------------------------------------------------------------------
' Date Created : April 20, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Date Edited  : April 20, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : GetStationFolder
' Description  : This function check if the Root Folder exists. If not it creates it.
'                It also creates a sub-folder which is renamed after the Station Name.
' Parameters   : String
' Returns      : String
'---------------------------------------------------------------------------------------
Function GetStationFolder(ByVal DFolder As String) As String

    Dim FSO As Object
    Dim CSPath As String, DestPath As String, FolderPath As String
    Dim ProcessedFolder As String, ProcessedFolderPath As String
    
    ' Disable all the pop-up menus
    Application.ScreenUpdating = False
       
    Set FSO = CreateObject("scripting.filesystemobject")

    DestPath = DFolder
    
    If Right(DestPath, 1) <> "\" Then
        FolderPath = Right(DestPath, 10)
        Debug.Print FolderPath
        CSPath = Replace(DestPath, FolderPath, "")
        Debug.Print CSPath
    Else
        FolderPath = Right(DestPath, 11)
        Debug.Print FolderPath
        CSPath = Replace(DestPath, FolderPath, "")
        Debug.Print CSPath
    End If
    
    GetStationFolder = CSPath
    Debug.Print GetStationFolder

ErrHandler:
    If Err.Number <> 0 Then
        Debug.Print Err.Number
        Err.Clear
        Resume Next
    End If
End Function
'---------------------------------------------------------------------
' Date Acquired : May 15, 2012
' Source : www.rondebruin.nl/saveas.htm
'---------------------------------------------------------------------
' Date Edited  : March 27, 2014
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : SaveCSV
' Description  : This function saves the workbook in .csv format in
'                any location desired by the user. It opens up in
'                compatibility format.
' Parameters   : Workbook, String, String
' Returns      : -
'---------------------------------------------------------------------
Function SaveCSV(wbDest As Workbook, ByVal MFolder As String, ByVal ClimateStation As String)

    Dim saveFile As String
    Dim FileFormatValue As Long
    
    ' Check the Excel version
    If Val(Application.Version) <= 9 Then
        FileFormatValue = xlCSV ' For CSV
        saveFile = MFolder & "All_" & ClimateStation
        wbDest.SaveAs saveFile, FileFormat:=FileFormatValue, CreateBackup:=False
        wbDest.Close SaveChanges:=False
    End If
        
    If Val(Application.Version) > 9 Then
        FileFormatValue = 6 ' For CSV
        saveFile = MFolder & "All_" & ClimateStation
        wbDest.SaveAs saveFile, FileFormat:=FileFormatValue, CreateBackup:=False
        wbDest.Close SaveChanges:=False
    End If
End Function
'---------------------------------------------------------------------
' Date Created : August 3, 2012
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : August 16, 2012
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : LastRow
' Description  : This function returns the last row count for the
'                activesheet.
' Parameters   : -
' Returns      : Long
'---------------------------------------------------------------------
Function LastRow() As Long

    Dim LastRowIndex As Long
    If WorksheetFunction.CountA(Cells) > 0 Then
       LastRowIndex = Cells.Find(What:="*", After:=Range("A1"), SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    End If
    LastRow = LastRowIndex
    
End Function
'---------------------------------------------------------------------
' Date Created : August 3, 2012
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : August 16, 2012
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : LastCol
' Description  : This function returns the last column count for the
'                activesheet.
' Parameters   : -
' Returns      : Long
'---------------------------------------------------------------------
Function LastCol() As Long

    Dim LastColIndex As Long
    If WorksheetFunction.CountA(Cells) > 0 Then
       LastColIndex = Rows(1).Find(What:="*", After:=Range("A1"), SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    End If
    LastCol = LastColIndex
    
End Function
'---------------------------------------------------------------------
' Date Created : August 3, 2012
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : August 3, 2012
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : RangeAddress
' Description  : This function finds the input string and returns its
'                address.
' Parameters   : String
' Returns      : Address String
'---------------------------------------------------------------------
Function RangeAddress(ByVal InputString As String)

    Dim Found As Range
    Dim DynamicAddress
    
    Set Found = Rows(1).Find(What:=InputString, SearchDirection:=xlNext, SearchOrder:=xlByColumns)
    With Found
        DynamicAddress = Found.Address
        Debug.Print DynamicAddress
    End With

    RangeAddress = DynamicAddress
    
End Function

'---------------------------------------------------------------------
' Date Created: June 4, 2012
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited: June 13, 2012
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : RowCheck
' Description  : This function will check if the worksheet is empty.
'                If worksheet is empty, function exits. Otherwise,
'                it checks for the last row. If found, selects one row
'                after the last known row.
' Parameters   : Worksheet
'---------------------------------------------------------------------
Function RowCheck(WKSheet As Worksheet)

    Dim LastRow As Long
    
    ' Activate correct worksheet
    WKSheet.Activate
    
    '-------------------------------------------------------------
    ' For Empty/New Workbook
    '-------------------------------------------------------------
    If WorksheetFunction.CountA(Cells) = 0 Then
        Range("A1").Select
        Exit Function
    End If

    '-------------------------------------------------------------
    ' Check for the last used row. Select the row after the
    ' last known row.
    '-------------------------------------------------------------
    LastRow = Cells.Find(What:="*", SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
    ActiveSheet.Range("A" & LastRow + 1).Select

End Function
'---------------------------------------------------------------------
' Date Acquired : May 28, 2012
' Source : http://msdn.microsoft.com/en-us/library/ff198177.aspx
'---------------------------------------------------------------------
' Date Edited  : June 13, 2012
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : FindRange
' Description  : This function selects the range using the last used
'                row and column even with missing values in
'                between many columns.
' Parameters   : Worksheet
'---------------------------------------------------------------------
Function FindRange(WKSheet As Worksheet)

    Dim FirstRow&, FirstCol&, LastRow&, LastCol&
    Dim myUsedRange As Range
        
    ' Activate the correct worksheet
    WKSheet.Activate
    
    ' Define variables
    FirstRow = Cells.Find(What:="*", SearchDirection:=xlNext, SearchOrder:=xlByRows).Row
    FirstCol = Cells.Find(What:="*", SearchDirection:=xlNext, SearchOrder:=xlByColumns).Column
    LastRow = Cells.Find(What:="*", SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
    LastCol = Cells.Find(What:="*", SearchDirection:=xlPrevious, SearchOrder:=xlByColumns).Column
    
    ' Select Range using FirstRow, FirstCol, LastRow, LastCol
    Set myUsedRange = Range(Cells(FirstRow, FirstCol), Cells(LastRow, LastCol))
    myUsedRange.Select
    
End Function
'---------------------------------------------------------------------
' Date Created : July 20, 2012
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : July 20, 2012
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : LastRowIndex
' Description  : This function selects the range depending on the row
'                that contains the header information.
' Parameters   : -
'---------------------------------------------------------------------
Function LastRowIndex(ByVal Index As Integer)

    Dim rowIndex As Integer
    Dim Found As Range
    Dim LastCell
    
    Range("A1").Select
    Set Found = Cells.Find(What:="Date/Time", After:=ActiveCell)
    LastCell = Found.Address
    Range(LastCell).Select
    rowIndex = Selection.Row
    
    '-------------------------------------------------------------
    ' Process metadata. The first retains the header information.
    '-------------------------------------------------------------
    If Index = 1 Then
        rowIndex = rowIndex - 1
    End If
        
    '-------------------------------------------------------------
    ' Delete metadata according to RowIndex.
    '-------------------------------------------------------------
    ActiveSheet.Rows("1:" & rowIndex).Select
    Selection.Delete

End Function
'---------------------------------------------------------------------
' Date Acquired : May 28, 2012
' Source : http://msdn.microsoft.com/en-us/library/ff198177.aspx
'---------------------------------------------------------------------
' Date Edited  : August 21 2012
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : FindValuesRange
' Description  : This function selects the values range only.
' Parameters   : -
'---------------------------------------------------------------------
Function FindValuesRange()

    Dim FirstRow&, FirstCol&, LastRow&, LastCol&
    Dim myUsedRange As Range
        
    ' Activate the correct worksheet
    
    ' Define variables
    FirstRow = 2
    FirstCol = 5
    LastRow = Cells.Find(What:="*", SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
    LastCol = Cells.Find(What:="*", SearchDirection:=xlPrevious, SearchOrder:=xlByColumns).Column
    
    ' Select Range using FirstRow, FirstCol, LastRow, LastCol
    Set myUsedRange = Range(Cells(FirstRow, FirstCol), Cells(LastRow, LastCol))
    myUsedRange.Select
    
End Function

