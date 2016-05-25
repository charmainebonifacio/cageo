Attribute VB_Name = "CCDST_3_DataDownload"
'---------------------------------------------------------------------------------------
' Date Created : February 19, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Date Edited  : August 14, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : ReturnURLElements
' Description  : The function returns a string of dynamic array with a specified length.
'                In this case, there are eight key elements to look for in a URL. All
'                eight elements must be found before the tool can proceed to scrape all
'                the climate data.
' Parameters   : Integer, Integer
' Returns      : String
'---------------------------------------------------------------------------------------
Function ReturnURLElements(ByVal DataInterval As Integer, ByVal Index As Integer) As String()
    
    Dim elementName() As String
    
    ' Disable all the pop-up menus
    Application.ScreenUpdating = False

    ' Initialize String Array of Elements Name
    ReDim elementName(1 To Index)
    If DataInterval = 49 Then
        Debug.Print "User chose HOURLY DATA."
        elementName(1) = "timeframe"
        elementName(2) = "Prov"
        elementName(3) = "StationID"
        elementName(4) = "hlyRange"
        elementName(5) = "Year"
        elementName(6) = "Month"
        elementName(7) = "Day"
    End If
    If DataInterval = 50 Then
        Debug.Print "User chose DAILY DATA."
        elementName(1) = "timeframe"
        elementName(2) = "Prov"
        elementName(3) = "StationID"
        elementName(4) = "dlyRange"
        elementName(5) = "Year"
        elementName(6) = "Month"
        elementName(7) = "Day"
    End If
    If DataInterval = 51 Then
        Debug.Print "User chose MONTHLY DATA."
        elementName(1) = "timeframe"
        elementName(2) = "Prov"
        elementName(3) = "StationID"
        elementName(4) = "mlyRange"
        elementName(5) = "Year"
        elementName(6) = "Month"
        elementName(7) = "Day"
    End If

    ' Print String Array of Element Name
    For U = LBound(elementName) To UBound(elementName)
        Debug.Print "The element is: " & elementName(U)
    Next U
    
    ReturnURLElements = elementName()

End Function
'---------------------------------------------------------------------------------------
' Date Created : February 19, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Date Edited  : February 19, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : InitializeElements
' Description  : The function returns an initialized string array.
' Parameters   : Integer
' Returns      : String
'---------------------------------------------------------------------------------------
Function InitializeElements(ByVal Index As Integer) As String()

    Dim elementValue() As String
    ReDim elementValue(1 To Index)      ' Initialize String Array of Value
    
    For v = LBound(elementValue) To UBound(elementValue)
        elementValue(v) = "."
    Next v
    
    InitializeElements = elementValue()
    
End Function
'---------------------------------------------------------------------------------------
' Date Acquired : February 19, 2013
' Source: http://msdn.microsoft.com/en-us/library/office/aa164473(v=office.10).aspx
'---------------------------------------------------------------------------------------
' Date Edited  : February 19, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : IsArrayEmpty
' Description  : This function determines whether the array contains any elements. It
'                returns a FALSE if it does contain elments. Otherwise, it returns a
'                TRUE.
' Parameters   : Variant
' Returns      : Boolean
'---------------------------------------------------------------------------------------
Function IsArrayEmpty(varArray As Variant) As Boolean

    Dim lngUBound As Long
   
    On Error Resume Next
    
    ' If the array is empty, an error occurs when you check the array's bounds.
    lngUBound = UBound(varArray)
    IsArrayEmpty = False
    If Err.Number <> 0 Then
       IsArrayEmpty = True
    End If

End Function
'---------------------------------------------------------------------------------------
' Date Created : February 25, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Date Edited  : August 14, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : URLParseAllElements
' Description  : This function parses the URL entered by the user.
'                This is for all the 7 elements.
' Parameters   : String, String, String, Integer, String, String
' Returns      : Boolean
'---------------------------------------------------------------------------------------
Function URLParseAllElements(ByVal URLLink As String, ByRef URLElement() As String, _
ByRef URLValue() As String, ByVal elementIndex As Integer, _
ByVal AssignedChar As String, ByVal DelChar As String)

    Dim valueOfElement As Variant, fullElement As Variant
    Dim elementCount As Integer
    
    ' Disable all the pop-up menus
    Application.ScreenUpdating = False
    
    ' Define Variables
    elementCount = 0
    
    '---------------------------------------------------------------------
    ' C. Last URL Validity Test
    ' All elements must be present before a file can be downloaded!
    ' This will check all the elements.
    '---------------------------------------------------------------------
    For i = LBound(URLElement) To UBound(URLElement)
        valueOfElement = URLParse(URLLink, URLElement(i), AssignedChar, DelChar)
        Debug.Print "Value of the Element: " & valueOfElement
        If IsEmpty(valueOfElement) = False Then
            elementCount = elementCount + 1
        End If
        fullElement = URLElement(i) + AssignedChar + valueOfElement
        Debug.Print "URL Element: " & fullElement
        URLValue(i) = valueOfElement
    Next i
    
    Debug.Print "Number of Valid Elements: " & elementCount
    If elementCount = 7 Then Debug.Print "Full Validity Test complete. URL passed the validity check."

End Function
'---------------------------------------------------------------------------------------
' Date Acquired : 08212012
' Source : http://stackoverflow.com/questions/11675454/how-to-parse-url-parameters-in-vba
'---------------------------------------------------------------------------------------
' Date Edited  : August 22, 2012
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : URLParse
' Description  : This function parses the URL and aids in checking
'                in conjunction with the validity test.
' Parameters   : Variant, String, String, String
' Returns      : Variant
'---------------------------------------------------------------------------------------
Function URLParse(Txt As Variant, Key As String, AssignChar As String, _
Delimiter As String) As Variant
    
    Dim StartPos As Integer, EndPos As Integer, Result As Variant

    If IsNull(Txt) Then
        Result = Null
    ElseIf Len(Key) = 0 Then
            EndPos = InStr(Txt, AssignChar)
        If EndPos = 0 Then
            Result = Trim(Txt)
        Else
            If InStrRev(Txt, " ", EndPos) = EndPos - 1 Then
                EndPos = InStrRev(Txt, Delimiter, EndPos - 2)
            Else: EndPos = InStrRev(Txt, Delimiter, EndPos)
            End If
            Result = Trim(Left(Txt, EndPos))
        End If
    Else
         StartPos = InStr(Txt, Key & AssignChar)
        'Allow for space between Key and Assignment Character
        If StartPos = 0 Then
            StartPos = InStr(Txt, Key & " " & AssignChar)
            If StartPos > 0 Then StartPos = StartPos + Len(Key & " " & AssignChar)
        Else
            StartPos = StartPos + Len(Key & AssignChar)
        End If
        If StartPos = 0 Then
            Parse = Null
        Else
            EndPos = InStr(StartPos, Txt, AssignChar)
            If EndPos = 0 Then
                If Right(Txt, Len(Delimiter)) = Delimiter Then
                    Result = Trim(Mid(Txt, StartPos, _
                                      Len(Txt) - Len(Delimiter) - StartPos + 1))
                Else
                    Result = Trim(Mid(Txt, StartPos))
                End If
            Else
                If InStrRev(Txt, Delimiter, EndPos) = EndPos - 1 Then
                    EndPos = InStrRev(Txt, Delimiter, EndPos - 2)
                Else
                    EndPos = InStrRev(Txt, Delimiter, EndPos)
                End If
                If EndPos < StartPos Then
                    Result = Trim(Mid(Txt, StartPos))
                Else
                    Result = Trim(Mid(Txt, StartPos, EndPos - StartPos))
                End If
            End If

        End If
    End If
    
    ' Results the result of the parsing script
    URLParse = Result

End Function
'---------------------------------------------------------------------------------------
' Date Created : August 21, 2012
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Date Edited  : February 25, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : URLParseRange
' Description  : This function checks parses the range data from the URL and converts
'                it into integer. It uses a loop statement to generate the new download
'                link and the new file to be downloaded.
' Parameters   : String, String, String, String, Integer
' Returns      : String
'---------------------------------------------------------------------------------------
Function URLParseRange(ByVal OriginalFolder As String, ByVal ClimateStationName As String, _
ByVal SIDElement As String, ByVal RangeElement As String, ByVal DataIntervalIndex As Integer)
    
    Dim StartDateRange As String, EndDateRange As String
    Dim StartDate As Date, EndDate As Date, CurrentDate As Date
    Dim YearIndex As Integer, MonthIndex As Integer
    Dim YearElement As String, MonthElement As String ' Hourly and Daily
    Dim SYearElement As String, SMonthElement As String ' Monthly
    Dim EYearElement As String, EMonthElement As String ' Monthly
    Dim downloadFileFromURL As String, downloadedFile As String
    Dim FolderName As String
    
    ' Disable all the pop-up menus
    Application.ScreenUpdating = False

    ' Parse Year
    StartDateRange = Left(RangeElement, 10)
    EndDateRange = Right(RangeElement, 10)
    StartDate = DateValue(StartDateRange)
    EndDate = DateValue(EndDateRange)
    CurrentDate = StartDate
    Debug.Print CurrentDate
    Debug.Print EndDate
    
    ' Original Name
    FolderName = OriginalFolder
    If Right(FolderName, 1) <> "\" Then
        FolderName = FolderName & "\"
    End If
    
    '---------------------------------------------------------------------
    ' Warn Users that the processing of downloading, processing and merging all files would
    ' take awhile to finish. Approximately 10-20 minutes, depending on the number of files
    ' on the queue.
    '---------------------------------------------------------------------
    Call CCDST_WARNINGMESAGE
    
    Application.DisplayStatusBar = True
    Application.StatusBar = "In progress: DOWNLOADING DATA."
    '---------------------------------------------------------------------
    ' Integer 49 = Hourly Data
    ' Integer 50 = Daily Data
    ' Integer 51 = Monthly Data
    '---------------------------------------------------------------------
    Select Case DataIntervalIndex
        Case 49 To 50
            If DataIntervalIndex = 49 Then
                Debug.Print "HOURLY DATA."
            End If
            If DataIntervalIndex = 50 Then
                Debug.Print "DAILY DATA."
            End If
            
            Do While CurrentDate <= EndDate
        
                YearIndex = Year(CurrentDate)
                MonthIndex = Month(CurrentDate)
                
                ' Convert integer back to string
                Debug.Print "Integer: Year " & YearIndex
                YearElement = YearIndex
                Debug.Print "Integer: Month " & MonthIndex
                MonthElement = MonthIndex
                Debug.Print "String: " & YearElement
                Debug.Print "Passing current parameters to change URL String"
                
                ' Initialize File to be Download by changing the URL String
                downloadFileFromURL = URLFileDownloadLink(SIDElement, RangeElement, YearElement, MonthElement, DataIntervalIndex)
                Debug.Print downloadFileFromURL
                
                ' Rename File to be downloaded
                downloadedFile = URLRenameFileDownload(ClimateStationName, YearElement, MonthElement, DataIntervalIndex)
                Debug.Print downloadedFile
                
                ' Download File using Shell Script
                Call DownloadURLtoFile(downloadFileFromURL, FolderName & downloadedFile)
                
                CurrentDate = DateAdd("m", 1, CurrentDate)
                Debug.Print CurrentDate
            Loop

        Case 51
            Debug.Print "MONTHLY DATA."
            ' Start Date Range
            YearIndex = Year(CurrentDate)
            MonthIndex = Month(CurrentDate)
            ' Convert integer back to string
            Debug.Print "Integer: Year " & YearIndex
            SYearElement = YearIndex
            Debug.Print "Integer: Month " & MonthIndex
            SMonthElement = MonthIndex
            
            ' End Date Range
            YearIndex = Year(EndDate)
            MonthIndex = Month(EndDate)
            Debug.Print "Integer: Year " & YearIndex
            EYearElement = YearIndex
            Debug.Print "Integer: Month " & MonthIndex
            EMonthElement = MonthIndex
            
            Debug.Print "String: " & SYearElement & EYearElement
            Debug.Print "Passing current parameters to change URL String"

            ' Initialize File to be Download by changing the URL String
            downloadFileFromURL = URLFileDownloadLink(SIDElement, RangeElement, YearElement, MonthElement, DataIntervalIndex)
            Debug.Print downloadFileFromURL

             ' Rename File to be downloaded
            downloadedFile = URLRenameFileDownloadMonth(ClimateStationName, SYearElement, SMonthElement, _
                             EYearElement, EMonthElement, DataIntervalIndex)
            Debug.Print downloadedFile
            
            ' Download File using Shell Script
            Call DownloadURLtoFile(downloadFileFromURL, FolderName & downloadedFile)

    End Select
    Application.StatusBar = False

End Function
'---------------------------------------------------------------------------------------
' Date Created : August 21, 2012
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Date Edited  : May 17, 2016
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : URLFileDownloadLink
' Description  : This function changes the URL link by changing the Station ID, Range
'                and Year. This generates the new download link.
' Parameters   : String, String, String, String, Integer
' Returns      : String
'---------------------------------------------------------------------------------------
Function URLFileDownloadLink(ByVal CurrentSID As String, ByVal CurrentRange As String, _
ByVal CurrentYear As String, ByVal CurrentMonth As String, ByVal DataIntervalIndex As Integer) As String
    
    Dim c_Domain As String ' CONST Variables
    Dim c_sector As String
    Dim c_submit As String
    Dim c_day As String
    Dim c_format As String
    Dim d_timeframe As String ' Dynamic Variables
    Dim d_SID As String
    Dim d_range As String
    Dim d_year As String
    Dim d_month As String
    Dim newURL As String
    
    ' Disable all the pop-up menus
    Application.ScreenUpdating = False
    
    ' http://climate.weather.gc.ca/climate_data/bulk_data_e.html?format=csv&stationID=1706&Year=${year}&Month=${month}&Day=14&timeframe=1&submit= Download+Data"

    ' CONST VARIABLES
    newURL = ""
    c_Domain = "http://climate.weather.gc.ca/" '##################################
    c_sector = "climate_data/bulk_data_e.html?" ' changed to climate_data
    c_format = "format=csv&"
    c_day = "Day=01&"
    c_submit = "submit=Download+Data" ' This is the last element that needs to be added
    
    '---------------------------------------------------------------------
    ' Integer 49 = Hourly Data
    ' Integer 50 = Daily Data
    ' Integer 51 = Monthly Data
    '---------------------------------------------------------------------
    Select Case DataIntervalIndex
        Case 49
            Debug.Print "HOURLY DATA."
            d_range = "hlyRange=" & CurrentRange & "&"
            d_timeframe = "timeframe=1&"
        Case 50
            Debug.Print "DAILY DATA."
            d_range = "dlyRange=" & CurrentRange & "&"
            d_timeframe = "timeframe=2&"
        Case 51
            Debug.Print "MONTHLY DATA."
            d_range = "mlyRange=" & CurrentRange & "&"
            d_timeframe = "timeframe=3&"
    End Select

    ' Dynamic Variables
    d_SID = "stationID=" & CurrentSID & "&" ' From StationID to stationID
    d_year = "Year=" & CurrentYear & "&" ' Defined Variable using For Loop
    d_month = "Month=" & CurrentMonth & "&"
    
    newURL = c_Domain + c_sector + c_format + d_SID + d_year + d_month + c_day + d_timeframe + c_submit
    Debug.Print "New URL: " & newURL ' The URL has changed from the last revision
    URLFileDownloadLink = newURL

End Function
'---------------------------------------------------------------------------------------
' Date Created : August 21, 2012
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Date Edited  : March 25, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : URLRenameFileDownload
' Description  : This function checks the changes the filename to be downloaded from
'                eng-hourly-MMDDYYYY-MMDDYYYY.csv to StationName_Hourly_YYYYMM.csv and
'                eng-daily-MMDDYYYY-MMDDYYYY.csv to StationName_Daily_YYYYMMDD_YYMMDD.csv
' Parameters   : String, String, String, Integer
' Returns      : String
'---------------------------------------------------------------------------------------
Function URLRenameFileDownload(ByVal StationName As String, ByVal CurrentYear As String, _
ByVal CurrentMonth As String, ByVal DataIntervalIndex As Integer) As String
    
    Dim FileName As String
    Dim DataType As String, FileExt As String, Separator As String
    Dim StartMonth As String, EndMonth As String
    Dim CMonth As String
    
    ' Disable all the pop-up menus
    Application.ScreenUpdating = False
    
    FileName = ""
    FileExt = ".csv"
    Separator = "_"
    StartMonth = "0101"
    EndMonth = "1231"
  
    '---------------------------------------------------------------------
    ' Integer 49 = Hourly Data
    ' Integer 50 = Daily Data
    ' Integer 51 = Monthly Data
    '---------------------------------------------------------------------
    Select Case DataIntervalIndex
        Case 49
            Debug.Print "Name download data to denote HOURLY DATA."
            DataType = "_Hourly_"
            CMonth = CurrentMonth
            If Len(CurrentMonth) < 2 Then ' Check the length of the month
                CMonth = "0" & CurrentMonth
            End If
            FileName = StationName & DataType & CurrentYear & CMonth & FileExt
        Case 50
            Debug.Print "Name download data to denote DAILY DATA."
            DataType = "_Daily_"
            FileName = StationName & DataType & CurrentYear & StartMonth & Separator & CurrentYear & EndMonth & FileExt
    End Select
    
    Debug.Print "Filename to be downloaded: " & FileName
    URLRenameFileDownload = FileName
    
End Function
'---------------------------------------------------------------------------------------
' Date Created : February 24, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Date Edited  : March 25, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : URLRenameFileDownloadMonth
' Description  : This function checks the changes the filename to be downloaded from
'                eng-monthly-MMDDYYY-MMDDYYY.csv to StationName_Monthly_YYYYMM_YYYYMM.csv
' Parameters   : String, String, String, String, String, Integer
' Returns      : String
'---------------------------------------------------------------------------------------
Function URLRenameFileDownloadMonth(ByVal StationName As String, ByVal CurrentYear As String, _
ByVal CurrentMonth As String, ByVal EndYear As String, _
ByVal EndMonth As String, ByVal DataIntervalIndex As Integer) As String
    
    Dim FileName As String
    Dim DataType As String, FileExt As String, Separator As String
    Dim CMonth As String, EMonth As String ' Temp

    ' Disable all the pop-up menus
    Application.ScreenUpdating = False
    
    FileName = ""
    FileExt = ".csv"
    Separator = "_"
    
    '---------------------------------------------------------------------
    ' Integer 49 = Hourly Data
    ' Integer 50 = Daily Data
    ' Integer 51 = Monthly Data
    '---------------------------------------------------------------------
    Select Case DataIntervalIndex
        Case 51
            Debug.Print "Name download data to denote MONTHLY DATA."
            DataType = "_Monthly_"
            CMonth = CurrentMonth
            EMonth = EndMonth
            If Len(CurrentMonth) < 2 Then
                CMonth = "0" & CurrentMonth
            End If
            If Len(EndMonth) < 2 Then
                EMonth = "0" & EndMonth
            End If
            FileName = StationName & DataType & CurrentYear & CMonth & Separator & EndYear & EMonth & FileExt
    End Select

    Debug.Print "Filename to be downloaded: " & FileName
    URLRenameFileDownloadMonth = FileName
    
End Function
'---------------------------------------------------------------------------------------
' Date Acquired: August 21, 2012
' Source: http://www.rondebruin.nl/win/s3/win026.htm
'---------------------------------------------------------------------------------------
' Date Edited  : April 20, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : CreateDestinationFolder
' Description  : This function check if the Root Folder exists. If not it creates it.
'                It also creates a sub-folder which is renamed after the Station Name.
' Parameters   : String, String, Integer
' Returns      : String
'---------------------------------------------------------------------------------------
Function CreateDestinationFolder(ByVal UserSelectedFolder As String, _
ByVal ClimateStationName As String, ByRef OFolderIndex As Integer) As String

    Dim FSO As Object
    Dim MainRootFolderPath As String, SubFolderPath As String
    Dim RootPath As String, OrigPath As String
    Dim ProPath As String, CSPath As String
    Dim RootFolder As String, OriginalFolder As String
    Dim ProcessedFolder As String, CSFolder As String
    Set FSO = CreateObject("scripting.filesystemobject")

    ' Disable all the pop-up menus
    Application.ScreenUpdating = False

    On Error GoTo ErrHandler

    ' Check if Folder has been selected, if not, go to default
    If IsEmpty(UserSelectedFolder) Then
        MainRootFolderPath = "C:"
    Else: 'Initialize Variables
        MainRootFolderPath = UserSelectedFolder
    End If
    If Right(MainRootFolderPath, 1) <> "\" Then
        MainRootFolderPath = MainRootFolderPath & "\"
    End If
    
    '---------------------------------------------------------------------
    ' Create Root Folder if it does not exist
    '---------------------------------------------------------------------
    RootFolder = "CCDST"
    RootPath = MainRootFolderPath & RootFolder
    If Right(RootPath, 1) <> "\" Then
        RootPath = RootPath & "\"
    End If
    If FSO.FolderExists(RootPath) = False Then
        Debug.Print "RootPath Folder doesn't exist"
        MkDir (RootPath)
    Else: Debug.Print "RootPath Folder exist"
    End If

    '---------------------------------------------------------------------
    ' Create Station Folder if it does not exist
    '---------------------------------------------------------------------
    CSPath = RootPath & ClimateStationName
    If Right(CSPath, 1) <> "\" Then
        CSPath = CSPath & "\"
    End If
    If FSO.FolderExists(CSPath) = False Then
        Debug.Print "CSPath Folder doesn't exist"
        MkDir (CSPath)
    Else: Debug.Print "CSPath Folder exist"
    End If

    '---------------------------------------------------------------------
    ' Create Original Folder
    '---------------------------------------------------------------------
    OriginalFolder = "_DOWNLOAD"
    OFolderIndex = Len(OriginalFolder) ' Processing Later
    OrigPath = CSPath & OriginalFolder
    If Right(OrigPath, 1) <> "\" Then
        OrigPath = OrigPath & "\"
    End If
    If FSO.FolderExists(OrigPath) = False Then
        Debug.Print "Original Path Folder doesn't exist"
        MkDir (OrigPath)
    Else: Debug.Print "Original Path Folder exist"
    End If

    '---------------------------------------------------------------------
    ' Create Processed Folder
    '---------------------------------------------------------------------
    ProcessedFolder = "_PROCESSED"
    ProPath = CSPath & ProcessedFolder
    If Right(ProPath, 1) <> "\" Then
        ProPath = ProPath & "\"
    End If
    If FSO.FolderExists(ProPath) = False Then
        Debug.Print "Processed Path Folder doesn't exist"
        MkDir (ProPath)
    Else: Debug.Print "Processed Path Folder exist"
    End If

    Debug.Print "Data download saved under main directory: " & OrigPath
    
    ' Remove "\" character
    If Right(OrigPath, 1) = "\" Then
        OrigPath = Left(OrigPath, Len(OrigPath) - 1)
    End If
    CreateDestinationFolder = OrigPath
    Debug.Print CreateDestinationFolder
    
ErrHandler:
    If Err.Number <> 0 Then
        Debug.Print Err.Number
        Err.Clear
        Resume Next
    End If
End Function
