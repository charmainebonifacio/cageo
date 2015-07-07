Attribute VB_Name = "CCDST_1_Main"
'---------------------------------------------------------------------------------------
' Date Created : February 19, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Date Edited  : March 25, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : CCDST_MAIN
' Description  : This is the main function that organizes the steps before the climate
'                date can be downloaded, processed and merged. It will check for user
'                input; and validate user input. If either the input or the valida-
'                tion fails, then the tool will notify the user of its requirement. It
'                will proceed to create a copy of the downloaded data; process and save
'                the changes. It should present a summary of the where the final
'                "clean" copy of the data is.
' Parameters   : String, String, String, String
' Returns      : -
'---------------------------------------------------------------------------------------
Function CCDST_MAIN(ByVal ClimateStation As String, ByVal URLLink As String, _
ByVal StartDateRange As String, ByVal EndDateRange As String) As Boolean
    
    Dim DataIntervalIndex As Integer, ResponseValue As Integer
    Dim OriginalFolder As String, ProcessedFolder As String, MainFolder As String
    Dim UserSelectedFolder As String, DataFolderPath As String
    Dim OriginalFolderInt As Integer
    Dim start_time As Date, end_time As Date
    Dim I_UserCheckTime As Long, II_ValidationTime As Long
    Dim III_DownloadTime As Long, IV_ProcessingTime As Long, V_MergeTme As Long
    Dim Status As Boolean
    Dim MessageSummary As String, SectionMessage As String, SummaryTitle As String

    
    ' Disable all the pop-up menus
    Application.ScreenUpdating = False
           
    ' Define Variables
    SummaryTitle = "CCDST Diagnostic Summary"
    DataIntervalIndex = 0
    OriginalFolderInt = 0
    MainFolder = ""
    
    '---------------------------------------------------------------------
    ' A. Before downloading, the tool needs to check user input and
    ' validate entries.
    '---------------------------------------------------------------------
    Application.DisplayStatusBar = True
    Application.StatusBar = "In progress: USER INPUT CHECK"
    '---------------------------------------------------------------------
    ' I. Check For USER INPUT
    '---------------------------------------------------------------------
    start_time = Now()
    Status = SECI_USERINPUTCHECK(ClimateStation, URLLink, StartDateRange, EndDateRange)
    end_time = Now()
    I_UserCheckTime = DateDiff("s", CDate(start_time), CDate(end_time))
    If Status = False Then GoTo End_Handler:
    Debug.Print "SECI_USERINPUTCHECK status: " & Status
    Application.StatusBar = False
    
    '---------------------------------------------------------------------
    ' II. Validate USER INPUT
    '---------------------------------------------------------------------
    Application.DisplayStatusBar = True
    Application.StatusBar = "In progress: DATA VALIDATION"
    start_time = Now()
    Status = SECII_URLVALIDITYCHECK(ClimateStation, URLLink, StartDateRange, EndDateRange, DataIntervalIndex)
    end_time = Now()
    II_ValidationTime = DateDiff("s", CDate(start_time), CDate(end_time))
    If Status = False Then GoTo End_Handler:
    Debug.Print "SECII_URLVALIDITYCHECK status: " & Status
    SectionMessage = CCDST_TIMER("I", I_UserCheckTime, "")
    MessageSummary = MessageSummary & SectionMessage & vbLf
    SectionMessage = CCDST_TIMER("II", II_ValidationTime, "")
    MessageSummary = MessageSummary & SectionMessage
    Debug.Print MessageSummary
    Application.StatusBar = False
    MessageSummary = ""

    '---------------------------------------------------------------------
    ' B. Once all user checks and validations are in place, proceed to
    ' download and process the data.
    '---------------------------------------------------------------------
    Application.DisplayStatusBar = True
    Application.StatusBar = "In progress: SELECT FOLDER."
    UserSelectedFolder = UserSelectsFolder     ' User Selects Folder
    ' This is the folder where data gets downloaded to (ie the Original Folder)
    OriginalFolder = CreateDestinationFolder(UserSelectedFolder, ClimateStation, OriginalFolderInt)
    Debug.Print OriginalFolder
    Application.StatusBar = False
    
    '---------------------------------------------------------------------
    ' III. Download File from Environment Canada
    '---------------------------------------------------------------------
    Application.DisplayStatusBar = True
    Application.StatusBar = "In progress: DOWNLOADING DATA."
    start_time = Now()
    Status = SECIII_DOWNLOADDATA(OriginalFolder, ClimateStation, URLLink, StartDateRange, _
        EndDateRange, DataIntervalIndex)
    end_time = Now()
    III_DownloadTime = DateDiff("s", CDate(start_time), CDate(end_time))
    SectionMessage = CCDST_TIMER("III", III_DownloadTime, "")
    MessageSummary = MessageSummary & SectionMessage
    If Status = False Then GoTo End_Handler:
    Debug.Print "SECIII_DATADOWNLOAD status: " & Status
    Application.StatusBar = False
    
    '---------------------------------------------------------------------
    ' IV. Preprocessed Files from Environment Canada
    '---------------------------------------------------------------------
    Application.DisplayStatusBar = True
    Application.StatusBar = "In progress: PROCESSING AND MERGING DATA"
    start_time = Now()
    Status = SECIV_PROCESSMERGEDATA(ClimateStation, MainFolder, OriginalFolder, ProcessedFolder, OriginalFolderInt, DataIntervalIndex)
    end_time = Now()
    IV_ProcessingTime = DateDiff("s", CDate(start_time), CDate(end_time))
    If Status = False Then GoTo End_Handler:
    Debug.Print "SECIV_PROCESSMERGEDATA status: " & Status
    Debug.Print ProcessedFolder
    Debug.Print MainFolder
    SectionMessage = CCDST_TIMER("IV", IV_ProcessingTime, MainFolder)
    MessageSummary = MessageSummary & SectionMessage
    Application.StatusBar = False
    '---------------------------------------------------------------------
    ' Output Summary
    '---------------------------------------------------------------------
    Application.StatusBar = "The Canadian Climate Data Scraping Tool was successful."
    MsgBox MessageSummary, vbOKOnly, SummaryTitle
    Application.StatusBar = False
    
End_Handler:
    If Status = False Then CCDST_MAIN = False
    If Status = True Then CCDST_MAIN = True
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
' Title        : SECI_USERINPUTCHECK
' Description  : This function serves as the first check which looks for user input. If
'                there is entry, then checks for validation.
' Parameters   : String, String, String, String
' Returns      : Boolean
'---------------------------------------------------------------------------------------
Function SECI_USERINPUTCHECK(ByVal ClimateStation As String, ByVal URLLink As String, _
ByVal StartDateRange As String, ByVal EndDateRange As String) As Boolean
    Dim IA_Input As Boolean, IB_Input As Boolean, IC_Input As Boolean
    Dim NotifyUser As String
    Dim MessageSummary As String, SectionMessage As String, SummaryTitle As String
    
    ' Disable all the pop-up menus
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = True
    Application.StatusBar = "In progress: USER INPUT CHECK"
    
    ' Define Booleans
    SECI_USERINPUTCHECK = True
    IA_Input = True
    IB_Input = True
    IC_Input = True
    
    SummaryTitle = "CCDST Error Summary"
    MessageSummary = "The tool encountered an error. Please try again." & vbLf
    MessageSummary = MessageSummary & vbLf
    
    On Error GoTo ErrHandler
    Debug.Print "Proceeding with SECI_USERINPUTCHECK function."
    
    '---------------------------------------------------------------------
    ' Part I. Check User Entry. If there is at least one no entry,
    ' end function right away.
    '---------------------------------------------------------------------
    ' Check if both Station Name and URL are empty.
    If (Len(ClimateStation) = 0 And Len(URLLink) = 0) Then
        NotifyUser = "IA"
        SectionMessage = CCDST_USERNOTIFICATION(NotifyUser)
        MessageSummary = MessageSummary & SectionMessage
        MsgBox MessageSummary, vbOKOnly, SummaryTitle
        SECI_USERINPUTCHECK = False
        GoTo EndHandler
    End If
    
    ' Check if Station Name is empty and URL is not empty.
    If (Len(ClimateStation) = 0 And Len(URLLink) <> 0) Then
        NotifyUser = "IB"
        SectionMessage = CCDST_USERNOTIFICATION(NotifyUser)
        MessageSummary = MessageSummary & SectionMessage
        If Len(URLLink) <> 0 Then
            If CheckWebsiteRoot(URLLink) = False Then
                NotifyUser = "IE"
                SectionMessage = CCDST_USERNOTIFICATION(NotifyUser)
                MessageSummary = MessageSummary & SectionMessage
            End If
        End If
        MsgBox MessageSummary, vbOKOnly, SummaryTitle
        SECI_USERINPUTCHECK = False
        GoTo EndHandler
    End If

    ' Check if Station Name is not empty and URL is empty.
    If (Len(ClimateStation) <> 0 And Len(URLLink) = 0) Then
        If Len(ClimateStation) <> 0 Then
            If IsLegalStationName(ClimateStation) = False Then
                NotifyUser = "ID"
                SectionMessage = CCDST_USERNOTIFICATION(NotifyUser)
                MessageSummary = MessageSummary & SectionMessage
            End If
        End If
        NotifyUser = "IC"
        SectionMessage = CCDST_USERNOTIFICATION(NotifyUser)
        MessageSummary = MessageSummary & SectionMessage
        MsgBox MessageSummary, vbOKOnly, SummaryTitle
        SECI_USERINPUTCHECK = False
        GoTo EndHandler
    End If

    '---------------------------------------------------------------------
    ' Part II. If there is entry, check User Entry.
    '---------------------------------------------------------------------
    ' Check Station Input for invalid characters.
    If Len(ClimateStation) <> 0 Then
        If IsLegalStationName(ClimateStation) = False Then
            IA_Input = False
            NotifyUser = "ID"
            SectionMessage = CCDST_USERNOTIFICATION(NotifyUser)
            MessageSummary = MessageSummary & SectionMessage
        End If
    End If

    ' Validate ROOT website. It must contain specific string.
    If Len(URLLink) <> 0 Then
        If CheckWebsiteRoot(URLLink) = False Then
            IB_Input = False
            NotifyUser = "IE"
            SectionMessage = CCDST_USERNOTIFICATION(NotifyUser)
            MessageSummary = MessageSummary & SectionMessage
        End If
    End If
    
    ' Check Start and End Date input for Optional User Entry
    If (Len(StartDateRange) = 0) Or (Len(EndDateRange) = 0) Then
        Debug.Print "If there is no specified input then use default date range from URL."
        Debug.Print "A non-specified date range entry is a valid entry."
    Else
        If (CheckUserDateRangeInput(StartDateRange) = False) And (CheckUserDateRangeInput(EndDateRange) = False) Then
            Debug.Print "User specified date range is invalid."
            IC_Input = False
            NotifyUser = "IF"
            SectionMessage = CCDST_USERNOTIFICATION(NotifyUser)
            MessageSummary = MessageSummary & SectionMessage
        End If
    End If
    
    '---------------------------------------------------------------------
    ' Check all Boolean within this function. If all are TRUE, proceed.
    '---------------------------------------------------------------------
    Debug.Print "End of SECI_USERINPUTCHECK function."
    If (IA_Input And IB_Input And IC_Input) = False Then
        SECI_USERINPUTCHECK = False
        MsgBox MessageSummary, vbOKOnly, SummaryTitle
        GoTo EndHandler
    End If
    Application.StatusBar = False
ErrHandler:
    If Err.Number <> 0 Then
        Debug.Print "Error Number for debugging Section I: " & Err.Number
        Err.Clear
        Resume Next
    End If
EndHandler:
End Function
'---------------------------------------------------------------------------------------
' Date Created : February 19, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Date Edited  : August 14, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : SECII_URLVALIDITYCHECK
' Description  : This function serves as the second section of the tool. It checks for
'                specific URL elements. It also checks the validity of date range
'                specified, and whether it is out of scope or not.
' Parameters   : String, String, String, String, String
' Returns      : Boolean
'---------------------------------------------------------------------------------------
Function SECII_URLVALIDITYCHECK(ByVal ClimateStation As String, ByVal URLLink As String, _
ByVal StartDateRange As String, ByVal EndDateRange As String, ByRef DataInterval As Integer) As Boolean

    Dim URLElements() As String, URLValue() As String
    Dim URLElementIndex As Integer
    Dim AssignedCharacter As String, DelimeterCharacter As String
    Dim URLDataIntervalIndex As Integer
    Dim RangeElement As String
    Dim IIA_Input As Boolean, IIB_Input As Boolean, IIC_Input As Boolean
    Dim NotifyUser As String
    Dim MessageSummary As String, SectionMessage As String, SummaryTitle As String
 
    ' Disable all the pop-up menus
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = True
    Application.StatusBar = "In progress: DATA VALIDATION"

    ' Define Booleans
    SECII_URLVALIDITYCHECK = True
    IIA_Input = True
    IIB_Input = True
    IIC_Input = True

    ' Define Variables
    URLElementIndex = 7
    AssignedCharacter = "="
    DelimeterCharacter = "&"
        
    SummaryTitle = "CCDST Error Summary"
    MessageSummary = "The tool encountered an error. Please try again." & vbLf
    MessageSummary = MessageSummary & vbLf
    
    On Error GoTo ErrHandler
    Debug.Print "Proceeding with SECII_URLVALIDITYCHECK function."

    '---------------------------------------------------------------------
    ' IIA. Validate Climate Data Online Website.
    '---------------------------------------------------------------------
    If CheckWebsiteValidity(URLLink) = False Then
        Debug.Print "The user must enter the appropriate link address."
        IIA_Input = False
        NotifyUser = "IIA"
        SectionMessage = CCDST_USERNOTIFICATION(NotifyUser)
        MessageSummary = MessageSummary & SectionMessage
    End If

    '---------------------------------------------------------------------
    ' IIB. Validate Link Address Elements.
    '---------------------------------------------------------------------
    URLDataIntervalIndex = ReturnURLTimeFrame(URLLink, URLElementIndex, AssignedCharacter, DelimeterCharacter)
    DataInterval = URLDataIntervalIndex ' Initialize the datainterval
    If CheckURLValidity(URLLink, URLElementIndex, URLDataIntervalIndex, AssignedCharacter, DelimeterCharacter, RangeElement) = False Then
        Debug.Print "The user must enter the correct URL."
        IIB_Input = False
        NotifyUser = "IIB"
        SectionMessage = CCDST_USERNOTIFICATION(NotifyUser)
        MessageSummary = MessageSummary & SectionMessage
    End If
    
    '---------------------------------------------------------------------
    ' IIC. Check user specified data range is within the scope of default date range.
    '---------------------------------------------------------------------
    Debug.Print "The default data range element is: " & RangeElement
    If (Len(StartDateRange) And Len(EndDateRange)) > 5 Then
        Debug.Print "User entered a date range. Check for validity."
        If CheckValidityDateRange(StartDateRange, EndDateRange, RangeElement) = False Then
            IIC_Input = False
            NotifyUser = "IIC"
            SectionMessage = CCDST_USERNOTIFICATION(NotifyUser)
            MessageSummary = MessageSummary & SectionMessage
        End If
    End If

    '---------------------------------------------------------------------
    ' Check all Boolean within this function. If all are TRUE, proceed.
    '---------------------------------------------------------------------
    Debug.Print "End of SECII_URLVALIDITYCHECK function."
    If (IIA_Input And IIB_Input And IIC_Input) = False Then
        SECII_URLVALIDITYCHECK = False
        MsgBox MessageSummary, vbOKOnly, SummaryTitle
        GoTo EndHandler
    End If
    Application.StatusBar = False

ErrHandler:
    If Err.Number <> 0 Then
        Debug.Print "Error Number for debugging Section II: " & Err.Number
        Err.Clear
        Resume Next
    End If
EndHandler:
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
' Title        : SECIII_DATADOWNLOAD
' Description  : This function serves as the second section of the tool where the user
'                entry for the URL is checked for further validation. Then it checks
'                that the data interval selected matches the user's URL entry. Lastly,
'                the functions checks the date range specified
' Parameters   : String, String, String, String, String
' Returns      : Boolean
'---------------------------------------------------------------------------------------
Function SECIII_DOWNLOADDATA(ByVal OrigFolder As String, ByVal ClimateStation As String, ByVal URLLink As String, _
ByVal StartDateRange As String, ByVal EndDateRange As String, ByVal DataInterval As Integer) As Boolean
    
    Dim AssignedCharacter As String, DelimeterCharacter As String
    Dim URLElements() As String, URLValue() As String
    Dim URLElementIndex As Integer
    Dim NewRangeElement As String

     ' Disable all the pop-up menus
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = True
    Application.StatusBar = "In progress: DOWNLOADING DATA."

    ' Define Variables
    URLElementIndex = 8
    URLDataIntervalIndex = 0
    AssignedCharacter = "="
    DelimeterCharacter = "&"
    
    SummaryTitle = "CCDST Error Summary"
    MessageSummary = "The tool encountered an error. Please try again." & vbLf
    MessageSummary = MessageSummary & vbLf

    On Error GoTo ErrHandler
    Debug.Print "Proceeding with SECIII_DOWNLOADCLIMATEDATA function."
    
    ' Initialize URL Element Array depending on user selection
    URLElements = ReturnURLElements(DataInterval, URLElementIndex)
    URLValue = InitializeElements(URLElementIndex)
    If IsArrayEmpty(URLElements) = True Then
        Debug.Print "Array is empty."
        SECIII_DATADOWNLOAD = False
        GoTo EndHandler
    End If
    
    Debug.Print "Array is NOT empty."    ' Proceed to parse URL.
    Call URLParseAllElements(URLLink, URLElements, URLValue, URLElementIndex, AssignedCharacter, DelimeterCharacter)
    
    ' Print NEW Values for the URLValue String
    For X = LBound(URLElements) To UBound(URLElements)
        Debug.Print "Element: " & URLElements(X) & ". NEW value is: " & URLValue(X) & "."
        Debug.Print URLElements(X) & ": " & URLValue(X)
    Next X

    ' For the User Specified Date Range
    If Len(StartDateRange) <> 0 Or Len(EndDateRange) <> 0 Then
        NewRangeElement = UserDateRangeInput(StartDateRange, EndDateRange)
        If Len(NewRangeElement) > 10 Then
            Call URLParseRange(OrigFolder, ClimateStation, URLValue(3), NewRangeElement, DataInterval)
        End If
    Else: Call URLParseRange(OrigFolder, ClimateStation, URLValue(3), URLValue(4), DataInterval) ' Use Default Range
    End If
    
    SECIII_DOWNLOADDATA = True
    Application.StatusBar = False
    
ErrHandler:
    If Err.Number <> 0 Then
        Debug.Print "Error Number for debugging Section III: " & Err.Number
        Err.Clear
        Resume Next
    End If
EndHandler:
End Function
'---------------------------------------------------------------------------------------
' Date Created : February 19, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Date Edited  : March 26, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : SECIV_PROCESSMERGEDATA
' Description  : This function serves as the fourth section of the tool where the data
'                will be processed.
' Parameters   : String, String, String, String, Integer
' Returns      : Boolean
'---------------------------------------------------------------------------------------
Function SECIV_PROCESSMERGEDATA(ByVal ClimateStation As String, ByRef MasterDestinationFolder As String, ByVal OrigFolder As String, _
ByRef DestFolder As String, ByVal OFolderInt As Integer, ByVal DataInterval As Integer) As Boolean
                
     ' Disable all the pop-up menus
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = True
    Application.StatusBar = "In progress: PROCESSING AND MERGING DATA"
    
    SECIV_PROCESSMERGEDATA = True
    
    '-------------------------------------------------------------
    ' Make a copy of the RAW Data. This is important so that the
    ' user has an extra copy of the RAW unprocessed data. The
    ' purpose of this function is to eliminate the need to re-down
    ' load the data in the first place.
    '-------------------------------------------------------------
    DestFolder = CopyOriginalFolder(OrigFolder, OFolderInt)
    
    ' Station Folder
    MasterDestinationFolder = GetStationFolder(DestFolder)
    If Right(MasterDestinationFolder, 1) <> "\" Then
        MasterDestinationFolder = MasterDestinationFolder & "\"
    End If
    '-------------------------------------------------------------
    ' Process Data by deleting unnecessary columns,
    ' replacing Header strings and replacing missing values.
    ' Save Destination Workbook once data processing is finished.
    '-------------------------------------------------------------
    If ProcessMergeData(ClimateStation, DestFolder, MasterDestinationFolder, DataInterval) = False Then SECIV_PROCESSMERGEDATA = False
    
    Application.StatusBar = False
    
End Function
