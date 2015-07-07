Attribute VB_Name = "CCDST_2_DataInputValidation"
'---------------------------------------------------------------------------------------
' Date Acquired : March 26, 2013
' Source : http://www.codeforexcelandoutlook.com/excel-vba/validate-filenames/
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Date Edited  : March 26, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : IsLegalStationName
' Description  : This function returns TRUE if valid characters are entered. False,
'                otherwise.
' Parameters   : String
' Returns      : Boolean
'---------------------------------------------------------------------------------------
Function IsLegalStationName(ByVal ClimateStation As String) As Boolean
    IsLegalStationName = True
    Dim InvalidChar()
    Dim CharIndex As Integer
    Dim InvalidIndex As Integer

    CharIndex = 28
    
    ReDim InvalidChar(0 To CharIndex)
    InvalidChar(0) = "["
    InvalidChar(1) = "]"
    InvalidChar(2) = "("
    InvalidChar(3) = ")"
    InvalidChar(4) = "/"
    InvalidChar(5) = "\"
    InvalidChar(6) = "{"
    InvalidChar(7) = "}"
    InvalidChar(8) = "|"
    InvalidChar(9) = "!"
    InvalidChar(10) = "@"
    InvalidChar(11) = "#"
    InvalidChar(12) = "$"
    InvalidChar(13) = "%"
    InvalidChar(14) = "^"
    InvalidChar(15) = "&"
    InvalidChar(16) = "*"
    InvalidChar(17) = ";"
    InvalidChar(18) = ":"
    InvalidChar(19) = """"
    InvalidChar(20) = "'"
    InvalidChar(21) = "<"
    InvalidChar(22) = ">"
    InvalidChar(23) = "?"
    InvalidChar(24) = "="
    InvalidChar(25) = "-"
    InvalidChar(26) = "+"
    InvalidChar(27) = "~"
    InvalidChar(28) = "`"
    
    For X = LBound(InvalidChar) To UBound(InvalidChar)
        InvalidIndex = InStr(ClimateStation, InvalidChar(X))
        If InvalidIndex > 0 Then
            IsLegalStationName = False
            Exit Function
        End If
    Next X
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
' Title        : ReturnDataIntervalIndex
' Description  : This function returns an integer that corresponds to the data interval
'                option on the interface.
' Parameters   : String, String
' Returns      : Integer
'---------------------------------------------------------------------------------------
Function ReturnDataIntervalIndex(ByVal Index As Integer) As Integer
    ReturnDataIntervalIndex = Index
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
' Title        : CheckUserDateRangeInput
' Description  : This function returns a FALSE if the date range entered by the user
'                is invalid. Otherwise, it returns TRUE. No entry is a valid entry.
' Parameters   : String, String
' Returns      : Boolean
'---------------------------------------------------------------------------------------
Function CheckUserDateRangeInput(ByVal InputDateRange As String) As Boolean

    Dim InputDate As Date     ' yyyy-mm-dd
    
    CheckUserDateRangeInput = False
    
    ' Check if the date entered is a valid entry
    If IsDate(InputDateRange) = True Then
        CheckUserDateRangeInput = True
        InputDate = DateValue(InputDateRange)
        NewDate = InputDate
    End If
    
    'EXCEPTIONS: If there is no entry, then proceed.
    If Len(InputDateRange) = 0 Then
        CheckUserDateRangeInput = True
    End If

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
' Title        : UserDateRangeInput
' Description  : This function returns the valid date range specified by the user to be
'                used in another function within this module. Environment Canada's URL
'                for any date range is specified as: YYYY-MM-DD|YYYY-MM-DD
' Parameters   : String, String
' Returns      : String
'---------------------------------------------------------------------------------------
Function UserDateRangeInput(ByVal StartDateRange As String, ByVal EndDateRange As String) As String

    Dim UserInput As String
    Dim Connector As String
    Dim UserStartRangeInput As Date, UserEndRangeInput As Date
    
    Connector = "|"
    UserStartRangeInput = DateValue(StartDateRange)
    UserEndRangeInput = DateValue(EndDateRange)
    UserInput = UserStartRangeInput & Connector & UserEndRangeInput
    Debug.Print "The combined user input date range is: " & UserInput
    UserDateRangeInput = UserInput

End Function
'---------------------------------------------------------------------------------------
' Date Created : February 25, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Date Edited  : March 25, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : ReturnURLTimeFrame
' Description  : This function returns the timeframe entered by the user input. This
'                should distinguish between the three data intervals.
' Parameters   : String, Integer, String, String
' Returns      : Integer
'---------------------------------------------------------------------------------------
Function ReturnURLTimeFrame(ByVal URLLink As String, ByVal elementIndex As Integer, _
ByVal AssignedChar As String, ByVal DelChar As String) As Integer
    
    Dim keyTimeFrame As String, TimeFrameValue As Variant
    Dim URLDataIntervalIndex As Integer
    
    ' Define Variables
    elementCount = 0
    keyTimeFrame = "timeframe"
    ReturnURLTimeFrame = 0
    
    '---------------------------------------------------------------------
    ' The timeframe value should match the data interval type.
    ' Timeframe 1 = Hourly Data
    ' Timeframe 2 = Daily Data
    ' Timeframe 3 = Monthly Data
    '---------------------------------------------------------------------
    ' Parse The Specific Keyword To Check for User selection and URL submitted.
    TimeFrameValue = URLParse(URLLink, keyTimeFrame, AssignedChar, DelChar)
    Debug.Print "The timeframe value is: " & TimeFrameValue
    
    ' If the Value is NULL, URL submitted is invalid
    If IsEmpty(TimeFrameValue) Then
        Debug.Print "Invalid URL. Element did not pass validity check. "
        Exit Function
    End If
    
    URLDataIntervalIndex = Asc(TimeFrameValue) ' Convert this to Integer
    ReturnURLTimeFrame = URLDataIntervalIndex
    
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
' Title        : CheckWebsiteRoot
' Description  : This function checks the validity of the website entered by the user.
' Parameters   : String
' Returns      : Boolean
'---------------------------------------------------------------------------------------
Function CheckWebsiteRoot(ByVal URLLink As String) As Boolean
    
    Dim Website As String
   
    ' Define Variables
    CheckWebsiteRoot = False
    Website = "climate.weather.gc.ca" '##########################################################################

    ' Some parts of the URL make it valid, however, the entire string is invalid.
    If InStr(URLLink, Website) > 1 Then
        CheckWebsiteRoot = True
        Debug.Print "The URL website is from the government website. This is valid."
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
' Title        : CheckWebsiteValidity
' Description  : This function checks the validity of the website entered by the user.
' Parameters   : String
' Returns      : Boolean
'---------------------------------------------------------------------------------------
Function CheckWebsiteValidity(ByVal URLLink As String) As Boolean
    
    Dim Website As String
    
    ' Define Variables
    CheckWebsiteValidity = True
    Website = "http://climate.weather.gc.ca/climateData/"

    ' Some parts of the URL make it valid, however, the entire string is invalid.
    If Left(URLLink, 41) <> Website Then
        CheckWebsiteValidity = False
        Debug.Print "Some elements of the URL are invalid. Please try again."
    End If
    
End Function
'---------------------------------------------------------------------------------------
' Date Created : February 25, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Date Edited  : March 15, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : CheckURLValidity
' Description  : This function checks the validity of the website entered by the user.
' Parameters   : String, Integer, Integer, String, String, String
' Returns      : Boolean
'---------------------------------------------------------------------------------------
Function CheckURLValidity(ByVal URLLink As String, ByVal elementIndex As Integer, _
ByVal URLDataIntervalIndex As Integer, ByVal AssignedChar As String, ByVal DelChar As String, ByRef RangeElement As String) As Boolean
    
    Dim keywordToCheck As String
    Dim fullElement As String
    Dim TimeFrameValue As Variant
    
    ' Define Variables
    CheckURLValidity = True
    elementCount = 0
    keywordToCheck = "Nothing"

     ' Disable all the pop-up menus
    Application.ScreenUpdating = False

    '---------------------------------------------------------------------
    ' The timeframe value should match the data interval type.
    ' Integer 49 = Hourly Data
    ' Integer 50 = Daily Data
    ' Integer 51 = Monthly Data
    '---------------------------------------------------------------------
    Select Case URLDataIntervalIndex
        Case 49
            Debug.Print "Timeframe corresponds to HOURLY DATA."
            keywordToCheck = "hlyRange"
        Case 50
            Debug.Print "Timeframe corresponds to DAILY DATA."
            keywordToCheck = "dlyRange"
        Case 51
            Debug.Print "Timeframe corresponds to MONTHLY DATA."
            keywordToCheck = "mlyRange"
        Case Else
            Debug.Print "Timeframe does not exist. Try again."
    End Select
    
    '---------------------------------------------------------------------
    ' Select the appropriate Data from the URL. This must correspond to the
    ' data interval selected by the USER.
    '---------------------------------------------------------------------
    Debug.Print "Keyword to check is: " & keywordToCheck
    
    ' Parse The Specific Keyword To Check for Validity of the URL submitted.
    TimeFrameValue = URLParse(URLLink, keywordToCheck, AssignedChar, DelChar)
    Debug.Print "Checking validity of URL submitted... " & TimeFrameValue
    
    ' If the Value is NULL, URL submitted is invalid
    If IsEmpty(TimeFrameValue) Then
        Debug.Print "Invalid URL. Element did not pass validity check. "
        CheckURLValidity = False
        Exit Function
    End If
    
    RangeElement = TimeFrameValue
    fullElement = keywordToCheck + AssignedChar + RangeElement
    Debug.Print "Valid URL. Element passed validity check: " & fullElement
    
End Function
'---------------------------------------------------------------------------------------
' Date Created : February 25, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Date Edited  : March 15, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : CheckValidityDateRange
' Description  : This function checks the validity of the date range entered by the
'                user.
' Parameters   : String, String, String
' Returns      : Boolean
'---------------------------------------------------------------------------------------
Function CheckValidityDateRange(ByVal StartDateRange As String, _
ByVal EndDateRange As String, ByVal RangeElement As String) As Boolean

    Dim URLStartRangeInput As String, URLEndRangeInput As String
    Dim StartDate As Date, EndDate As Date
    Dim StartDateInput As Date, EndDateInput As Date
    Dim StartDateURL As Date, EndDateURL As Date
    Dim YearIndex As Integer, MonthIndex As Integer
    Dim YearElement As String, MonthElement As String
    
    CheckValidityDateRange = True
    
     ' Disable all the pop-up menus
    Application.ScreenUpdating = False
    
    ' Parse Year from RangeElement (URL)
    ' Convert string to Date Object
    URLStartRangeInput = Left(RangeElement, 10)
    URLEndRangeInput = Right(RangeElement, 10)
    StartDateURL = DateValue(URLStartRangeInput)
    EndDateURL = DateValue(URLEndRangeInput)
    Debug.Print "This is the default date range."
    Debug.Print StartDateURL & "-" & EndDateURL
    
    ' Convert User Input Date Range into Date Object
    StartDateInput = DateValue(StartDateRange)
    EndDateInput = DateValue(EndDateRange)
    Debug.Print "This is the user specified date range."
    Debug.Print StartDateInput & "-" & EndDateInput
    
    If (StartDateURL <= StartDateInput) And (EndDateURL >= EndDateInput) Then
        Debug.Print "RangeElement start date is: " & StartDateURL & " and user input start date is: " & StartDateInput
        Debug.Print "RangeElement end date is: " & EndDateURL & " and user input end date is: " & EndDateInput
        CheckValidityDateRange = True
    Else: CheckValidityDateRange = False
    End If
 
End Function

