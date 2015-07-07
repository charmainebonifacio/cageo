Attribute VB_Name = "CCDST_NOTIFICATIONS"
'---------------------------------------------------------------------------------------
' Date Created : February 19, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Date Edited  : February 19, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : CCDST_USERNOTIFICATION
' Description  : This function will notify user according to the notification that the
'                specific code it receives.
' Parameters   : String
' Returns      : String
'---------------------------------------------------------------------------------------
Function CCDST_USERNOTIFICATION(ByVal Notification As String) As String

    Dim NotifyUser As String
    
    Select Case Notification
        Case "IA"
            Debug.Print "__User did not enter Station Name and/or URL string."
            NotifyUser = " > Empty station name." & vbLf
            NotifyUser = NotifyUser & " > Empty link address." & vbLf
            Debug.Print "Should have gotten: " & NotifyUser
                
        Case "IB"
            Debug.Print "__User did not enter a station name."
            NotifyUser = " > Empty station name." & vbLf
            Debug.Print "Should have gotten: " & NotifyUser
                
        Case "IC"
            Debug.Print "__User did not enter a link."
            NotifyUser = " > Empty link address." & vbLf
            Debug.Print "Should have gotten: " & NotifyUser
                                                        
        Case "ID"
            Debug.Print "__User entered invalid characters in the textbox field."
            NotifyUser = " > Invalid station name." & vbLf
            Debug.Print "Should have gotten: " & NotifyUser
                                                        
        Case "IE"
            Debug.Print "__The website entered did not pass the validity check."
            NotifyUser = " > Invalid link address." & vbLf
            Debug.Print "Should have gotten: " & NotifyUser
                                                        
        Case "IF"
            Debug.Print "__User entered invalid date range in the text field."
            NotifyUser = " > Invalid date format." & vbLf
            Debug.Print "Should have gotten: " & NotifyUser
                                                                                                                
        Case "IIA"
            Debug.Print "__The entire URL entered did not pass the validity check."
            NotifyUser = " > Invalid link address copied from Environment Canada." & vbLf
            Debug.Print "Should have gotten: " & NotifyUser
                                                                                                               
        Case "IIB"
            Debug.Print "__Missing important elements."
            NotifyUser = " > The link address is missing important elements." & vbLf
            Debug.Print "Should have gotten: " & NotifyUser
            
        Case "IIC"
            Debug.Print "__User input for date range was invalid."
            NotifyUser = " > Invalid date range. Data is not available from the website." & vbLf
            Debug.Print "Should have gotten: " & NotifyUser

       Case "IIIA"
            Debug.Print "__URL Elements could not be initialized."
            NotifyUser = " > URL Elements could not be initialized. Invalid URL Link." & vbLf
            Debug.Print "Should have gotten: " & NotifyUser

        Case "IIIB"
            Debug.Print "Tool has finished downloading data from Environment Canada."
            
        Case "IV"
            Debug.Print "Tool has finished processing data."
        
        Case "VA"
            Debug.Print "Tool has finished merging data."

        Case "VB" ' For when user decides to not follow instructions.
            Debug.Print "Tool has finished processing and merging data."
        
    End Select
    
    CCDST_USERNOTIFICATION = NotifyUser
    
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
' Title        : CCDST_TIMER
' Description  : This function will notify user how much time has elapsed to complete
'                any procedure, mainly the download and processin steps.
' Parameters   : String, Long, String
' Returns      : String
'---------------------------------------------------------------------------------------
Function CCDST_TIMER(ByVal Notification As String, ByVal TimeElapsed As Long, ByVal MFolder As String) As String

    Dim NotifyUser As String
    
    Select Case Notification
        Case "I"
            Debug.Print "SECTION I Timer... Checking"
            Debug.Print " >> Finished checking user input and selection."
            NotifyUser = "User input check took " & TimeElapsed & " seconds."
            Debug.Print "Should have gotten: " & NotifyUser
        Case "II"
            Debug.Print "SECTION II Timer... Validating"
            Debug.Print " >> Finished val.idating user input and selection"
            NotifyUser = "Validation took " & TimeElapsed & " seconds."
            Debug.Print "Should have gotten: " & NotifyUser
        Case "III"
            Debug.Print "SECTION III Timer... Downloading"
            NotifyUser = "DIAGNOSTIC SUMMARY." & vbLf
            NotifyUser = NotifyUser & vbLf
            NotifyUser = NotifyUser & " >> Downloading climate data took " & TimeElapsed & " seconds." & vbLf
            Debug.Print "Should have gotten: " & NotifyUser
        Case "IV"
            Debug.Print "SECTION IV Timer... Processing and Merging"
            NotifyUser = NotifyUser & " >> Processing climate data took " & TimeElapsed & " seconds." & vbLf
            NotifyUser = NotifyUser & " >> The climate data files can be found here: " & vbLf
            NotifyUser = NotifyUser & MFolder & vbLf
            Debug.Print "Should have gotten: " & NotifyUser
    End Select

    CCDST_TIMER = NotifyUser
    
End Function
'---------------------------------------------------------------------------------------
' Date Created : February 19, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Date Edited  : March 29, 2014
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : WarningMessage
' Description  : This function will notify user that tool is currently processing the
'                user request to download, process and merge files.
' Parameters   : -
' Returns      : -
'---------------------------------------------------------------------------------------
Function CCDST_WARNINGMESAGE()

    Dim WarningPrompt As String
    Dim WindowTitle As String
    
    Application.DisplayStatusBar = True
    Application.StatusBar = "Warning Message Activated. Click ok to proceed."
    
    WindowTitle = "The Canadian Climate Data Scraping Tool"
    WarningPrompt = "WARNING: PLEASE READ!" & vbCrLf
    WarningPrompt = WarningPrompt & vbCrLf
    WarningPrompt = WarningPrompt & "The request to download, process and merge data " & _
        "from Environment Canada may take over ten minutes to complete depending on " & _
        "the date range (among other factors). A DIAGNOSTIC " & _
        "SUMMARY window will appear when the download is complete. " & vbCrLf
    WarningPrompt = WarningPrompt & vbCrLf
    WarningPrompt = WarningPrompt & "Please click [OK] to proceed. " & _
        "Microsoft Excel will not be responding to other task at after this point." & vbCrLf
    WarningPrompt = WarningPrompt & vbCrLf
    MsgBox WarningPrompt, vbOKOnly + vbExclamation, WindowTitle
    
    Application.StatusBar = False
    
End Function

