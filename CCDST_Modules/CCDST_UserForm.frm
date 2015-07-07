VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8400
   OleObjectBlob   =   "CCDST_UserForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Date Created : November 18, 2012
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Date Edited  : March 27, 2014
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Purpose      : The Canadian Climate Data Scraping Tool
' Course       : GEOGRAPHY 4990
' Supervisor   : Chris Hugenholtz
'---------------------------------------------------------------------------------------
Private Sub CommandButton1_Click()

    Dim StationName As String, URLPath As String, TmpName As String
    Dim StartDateInput As String, EndDateInput As String
    
    StationName = UserForm1.TextBox1.Value
    URLPath = UserForm1.TextBox2.Value
    StartDateInput = UserForm1.TextBox3.Value
    EndDateInput = UserForm1.TextBox4.Value
    Debug.Print "Station Name User Input: " & StationName
    Debug.Print "URL User Input: " & URLPath
    Debug.Print "Start Date User Input: " & StartDateInput
    Debug.Print "End Date User Input: " & EndDateInput
    
    'Check Station Name Here
    If InStr(StationName, " ") <> 0 Then
        TmpName = Replace(StationName, " ", "_")
        StationName = TmpName
    End If
    
    UserForm1.Hide
    
    If Val(Application.Version) <= 9 Then MsgBox "You are using Microsoft Excel 2003 and older."
    If Val(Application.Version) > 9 Then
        Debug.Print "Microsoft Excel 2007 or higher."
        If CCDST_MAIN(StationName, URLPath, StartDateInput, EndDateInput) = True Then Unload Me ' Clear Form
    End If
    Start_Here ' Start Tool Again
End Sub

Private Sub CommandButton2_Click()
' Clear Form
    Unload Me
    Debug.Print "Form cleared of user entry."
    Start_Here
End Sub

