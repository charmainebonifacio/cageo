Attribute VB_Name = "CCDST_0_Start"
'---------------------------------------------------------------------
' Date Created : August 1, 2012
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : March 27, 2014
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : The Canadian Climate Data Scraping Tool
' Description  : The purpose of this macro is to contain all macros
'                for scraping data from the Environment Canada.
'---------------------------------------------------------------------
Sub Start_Here()
   
    Dim button1 As String, button2 As String, button3 As String
    Dim button4 As String, button5 As String, button6 As String
    Dim strLabel1 As String, strLabel2 As String
    Dim strLabel3 As String, strLabel4 As String
    Dim strLabel5 As String, strLabel6 As String
    Dim strLabel7 As String, strLabel8 As String
    Dim userFormCaption As String
    
    ' Disable all the pop-up menus
    Application.ScreenUpdating = False

    ' Label Strings
    userFormCaption = "The Canadian Climate Data Scraping Tool"
    
    button1 = "DOWNLOAD DATA"
    button2 = "CLEAR FORM"

    strLabel1 = "Enter Station Name: "
    strLabel2 = "Enter Link Address: "
    strLabel3 = "Start (YYYY-MM-DD): "
    strLabel4 = "End (YYYY-MM-DD): "

    ' UserForm Initialize
    UserForm1.Caption = userFormCaption
    UserForm1.Frame1.Caption = ""
    UserForm1.Frame2.Caption = "OPTIONAL DATE RANGE"
    UserForm1.Frame2.Font.Size = 10
    
    UserForm1.Label1 = strLabel1
    UserForm1.Label2 = strLabel2
    UserForm1.Label3 = strLabel3
    UserForm1.Label4 = strLabel4
    UserForm1.Label1.Font.Size = 9
    UserForm1.Label1.Font.Size = 9
    UserForm1.Label2.Font.Size = 9
    UserForm1.Label3.Font.Size = 9
    UserForm1.Label4.Font.Size = 9
   
    UserForm1.CommandButton1.Caption = button1
    UserForm1.CommandButton2.Caption = button2
    
    UserForm1.CommandButton1.Font.Bold = True
    UserForm1.CommandButton2.Font.Bold = True
    
    Application.StatusBar = "The Canadian Climate Data Scraping Tool is ready."

    UserForm1.Show

End Sub
