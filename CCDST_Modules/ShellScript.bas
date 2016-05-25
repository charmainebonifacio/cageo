Attribute VB_Name = "ShellScript_DoNotChange"
Option Explicit
'---------------------------------------------------------------------------------------
' Date Created : April 6, 2013
' Source : http://msdn.microsoft.com/en-us/library/office/ee691831(v=office.14).aspx
'---------------------------------------------------------------------------------------
' Date Edited  : March 27, 2014
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : -
' Description  : Checks whether the code base is the new Microsoft Visual Basic for
'                Applications 7.0. If so, PtrSafe quantifier must be included for
'                64-Bit and Microsoft Office 2010. Otherwise, use regular Declare for
'                32-bit systems and older Microsoft Office
'---------------------------------------------------------------------------------------
#If Win64 Then ' For 64-Bit Compatibility
    Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" _
        Alias "URLDownloadToFileA" (ByVal pCaller As Long, _
        ByVal szURL As String, ByVal szFileName As String, _
        ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
#Else
    Private Declare Function URLDownloadToFile Lib "urlmon" _
        Alias "URLDownloadToFileA" (ByVal pCaller As Long, _
        ByVal szURL As String, ByVal szFileName As String, _
        ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
#End If
'---------------------------------------------------------------------------------------
' Date Acquired : August 21, 2012
' Source : http://www.digitalcoding.com/Code-Snippets/VB/Visual-Basic-Code-Snippet-Download-File-from-URL.html
' Source : http://www.hitechcoach.com/index.php?option=com_content&view=article&id=44:download-a-file-from-a-url&catid=27:vba
'---------------------------------------------------------------------------------------
' Date Edited  : April 6, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : DownloadURLtoFile
' Description  : Download File from the given URL Path and saves it the Folder Path.
'---------------------------------------------------------------------------------------
Public Sub DownloadURLtoFile(pURL As String, pFullFilePath As String)

    Call URLDownloadToFile(0, pURL, pFullFilePath, 0, 0)

End Sub

