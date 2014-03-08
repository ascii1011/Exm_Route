Attribute VB_Name = "mod_url"
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long

' Open the default browser on a given URL
' Returns True if successful, False otherwise

Public Function OpenBrowser(ByVal URL As String) As Boolean
    Dim res As Long
    
    ' it is mandatory that the URL is prefixed with http:// or https://
    If InStr(1, URL, "http", vbTextCompare) <> 1 Then
        URL = "http://" & URL
    End If
    
    res = ShellExecute(0&, "open", URL, vbNullString, vbNullString, _
        vbNormalFocus)
    OpenBrowser = (res > 32)
End Function

