Attribute VB_Name = "mod_files"
Option Explicit


Public Function FileExists(FileName As String) As Boolean
    FileExists = Dir(FileName) <> ""
End Function
