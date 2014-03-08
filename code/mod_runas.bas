Attribute VB_Name = "mod_runas"
Option Explicit

Private Const LOGON_WITH_PROFILE = &H1&
Private Const LOGON_NETCREDENTIALS_ONLY = &H2&
Private Const CREATE_DEFAULT_ERROR_MODE = &H4000000
Private Const CREATE_NEW_CONSOLE = &H10&
Private Const CREATE_NEW_PROCESS_GROUP = &H200&
Private Const CREATE_SEPARATE_WOW_VDM = &H800&
Private Const CREATE_SUSPENDED = &H4&
Private Const CREATE_UNICODE_ENVIRONMENT = &H400&
Private Const ABOVE_NORMAL_PRIORITY_CLASS = &H8000&
Private Const BELOW_NORMAL_PRIORITY_CLASS = &H4000&
Private Const HIGH_PRIORITY_CLASS = &H80&
Private Const IDLE_PRIORITY_CLASS = &H40&
Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const REALTIME_PRIORITY_CLASS = &H100&

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type

Private Type STARTUPINFO
    cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Byte
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type RunAs_Struct
    Username As String
    Password As String
    Domain As String
    ApplicationName As String
    CommandLine As String
    CurrentDirectory As String
End Type

Public Ra As RunAs_Struct

Private Declare Function CreateProcessWithLogon Lib "Advapi32" Alias "CreateProcessWithLogonW" (ByVal lpUsername As Long, ByVal lpDomain As Long, ByVal lpPassword As Long, ByVal dwLogonFlags As Long, ByVal lpApplicationName As Long, ByVal sCmdLine As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal sCurDir As Long, lpStartupInfo As STARTUPINFO, lpProcessInfo As PROCESS_INFORMATION) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

'example:
'Ra.UserName = "bsmith"
'Ra.Password = "password"
'Ra.Domain = "domain.com"
'Ra.CurrentDirectory = ""
'Ra.ApplicationName = "C:\windows\NOTEPAD.EXE"
'Ra.CommandLine = ""

Public Function RunAs() As Boolean
    Dim StartInfo As STARTUPINFO
    Dim ProcessInfo As PROCESS_INFORMATION
    
    RunAs = False
    
On Error GoTo RunAsErr:

    StartInfo.cb = LenB(StartInfo) 'initialize structure
    StartInfo.dwFlags = 0&
    
    CreateProcessWithLogon StrPtr(Ra.Username), _
                        StrPtr(Ra.Domain), _
                        StrPtr(Ra.Password), _
                        LOGON_WITH_PROFILE, _
                        StrPtr(Ra.ApplicationName), _
                        StrPtr(Ra.CommandLine), _
                        CREATE_DEFAULT_ERROR_MODE Or CREATE_NEW_CONSOLE Or CREATE_NEW_PROCESS_GROUP, _
                        ByVal 0&, _
                        StrPtr(Ra.CurrentDirectory), _
                        StartInfo, ProcessInfo
    
    RunAs = True
    
    CloseHandle ProcessInfo.hThread
    CloseHandle ProcessInfo.hProcess
    Exit Function
    
RunAsErr:
    Exit Function
End Function
