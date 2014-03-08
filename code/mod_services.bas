Attribute VB_Name = "mod_services"
Option Explicit



''''''''''''''Finds a service on the local computer'''''''''
'input: sSrvs (as the service name, not display name)
'output: (outputs the state of the service)
    'ex: null or "" = no service was found
    '    "Found: no state" = found service, but no state found
    '    "Running", "Stopped", etc...  (states)
    
Function ServiceExists(sSrvs As String) As String
    Dim wmi_Srvs
    Dim oSrvs
    Dim oSrv
    Dim strComputer As String
    Dim sName As String, sState As String
    
    sSrvs = LCase(sSrvs)
    
    ServiceExists = ""   'Does not exist
    strComputer = "."
    
On Error GoTo Err:
    Set wmi_Srvs = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
    Set oSrvs = wmi_Srvs.ExecQuery("Select * from Win32_Service", , 48)
    
    For Each oSrv In oSrvs
    
        ServiceExists = "Found: no state"   'Does not exist
        sName = LCase(oSrv.name)
        sState = oSrv.State
        'List1.AddItem "Name: " & sName & "   [State:" & sState & "]::"
        
        If sSrvs = sName Then
            ServiceExists = sState
            Exit For
        End If
        
    Next
    
Err:

    Set wmi_Srvs = Nothing
    Set oSrvs = Nothing
    Set oSrv = Nothing
    Exit Function
End Function
