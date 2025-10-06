On Error Resume Next
strComputer = "."
Set svc = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set adapters = svc.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
For Each cfg In adapters
    Set nic = GetObject("winmgmts:\\" & strComputer & "\root\cimv2:Win32_NetworkAdapter.DeviceID='" & cfg.Index & "'")
    If nic.NetConnectionStatus = 2 Then
        For Each ip In cfg.IPAddress
            If InStr(ip, ":") = 0 Then
                If cfg.DHCPEnabled Then
                    WScript.Echo cfg.Description & " - " & ip & " (DHCP)"
                Else
                    WScript.Echo cfg.Description & " - " & ip & " (Static)"
                End If
            End If
        Next
    End If
Next
