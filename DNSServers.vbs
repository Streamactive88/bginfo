On Error Resume Next
strComputer = "."
Set svc = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set adapters = svc.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
For Each cfg In adapters
    Set nic = GetObject("winmgmts:\\" & strComputer & "\root\cimv2:Win32_NetworkAdapter.DeviceID='" & cfg.Index & "'")
    If nic.NetConnectionStatus = 2 Then
        If IsArray(cfg.DNSServerSearchOrder) Then
            For Each dns In cfg.DNSServerSearchOrder
                WScript.Echo cfg.Description & " - " & dns
            Next
        End If
    End If
Next
