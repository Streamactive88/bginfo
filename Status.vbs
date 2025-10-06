On Error Resume Next
Const JOINED_DOMAIN = 1
Const JOINED_WORKGROUP = 2
Dim joinStatus
joinStatus = ""

Set svc = GetObject("winmgmts:\\.\root\cimv2")
Set colCS = svc.ExecQuery("SELECT DomainRole, PartOfDomain FROM Win32_ComputerSystem")
For Each objCS in colCS
    If objCS.PartOfDomain Then
        joinStatus = "Domain Joined"
    Else
        joinStatus = "Standalone / Workgroup"
    End If
Next

Set shell = CreateObject("WScript.Shell")
On Error Resume Next
Azure1 = shell.RegRead("HKLM\SOFTWARE\Microsoft\Enrollments\")
Azure2 = shell.RegRead("HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\CloudDomainJoin\JoinInfo")

If InStr(LCase(Azure1), "azuread") > 0 Then joinStatus = "Azure AD Joined"
If InStr(LCase(Azure2), "azuread") > 0 Then joinStatus = "Hybrid AD Joined"
If Len(joinStatus) = 0 Then joinStatus = "Unknown"

WScript.Echo joinStatus
