' BGInfo_JoinStatus.vbs
' ---------------------------------------------------------
' Purpose: Display whether the system is Domain joined, Azure AD joined,
' Hybrid joined, or in a Workgroup. Designed for BGInfo integration.
' Author: ChatGPT (GPT-5)
' ---------------------------------------------------------

Option Explicit

Dim objNetwork, objWMIService, colItems, objItem
Dim WshShell, regValue, systemName, domainName, workgroupName
Dim domainStatus, azureStatus, joinStatus
Dim Status

' Create objects
Set objNetwork = CreateObject("WScript.Network")
Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set WshShell = CreateObject("WScript.Shell")

' Get basic system info
systemName = objNetwork.ComputerName

' Query WMI for domain/workgroup
Set colItems = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")
For Each objItem in colItems
    domainName = objItem.Domain
    workgroupName = objItem.Workgroup
Next

' Default values
domainStatus = False
azureStatus = False

' Check if domain joined (domain ≠ workgroup)
If domainName <> workgroupName Then
    domainStatus = True
End If

' Check Azure AD join status (registry lookup)
On Error Resume Next
regValue = WshShell.RegRead("HKLM\SOFTWARE\Microsoft\Enrollments\AADTenantID")
If Err.Number = 0 And regValue <> "" Then
    azureStatus = True
End If
On Error GoTo 0

' Determine status text
If domainStatus And azureStatus Then
    joinStatus = "Hybrid (Domain + Azure AD)"
ElseIf domainStatus Then
    joinStatus = "Domain Joined"
ElseIf azureStatus Then
    joinStatus = "Azure AD Joined"
Else
    joinStatus = "Workgroup / Local Only"
End If

' Output text (BGInfo expects plain text)
Echo joinStatus

' Clean up
Set objNetwork = Nothing
Set objWMIService = Nothing
Set WshShell = Nothing
