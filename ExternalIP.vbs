On Error Resume Next
Dim http, strIP
Set http = CreateObject("MSXML2.XMLHTTP")
http.Open "GET", "https://api.ipify.org/", False
http.Send
strIP = Trim(http.ResponseText)
If Len(strIP) = 0 Then
    WScript.Echo "Unavailable"
Else
    WScript.Echo strIP
End If
