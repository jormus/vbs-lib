
Dim url: url = "http://gensvw1206:1880/smartfacts"
Dim payload: payload = "{" & chr(34) & "infoToSend" & chr(34) & ":" & chr(34) & "theInfo" & chr(34) & "}"

Call testHTTP(url,payload)
Wscript.Quit(0)







Sub testHTTP(Byval url, Byval payload)
	Dim objHttp: Set objHttp = CreateObject("MSXML2.ServerXMLHTTP")
    
    'ignore cert errors
    'objHttp.SetOption 2, SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS
    objHttp.open "POST", url, False

    objHttp.SetRequestHeader "Content-Type", "application/json"
    objHttp.SetRequestHeader "User-Agent", "My Agent String"

    

    objHttp.send payload




    If Not objHttp.status = 200 Then
        'error response
        WScript.Echo objHttp.status & ": " & objHttp.statusText
    Else
        WScript.Echo objHttp.status & ": " & objHttp.statusText
        'WScript.Echo objHttp.responseBody
        WScript.Echo objHttp.responseText
    End If
End Sub



Function CreatePayload






End Function