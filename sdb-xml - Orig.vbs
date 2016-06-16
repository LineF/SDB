Const XMLPath = "\\weishaupt.int\SYSVOL\weishaupt.int\scripts\"
'Const XMLPath = ".\"
Const SystemsXML = "wsapsys.xml"
Const UsersXML = "wsaplogon.xml"

Dim xmlDoc, xmlUDoc
Dim strSaplogonIni
ReDim ArrLogon(0)

'IncludeFile
Include("AlgFunc.vbs")

InitIE
' Standard Benutzerinfo
LDAPUserInfo()

' Übergabeparameter auslesen
Dim UserId, Country, LSC, SaplogonIniFile

'Main Program
If Wscript.arguments.count = 2 Then
	UserId = Wscript.Arguments.Item(0)
	Country = Wscript.Arguments.Item(1)
	LSC = 0
Else
'  wscript.echo ";Username and Country Required"
'  Wscript.Quit
	UserID = strusername
	Country = strCountry
	If strLAN Then
		LSC = 0
	Else
		LSC = 1
	End If
End If

Call SapGui(UserId,Country)
WriteText "... SAPLOGON konfiguriert fuer " & UserId
'wscript.sleep 30000
'End Main Program
CloseIE


'****************************************************************

Private Function SapGui(strUser, strCountry)

Dim SaplogonIniFile
SaplogonIniFile = ReadEnvironment("APPDATA") & "\saplogon.ini"
SetEnvironment "USER", "SAPLOGON_INI_FILE", SaplogonIniFile
strSaplogonIni = "; saplogon.ini -> " & strUser & "  (" & strCountry & ")" & CRLF & CRLF

' Laden von SapSystemKonfiguration
Set xmlDoc = CreateObject("Microsoft.XMLDOM")
xmlDoc.Async = "False"
xmlDoc.Load(XMLPath & SystemsXML)

' Laden von BenutzerSysteme
Set xmlUDoc = CreateObject("Microsoft.XMLDOM")
xmlUDoc.Async = "False"
xmlUDoc.Load(XMLPath & UsersXML)

' Benutzersysteme mit SapSysteme verbinden
QueryUserSysXML LCase(strUser), UCase(strCountry) 
QueryXML("SapSystems_Table/Defaults")

' Saplogon.ini schreiben
If ubound(arrLogon) > 0 Then
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.OpenTextFile(SaplogonIniFile, 2,-2)
	objFile.Write strSapLogonIni
	objFile.Close
	'wscript.echo strSapLogonIni
End If

End Function


'****************************************************************
Private Function QueryXML(strQuery)

Dim r

Set colItem = xmlDoc.selectNodes(strQuery)
For Each objItem in colItem
	For Each objChildItem in objItem.childNodes
		If objChildItem.nodetypestring <> "comment" Then
			strSaplogonIni = strSaplogonIni & CRLF
			strSaplogonIni = strSaplogonIni & "[" & objChildItem.nodename & "]" & CRLF
			If objChildItem.hasChildNodes Then
				If not objChildItem.firstChild.hasChildNodes Then
					QuerySys objChildItem.nodename,objChildItem.text
				End If
			Else
				QuerySys objChildItem.nodename,objChildItem.text
			End If
		End If
		If objChildItem.hasChildNodes Then
			For Each objChildChildItem in objChildItem.ChildNodes
				If objChildItem.nodetypestring <> "comment" And objChildChildItem.hasChildNodes Then
					strSaplogonIni = strSaplogonIni & objChildChildItem.nodename & "=" & objChildChildItem.text & CRLF
				End If
			Next
		End If
	Next
Next
End Function

'****************************************************************
Private Function QuerySys(strQuery,strdefault)

Dim strSystem, strSysQuery, n
n = 1
FOR EACH strSystem IN arrLogon
	strSysQuery = strSystem & "/" & strQuery
	Set NodeList = XMLDoc.getElementsByTagName(strSysQuery) 
	If NodeList.Length = 0 Then
		If strQuery = "LowSpeedConnection" Then
			strSaplogonIni = strSaplogonIni & "Item" & n & "=" & LSC & CRLF
		Else
			strSaplogonIni = strSaplogonIni & "Item" & n & "=" & strdefault & CRLF
		End If
	Else
		strSaplogonIni = strSaplogonIni & "Item" & n & "=" & NodeList.Item(0).Text & CRLF
	End If
	n = n + 1
Next
End Function

'****************************************************************
Private Function QueryUserSysXML(strQuery,strCountry)

Dim n, lAll
n = 0
Set colItem = XMLUDoc.getElementsByTagName(strQuery) 
If colItem.Length = 0 Then
	If strCountry = "DE" Then
		Set colItem = XMLUDoc.getElementsByTagName("default.de")
	Else
		Set colItem = XMLUDoc.getElementsByTagName("default")
	End If
	lAll = FALSE
ElseIf colItem.Item(0).text = "ALL" Then
	lAll = TRUE
Else
	lAll = FALSE
End If
If lAll Then
	QueryAllSysXML()
Else
	For Each objItem in colItem
		For Each objChildItem in objItem.childNodes
			redim preserve arrLogon(n)
			arrLogon(n) = UCase(objChildItem.text)
			n = n + 1
		Next
	Next
End If
End Function


'****************************************************************
Private Function QueryAllSysXML()

Dim n
n = 0
Set NodeList = XMLDoc.documentElement.childNodes
For Each xNode In Nodelist
	If xNode.nodetypestring <> "comment" And xNode.nodeName <> "Defaults" Then
		arrLogon(n) = UCase(xNode.nodeName)
		n = n + 1
		redim preserve arrLogon(n)
	End If
Next
redim preserve arrLogon(n-1)
End Function

'********************************************************************
' Use IncludeFile
'********************************************************************
Private Sub Include(strFile)

   Const ForReading = 1
   Dim fso: set fso = CreateObject("Scripting.FileSystemObject")
   Dim IncludeFullName: IncludeFullName = fso.GetParentFolderName(WScript.ScriptFullName) & "\" & strFile
   Dim f: set f = fso.OpenTextFile(IncludeFullName,ForReading)
   Dim s: s = f.ReadAll()
   ExecuteGlobal s
   End Sub
