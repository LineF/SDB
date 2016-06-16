Const SdbXML = "sdb-sonder.xml"
Dim xmlSdb
Dim strSdbCsv, subNodes(20)

'IncludeFile
Include("AlgFunc.vbs")

SdbCsvFile = "sdb.csv"
strSdbCsv = "; SDB-CSV-File" & CRLF

' Laden von SDB
Set xmlSdb = CreateObject("Microsoft.XMLDOM")
xmlSdb.Async = "False"
xmlSdb.Load(SdbXML)

cnt = xmlSdb.selectSingleNode("SdbRoot/SdbXml/sdbC/Service/Name").attributes.getNamedItem("loop").Text
WriteText "count = " & cnt

Set objItem = xmlSdb.selectSingleNode("SdbRoot/SdbXml/sdbC/Service")

For i = 0 to objItem.childNodes.length-1
	Set subNodes(i) = objItem.childNodes.item(i)
	WriteText subNodes(i).nodeName
Next
	

'For Each objChildItem in objItem.childNodes
'Next

'	WriteText objChildItem.nodeName
'			If objChildItem.nodetypestring <> "comment" Then
'			strSaplogonIni = strSaplogonIni & CRLF
'			strSaplogonIni = strSaplogonIni & "[" & objChildItem.nodename & "]" & CRLF
'			If objChildItem.hasChildNodes Then
'				If not objChildItem.firstChild.hasChildNodes Then
'					QuerySys objChildItem.nodename,objChildItem.text
'				End If
'			Else
'				QuerySys objChildItem.nodename,objChildItem.text
'			End If
'		End If
'		If objChildItem.hasChildNodes Then
'			For Each objChildChildItem in objChildItem.ChildNodes
'				If objChildItem.nodetypestring <> "comment" And objChildChildItem.hasChildNodes Then
'					strSaplogonIni = strSaplogonIni & objChildChildItem.nodename & "=" & objChildChildItem.text & CRLF
'				End If
'			Next
'		End If
'Next



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
