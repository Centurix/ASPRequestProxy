<%
' ASPProxyRequest
' AUTHOR: Chris Read
'
' Take the multipart/form-data and process the sucka
' CONTENT_TYPE holds the post type and boundary in IE, Opera, Netscape
'
' Usage
'
'Set objRequest = New ProxyRequest
'
'Response.Write objRequest("file1_mimetype")
'Response.Write objRequest("file1_filename") & "<br>"
'Response.Write objRequest("file1_sourcepath") & "<br>"
'
'objRequest.SaveItem "file1",Server.MapPath(objRequest("file1_filename"))

Class ProxyRequest
	Private m_objStorage
	Private m_objForm
	Private m_MultiPart

	Public Cookies
	Public ClientCertificate
	Public ServerVariables
	Public QueryString
	Public TotalBytes
	
	Public Function BinaryRead()
		m_objStorage.Position = 0
		BinaryRead = m_objStorage.Read
	End Function

	Public Sub SaveItem(sIndex,sFilename) ' Save and item to file
		Dim aTemp, objCopy
		
		Set objCopy = Server.CreateObject("ADODB.Stream")
		
		objCopy.Open 
		
		if m_objForm.Exists(sIndex) then
			aTemp = Split(m_objForm.Item(sIndex),"|")
			if UBound(aTemp) > 0 then
				m_objStorage.Position = CLng(aTemp(0)) - 1
				if not m_objForm.Exists(sIndex & "_filename") then
					if m_MultiPart then
						sTemp = ASCIIToUNICODE(m_objStorage.Read(CLng(aTemp(1))))
					else
						sTemp = URLDecode(ASCIIToUNICODE(m_objStorage.Read(CLng(aTemp(1)))))
					end if
					objCopy.Type = 2
					objCopy.Charset = "x-ansi"
					objCopy.WriteText sTemp
				else
					objCopy.Type = 1
					m_objStorage.CopyTo objCopy,CLng(aTemp(1))
				end if
				objCopy.Position = 0
				objCopy.SaveToFile sFilename
			end if
		end if
		
	End Sub

	Public Property Get Form(sIndex)
		Dim aTemp
		
		if m_objForm.Exists(sIndex) then
			aTemp = Split(m_objForm.Item(sIndex),"|")
			if UBound(aTemp) > 0 then
				' We have a winner
				m_objStorage.Position = CLng(aTemp(0))
				Form = m_objStorage.Read(aTemp(1))
			end if
		end if
	End Property

	Public Default Property Get Item(sIndex) ' Default property
		Dim aTemp, sTemp
		
		Item = ""
		
		if Right(sIndex,9) = "_filename" or Right(sIndex,9) = "_mimetype" or Right(sIndex,11) = "_sourcepath" then
			' Just return the information
			Item = m_objForm.Item(sIndex)
		elseif m_objForm.Exists(sIndex) then
			aTemp = Split(m_objForm.Item(sIndex),"|")
			if UBound(aTemp) > 0 then
				m_objStorage.Position = CLng(aTemp(0)) - 1
				if not m_objForm.Exists(sIndex & "_filename") then
					if m_MultiPart then
						sTemp = ASCIIToUNICODE(m_objStorage.Read(CLng(aTemp(1))))
					else
						sTemp = URLDecode(ASCIIToUNICODE(m_objStorage.Read(CLng(aTemp(1)))))
					end if
					Item = sTemp
				else
					Item = m_objStorage.Read(aTemp(1))
				end if
			end if
		else
			Item = QueryString(sIndex)
		end if
	End Property

	Private Function ASCIIToUNICODE(sText)
		Dim lTemp, sTemp
		
		sTemp = ""
		
		if IsNull(sText) then
			ASCIIToUNICODE = ""
		else
			For lTemp = 1 To LenB(sText)
				sTemp = sTemp & Chr(AscB(MidB(sText,lTemp,1)))
			Next
		
			ASCIIToUNICODE = sTemp
		end if
	End Function
	
	Private Function UNICODEToASCII(sText)
		Dim lTemp, sTemp
		
		sTemp = ""
		
		if IsNull(sText) then
			UNICODEToASCII = ""
		else
			For lTemp = 1 To Len(sText)
				sTemp = sTemp & ChrB(Asc(Mid(sText,lTemp,1)))
			Next
		
			UNICODEToASCII = sTemp
		end if
	End Function	
	
	Private Function URLDecode(sText)
		Dim reSearch, objMatches, objMatch
		
		Set reSearch = New RegExp
		
		reSearch.Pattern = "%[0-9,A-F]{2}"
		
		reSearch.Global = True
		reSearch.Multiline = True
		reSearch.IgnoreCase = True

		' Regain our spaces
		sText = Replace(sText,"+"," ")
		
		Set objMatches = reSearch.Execute(sText)
		
		For Each objMatch In objMatches
			sText = Replace(sText,objMatch.Value,Chr(CInt("&H" & Right(objMatch.Value,2))))
		Next
		
		URLDecode = sText
    End Function

	Public Function ExportRequest()
		ExportRequest = BinaryRead()
	End Function

	' Load a request session back into this object
	Public Sub ImportRequest(sRequest)
		Dim sText, lTemp
		Dim lContent, lContentStart, lContentLength
		Dim aPairs, aTokens
		
		TotalBytes = LenB(sRequest)
		
		Set m_objStorage = Server.CreateObject("ADODB.Stream")
		Set m_objForm = Server.CreateObject("Scripting.Dictionary")
		
		m_objStorage.Type = 1 ' adTypeBinary
		m_objStorage.Open

		m_objStorage.Write sRequest
		
		m_objStorage.Position = 0

		sText = m_objStorage.Read() ' Referencial string

		' We don't load multi-part forms
		m_MultiPart = False
		m_objStorage.Position = 0
		m_objStorage.Type = 2
		m_objStorage.Charset = "x-ansi"
		aPairs = Split(m_objStorage.ReadText,"&")
		lContentStart = 0
		lContentLength = 0
					
		For lTemp = 0 To UBound(aPairs)
			aTokens = Split(aPairs(lTemp),"=")
			lContentStart = lContentStart + Len(aTokens(0)) + 2
			m_objForm.Add aTokens(0),CStr(lContentStart) & "|" & Len(aTokens(1))
			lContentStart = lContentStart + Len(aTokens(1))
		Next

		m_objStorage.Position = 0
		m_objStorage.Type = 1
	End Sub


	Private Sub Class_Initialize()
		Dim aContentType, sBoundary, sText, lTemp
		Dim sHeader, lHeaderStart, lHeaderLength
		Dim sName, lNameStart, lNameLength
		Dim sFilename, lFilenameStart, lFilenameLength
		Dim sContentType, lContentTypeStart, lContentTypeLength
		Dim lContent, lContentStart, lContentLength
		Dim lPosition, aPairs, aTokens
		
		Set Cookies = Request.Cookies
		Set ClientCertificate = Request.ClientCertificate
		Set ServerVariables = Request.ServerVariables
		Set QueryString = Request.QueryString
		
		TotalBytes = Request.TotalBytes
		
		Set m_objStorage = Server.CreateObject("ADODB.Stream")
		Set m_objForm = Server.CreateObject("Scripting.Dictionary")
		
		m_objStorage.Type = 1 ' adTypeBinary
		m_objStorage.Open

		lPosition = 0
		
		' Speed up the BinaryRead method by reading in 64Kb blocks
		' This can cut a 5Mb transfer from 96 seconds to 4
		while lPosition < Request.TotalBytes
			m_objStorage.Write Request.BinaryRead(65535) ' 64Kb at a time
			lPosition = lPosition + 65535
		wend
		
		m_objStorage.Position = 0

		sText = m_objStorage.Read() ' Referencial string
		
		aContentType = Split(Request.ServerVariables("CONTENT_TYPE"),";")

		if UBound(aContentType) >= 0 then
			Select Case UCase(aContentType(0))
				Case "MULTIPART/FORM-DATA"
					m_MultiPart = True
					sBoundary = LeftB(sText,InStrB(sText,ChrB(13) & ChrB(10)) - 1)
					' Go on a discovery to find all the form elements
					lTemp = InStrB(sText,sBoundary)
					while lTemp > 0
						lHeaderStart = lTemp + LenB(sBoundary) + 2 ' Include CRLF pair
						lHeaderLength = InStrB(lTemp,sText,ChrB(13) & ChrB(10) & ChrB(13) & ChrB(10)) - lHeaderStart + 4
						if lHeaderLength > 0 then ' Element found, add it to our dictionary
							sHeader = ASCIIToUNICODE(MidB(sText,lHeaderStart,lHeaderLength))
							lNameStart = InStr(sHeader,"name=") + 6
							lNameLength = InStr(lNameStart,sHeader,Chr(34)) - lNameStart
							sName = Mid(sHeader,lNameStart,lNameLength)
							sFilename = ""
							sContentType = ""
							lFilenameStart = InStr(sHeader,"filename=") + 10
							if lFilenameStart > 10 then
								lFilenameLength = InStr(lFilenameStart,sHeader,Chr(34)) - lFilenameStart
								sFilename = Mid(sHeader,lFilenameStart,lFilenameLength)
							end if
							lContentTypeStart = InStr(sHeader,"Content-Type: ") + 14
							if lContentTypeStart > 15 then
								lContentTypeLength = InStr(lContentTypeStart,sHeader,Chr(13) & Chr(10)) - lContentTypeStart
								sContentType = Mid(sHeader,lContentTypeStart,lContentTypeLength)
							end if
							lContentStart = lHeaderStart + lHeaderLength
							lContentLength = InStrB(lContentStart,sText,sBoundary) - lContentStart - 2

							' Ok, we add this information to the dictionary
							m_objForm.Add sName,CStr(lContentStart) & "|" & CStr(lContentLength)
							if sFilename <> "" then
								sPath = Left(sFilename,InStrRev(sFilename,"\"))
								sFilename = Right(sFilename,Len(sFilename) - InStrRev(sFilename,"\"))
								m_objForm.Add sName & "_sourcepath",sPath
								m_objForm.Add sName & "_filename",sFilename
							end if
							if sContentType <> "" then
								m_objForm.Add sName & "_mimetype",sContentType
							end if
						end if
						lTemp = InStrB(lTemp + LenB(sBoundary),sText,sBoundary)
					wend
				Case Else
					m_MultiPart = False
					m_objStorage.Position = 0
					m_objStorage.Type = 2
					m_objStorage.Charset = "x-ansi"
					aPairs = Split(m_objStorage.ReadText,"&")
					lContentStart = 0
					lContentLength = 0
					
					For lTemp = 0 To UBound(aPairs)
						aTokens = Split(aPairs(lTemp),"=")
						lContentStart = lContentStart + Len(aTokens(0)) + 2
						m_objForm.Add aTokens(0),CStr(lContentStart) & "|" & Len(aTokens(1))
						lContentStart = lContentStart + Len(aTokens(1))
					Next

					m_objStorage.Position = 0
					m_objStorage.Type = 1
			End Select
		end if
	End Sub
	
	Private Sub Class_Terminate()
		Set Cookies = Nothing
		Set ClientCertificate = Nothing
		Set ServerVariables = Nothing
		Set QueryString = Nothing
		Set m_objStorage = Nothing
		Set m_objForm = Nothing
	End Sub
End Class
%>