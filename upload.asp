<%@ Language=VBScript %>
<!--#include file="class_request.asp"-->
<html>
<head>
<title>Uploaded file information</title>
</head>
<body>
<%
Dim objRequest, sText

Set objRequest = New ProxyRequest

Response.Write "file1 original filename='" & objRequest("file1_filename") & "'<br>"
Response.Write "file1 original path='" & objRequest("file1_sourcepath") & "'<br>"
Response.Write "file1 mimetype='" & objRequest("file1_mimetype") & "'<br>"

sText = objRequest("file1") ' Grab the data from the file
%>
</body>
</html>
