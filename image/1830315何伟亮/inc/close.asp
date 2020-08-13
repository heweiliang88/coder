<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html>
<html>
<%Const dbpath = "../"%>
<!--#include file="../core/Conn.asp"-->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>网站已关闭</title>
<link rel="shortcut icon" type="image/x-icon" href="<%=funpath%>favicon.ico" >
<link rel="Bookmark" type="image/x-icon" href="<%=funpath%>favicon.ico">
</head>
<body>
<%
If application(siternd & "55tropensite") = 1 Then
  response.Redirect("../")
  response.End()
Else
  response.Write application(siternd & "55trcloseintro")
End If
%>
</body>
</html>
