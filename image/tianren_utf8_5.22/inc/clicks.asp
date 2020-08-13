<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../core/Conn.asp"-->
<!--#include file="../core/fun/fun.asp"-->
<%
Const dbpath = "../"
id = tr_killstr(request.querystring("id"), 1, 8, 2, "", "", "", "", 2)
If id<>""Then
  Call changevalue("tr_article", " clicks=clicks+1 ", "id="&id&"")
  clicks = getfieldvalue("tr_article", "clicks", " and id="&id&" ", 2, " order by id desc ")
  If clicks = "" Then clicks = 0
  Response.Write("top.clicks.innerHTML="""&clicks&""";")
End If
%>
