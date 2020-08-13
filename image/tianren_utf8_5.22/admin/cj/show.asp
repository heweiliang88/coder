<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../../inc/charsetasp.asp"-->
<%
'采集内核：SK采集系统
'QQ81962480 天人文章管理系统
%>
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<link rel="stylesheet" type="text/css" href="../skin/default/style.css">
<title></title>
<script type="text/javascript" src="../js/hover.js"></script>
<link rel="shortcut icon" type="image/x-icon" href="../../favicon.ico">
<link rel="bookmark" type="image/x-icon" href="../../favicon.ico">
</head>

<body style="background:#fff; " class="w1">
<!--#include file="inc/setup.asp"-->

<%
If InStr(","&adminlimitidrange&",", ",5,")<1 Then contrl_message "您没有访问权限，请尝试其他版块。", "", "", "", ""
ID=Trim(skcj.G("ID"))
Lx=Trim(skcj.G("Lx"))
if ID="" then Response.end 
Select Case Lx
Case 1
	Sql="select * from SK_Article where ArticleID=" & id 
End Select
Set Rs=ConnItem.execute(Sql)
If not rs.eof then
%>
<table width="100%" align="center" border="0" cellpadding="3" cellspacing="1" id="mytable" class="size3" >
	<tr class="tr1" >
		<td class="td_title" colspan="17"><!--#include file="../incgoback.asp"-->■ 采集文章查看
</td>
	</tr>
  <%
  Select Case Lx
  Case 1
  	Call Get_Article
	Dim tRs,RsItem,StrTemp,Id,Sql
	     Dim ZdType,Zds,Zdo,Zd,SKZDStr,SKZDStr_arr
		 Sql="select * from SK_ZDField where Flag=1 Order By ID DESC"
		 Set tRs=ConnItem.execute(Sql)
		 if not tRs.bof and not tRs.eof then
		 	While not tRs.eof
				Id=tRs("ID")
				FieldName=tRs("FieldName")
				Response.Write "<tr>"
    			Response.Write "<td height=""25"" width=""30%"" align=""center"" valign=""middle"" >"& trs("FieldIntro") &"</td>"
    			Response.Write "<td align=""center"" valign=""middle"" >"& Rs(FieldName) &"</td>"
 				Response.Write "</tr>"
				tRs.movenext
			Wend
		End if
  End Select 
  %>
</table>
<%
End if
rs.close : set rs=nothing
Sub Get_Article
  Set ARs=ConnItem.execute("Select * from SK_Article Where ArticleID=" & Id)
  IF Not Rs.eof then
	  Response.Write "<tr><td height=""25"" width=""30%"" align=""center"" valign=""middle"" >标题</td>"
	  Response.Write "<td align=""center"" valign=""middle"" >"& Ars("Title") &"</td></tr>"
	  Response.Write "<tr><td height=""25"" width=""30%"" align=""center"" valign=""middle"" >作者</td>"
	  Response.Write "<td align=""center"" valign=""middle"" >"& Ars("Author") &"</td></tr>"
	  Response.Write "<tr><td height=""25"" width=""30%"" align=""center"" valign=""middle"" >来源</td>"
	  Response.Write "<td align=""center"" valign=""middle"" >"& Ars("CopyFrom") &"</td></tr>"
	  Response.Write "<tr><td height=""25"" width=""30%"" align=""center"" valign=""middle"" >录入者</td>"
	  Response.Write "<td align=""center"" valign=""middle"" >"& Ars("Inputer") &"</td></tr>"
	  Response.Write "<tr><td height=""25"" width=""30%"" align=""center"" valign=""middle"" >关键字</td>"
	  Response.Write "<td align=""center"" valign=""middle"" >"& Ars("Keyword") &"</td></tr>"
	  Response.Write "<tr><td height=""25"" width=""30%"" align=""center"" valign=""middle"" >更新时间</td>"
	  Response.Write "<td align=""center"" valign=""middle"" >"& Ars("UpdateTime") &"</td></tr>"
	  Response.Write "<tr><td height=""25"" width=""30%"" align=""center"" valign=""middle"" >文章内容</td>"
	  Response.Write "<td align=""center""   valign=""middle"" >"& Ars("Content")  &"</td></tr>"
	  Response.Write "<tr><td height=""25"" width=""30%"" align=""center"" valign=""middle"" >小图片</td>"
	  Response.Write "<td align=""center"" valign=""middle"" >"& Ars("picpath") &"</td></tr>"
  End if
  Ars.Close
  Set Ars=nothing
End Sub

Call CloseConnItem()
%>
<script type="text/javascript" >
showtable("mytable")
</script>
</body>
</html>
