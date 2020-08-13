<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="inc/charsetasp.asp"-->
<!DOCTYPE HTML>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!--#include file="core/Conn.asp"-->
<!--#include file="core/fun/fun.asp"-->
<!--#include file="core/fun/core.asp"-->
<!--#include file="core/class/tr_page.asp"-->
<!--#include file="skin/default/kp171014.asp"-->
<%'include此行代码供安装插件用勿删改%>
<%
if ishtml=2 then
funpath = "../"
else
funpath = "./"
end if
Const dbpath = "./"
Set ctr_page = New tr_page
id = tr_killstr(request.querystring("id"), 1, 8, 1, "", "", "", "", 2)
pageno = tr_killstr(request.querystring("pageno"), 1, 8, 1, "", "", "", "", 1)
dohtml = tr_killstr(request.querystring("dohtml"), 1, 8, 3, "", "", "", "", 1)
if dohtml<>1 then call jumphtml(id,pageno)
sql = "select tpagesize,colname,types,keywords,descriptions,queueid from tr_column where id="&id&" "
rs.Open sql, conn, 0, 1
If Not rs.EOF Then
  psize = rs("tpagesize")
  colname = rs("colname")
  types = rs("types")
  keywords = Rs("keywords")
  descriptions = Rs("descriptions")
  queueid = Rs("queueid")
else
response.write "<meta http-equiv=""Refresh"" content=""0; url="&funpath&""" />"&vbcrlf&"</head>"&vbcrlf&"</body>"&vbcrlf&"</html>"
response.End()
End If
rs.Close
rightinccolid=id

%>
<%'head此行代码供安装插件用勿删改%>
<!--#include file="inc/headtaginc.asp"-->
<title><%=colname%>_<%=application(siternd & "55trsitetitle")%></title>
<meta name="Keywords" content="<%=keywords%>">
<meta name="Description" content="<%=descriptions%>">
<base target="_blank">
<!--#include file="inc/phoneewmjs.asp"-->
<%'headend此行代码供安装插件用勿删改%>
</head>
<body>
<!--#include file="inc/head.asp"-->
<div class="container trblock clearfix " id="trblock178888">
<div class=" col-lg-9 <%=colbs9%> trrow1199 trmarb1199 trlist trovh" id="trleft178888">
<div class="trlisttitle1 "><span class="trcrumbs">当前位置： <a href="<%=funpath%>">网站首页</a> &gt;
<%call tr_crumbs(id,20)%>
</span><span class="trmcrumbs">当前位置： <a href="<%=outputurl(id, "", "", ishtml)%>"><%=left(colname,10)%></a></span> </div>
<div class="publicnr clearfix">
<%'body1此行代码供安装插件用勿删改%>
<%call tr_listtopnd(id,5,40,true,false,false,1,"new")%>
<%'body5此行代码供安装插件用勿删改%>
<%if types =1 then%>
<%call tr_articlelstnd(id,psize,40,false,true,false,13,"new","trlistul",1,1,data_type)%>
<%'body6此行代码供安装插件用勿删改%>
<%else%>
<%call tr_piclistnd(id,psize,158,173,12,1,"trimgul clearfix")%>
<%'body7此行代码供安装插件用勿删改%>
<%end if%>
</div>
<%'body2此行代码供安装插件用勿删改%>
</div>
<div class="col-lg-3 <%=colbs3%> trrow1199 trlistright" id="trlistright178888">
<%'body3此行代码供安装插件用勿删改%>
<!--#include file="inc/rightinc.asp"-->
<%'body4此行代码供安装插件用勿删改%>
</div>
</div>
<!--#include file="inc/bottomads.asp"-->
<div class="trpublicline clearfix "> </div>
<!--#include file="inc/foot.asp"-->
</body>
</html>
<%call CloseConn()%>
