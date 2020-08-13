<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="inc/charsetasp.asp"-->
<!DOCTYPE html>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<!--#include file="core/Conn.asp"-->
<!--#include file="core/fun/fun.asp"-->
<!--#include file="core/fun/core.asp"-->
<!--#include file="core/class/tr_page.asp"-->
<!--#include file="skin/default/kp171014.asp"-->
<%'include此行代码供安装插件用勿删改%>
<%
Const funpath = "./"
Const dbpath = "./"
Set ctr_page = New tr_page
keywords = tr_killstr(request.querystring("keywords"), 2, 30, 2, "", "", "", "", 1)
pageno = tr_killstr(request.querystring("pageno"), 1, 8, 1, "", "", "", "", 1)
%>
<%'head此行代码供安装插件用勿删改%>
<!--#include file="inc/headtaginc.asp"-->
<title>搜索<%=keywords%>_<%=application(siternd & "55trsitetitle")%></title>
<base target="_blank">
<!--#include file="inc/phoneewmjs.asp"-->
<%'headend此行代码供安装插件用勿删改%>
</head>
<body>
<!--#include file="inc/head.asp"-->
<div class="container trblock" id="trblock178888">
<div class="col-lg-9 <%=colbs9%> trrow1199 trmarb1199 trsearch" id="trleft178888">
<div class="trsearchtitle1 ">当前位置： <a href="#">网站首页</a> &gt; 搜索 <%=keywords%> </div>
<div class="publicnr">
<%'body1此行代码供安装插件用勿删改%>
<%call tr_articlelstnd(keywords,20,40,false,false,false,13,"new","trsearchul",1,2,data_type)%>
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
