<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="inc/charsetasp.asp"-->
<!DOCTYPE html>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="core/Conn.asp"-->
<!--#include file="core/fun/fun.asp"-->
<!--#include file="core/fun/core.asp"-->
<!--#include file="core/class/tr_page.asp"-->
<!--#include file="skin/default/kp171014.asp"-->
<!--#include file="skin/default/jj180328.asp"-->

<%'include此行代码供安装插件用勿删改%>
<%
Const funpath = "./"
Const dbpath = "./"
%>
<%'head此行代码供安装插件用勿删改%>
<!--#include file="inc/headtaginc.asp"-->
<title><%=application(siternd & "55trsitetitle")%></title>
<meta name="Keywords" content="<%=application(siternd & "55trsitekeywords")%>">
<meta name="Description" content="<%=application(siternd & "55trsitedescription")%>">
<base target="_blank">
<!--#include file="inc/phoneewmjs.asp"-->
<%'headend此行代码供安装插件用勿删改%>
</head>
<body>
<!--#include file="inc/head.asp"-->
<div class="container trindexcontrol" id="trindexcontrol178888">
<div class="col-lg-4 <%=colbs4%> trindexnew trrow991" id="trindexnew178888a">
<div class="trtitle3"><div class="trtitle3text trfl">最新内容</div></div>
<%call tr_trindexnewjj("",10,30,false, true, false, 13, "new",data_type)%>
</div>
<div class="col-lg-4 <%=colbs4%> trindexnew trindexcontrolhidden1199" id="trindexnew178888b">
<div class="trtitle3"><div class="trtitle3text trfl">精选推荐</div></div>
<%call tr_trindexnewjj("",10,30,false, true, false, 13, "commend",data_type)%>
</div>
<div class="col-lg-4 <%=colbs4%> trindexnew trindexcontrolhidden1199" id="trindexnew178888c">
<div class="trtitle3"><div class="trtitle3text trfl">热门阅读</div></div>
<%call tr_trindexhotjj("",10,30,true, false, false, 13, "hot",data_type)%>
</div>
</div>
<div class="container"> 
<div class="trad3 tradpadding1"> 
<script type="text/javascript"> _55tr_com('tr4')</script> 
</div>
</div>
<%'body1此行代码供安装插件用勿删改%>
<%'body2此行代码供安装插件用勿删改%>
<div class="container trcolumn clearfix" id="trcolblock178888">
<%
sql = "select top 10000 id,colname,isimgtext,jumpurl from tr_column where isindex=1 and isopen=1"
sql = sql & " order by orderid asc "
rs1.Open sql, conn, 0, 1
jj = 0
If Not rs1.EOF Then
  Do While Not rs1.EOF
    jj = jj + 1
    tmpurl = outputurl(rs1("id"), "", rs1("jumpurl"), ishtml)
%>
<%'body3此行代码供安装插件用勿删改%>
<div class="col-lg-4 col-md-6 <%=colbs4%> trnewlist <%=classstr%> trrow991">
<div class="trtitle3 ">
<div class="trtitle3text trfl"><a href="<%=tmpurl%>"><%=rs1("colname")%></a></div>
<span class="trmore trfr"> <a href="<%=tmpurl%>">more..</a> </span> </div>
<%'body4此行代码供安装插件用勿删改%>
<%
Call tr_imgtextnd(rs1("id"), 1, 80, 85, 28, 22, false)
call tr_bksmalllistnd(rs1("id"),12,30,false, true, false, 13, "new",2)
%>
</div>
<%
If jj = 3 Then jj = 0
rs1.movenext
Loop
End If
rs1.Close
%>
<%'body5此行代码供安装插件用勿删改%>
</div>
<!--#include file="inc/bottomads.asp"-->
<div class="container">
<div class="trpublicline clearfix " id="trlink178888">
<div class="trtitle4"> 友情链接 </div>
<%call tr_getlinknd(9999,88,31,20)%>
</div>
</div>
<!--#include file="inc/foot.asp"-->
</body>
</html>
<%call CloseConn()%>
