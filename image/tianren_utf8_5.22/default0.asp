<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="inc/charsetasp.asp"-->
<!DOCTYPE html>
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
<div class="container trnews trovh clearfix">
<div class="col-lg-3 <%=colbs3%> trfocusa " id="trfocusa178888">
<div class="row">
<%call tr_focusnd("",286,325,5,20,trfocuspc,trfocusm)%>
<div class="trad5"> <script type="text/javascript"> _55tr_com('tr2')</script></div>
<div class="tjyd" id="tjyd178888">
<div class="tjydtt">推荐阅读 </div>
<%call tr_smalllistkp2nd("",30,60,false, false, false, 13, "commend",data_type)%>
</div>
</div>
</div>
<div class=" col-lg-6 <%=colbs6%> trovh clearfix trrow1199 trmarb1199" id="trnewscenter178888">
<%call tr_indextopnd("",2,26,70,false)%>
<%call tr_smalllistkpnd("",50,60,false, false, true, 13, "new",data_type) %>
<div> </div>
</div>
<div class="col-lg-3 <%=colbs3%> trhotnews trfr trovh " id="trhotnews178888">
<div class="trtitle1 ">
<div class="trtitle1text trfl ">热门文章</div>
</div>
<%
call tr_hotsmalllistnd("",10,30,false, false, false, 13,"trhotnewsul")
%>
<div class="trad5"> <script type="text/javascript"> _55tr_com('tr10')</script></div>
<div class="trad5"> <script type="text/javascript"> _55tr_com('tr11')</script></div>
<div class="trad5"> <script type="text/javascript"> _55tr_com('tr12')</script></div>
</div>
</div>
<div class="container trad3 "> 
<script type="text/javascript"> _55tr_com('tr4')</script> 
</div>
<%'body1此行代码供安装插件用勿删改%>
<div class="container trrollimg " id="trtollimg178888">
<div class="trtitle2 ">
<div class="trtitle2text  ">滚动图片</div>
</div>
<%call tr_rollimgkpnd("",10,11,150,163,1140,210)%>
</div>
<%'body2此行代码供安装插件用勿删改%>
<div class="container trcolumn clearfix" id="trcolblock178888">
<%
sql = "select top 10000 id,colname,isimgtext,jumpurl from tr_column where isindex=1 and isopen=1 "
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
