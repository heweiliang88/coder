<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../inc/charsetasp.asp"-->
<!DOCTYPE html>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<!--#include file="../core/Conn.asp"-->
<!--#include file="../core/fun/fun.asp"-->
<!--#include file="../core/fun/core.asp"-->
<!--#include file="../core/class/tr_page.asp"-->
<!--#include file="../core/class/tr_user.asp"-->
<!--#include file="../core/class/tr_article.asp"-->
<!--#include file="../core/md5.asp"-->
<!--#include file="../skin/default/kp171014.asp"-->
<%'include此行代码供安装插件用勿删改%>
<%
Const funpath = "../"
Const dbpath = "../"
Set ctr_page = New tr_page
Set ctr_user = New tr_user
safemd5 = trmd5(ctr_user.truserislog(1))
pageno = tr_killstr(request.querystring("pageno"), 1, 8, 1, "", "", "", "", 1)
%>
<%'head此行代码供安装插件用勿删改%>
<!--#include file="../inc/headtaginc.asp"-->
<title>会员中心我的文章_<%=application(siternd & "55trsitetitle")%></title>
<!--#include file="../inc/phoneewmjs.asp"-->
<%'headend此行代码供安装插件用勿删改%>
</head>
<body>
<!--#include file="../inc/head.asp"-->
<div class="container trblock " id="trblock178888">
<div class="col-lg-9 <%=colbs9%> trrow1199 trmarb1199 trlist trovh" id="trleft178888">
<div class="trlisttitle1 "><span class="trcrumbs">当前位置： <a href="<%=funpath%>">网站首页</a> &gt;
      会员个人中心 &gt; 我的文章 </span><span class="trmcrumbs">当前位置： <a href="<%=funpath%>">网站首页</a> &gt;
我的文章 </span></div>
<div class="publicnr trpad1">
<%'body1此行代码供安装插件用勿删改%>
<table width="100%" align="center" border="0" cellpadding="3" cellspacing="1" id="mytable" class="trtable2a" >
<%
sql = "select id,title,ispass,addtime from tr_article where contrimd5='"&safemd5&"' and isdel=0 order by addtime desc "
page_size = 20
pagei = 0
n = (pageno -1) * page_size
Call ctr_page.create_rs(sql, rs, conn, page_size, pageno)
If Not rs.EOF Then
%>
<tr class="trfont4 nobor">
<td width="60%" align="center">标题</td>
<td width="8%" align="center">删除</td>
<td width="8%" align="center">状态</td>
<td width="24%" align="center">添加时间</td>
<%'body5此行代码供安装插件用勿删改%>
</tr>
<%
Do While Not rs.EOF
  If pagei>= page_size Then Exit Do
  pagei = pagei + 1
  n = n + 1
%>
<tr class="">
<td width="" align="center"><a href="<%if rs("ispass")=1 then%><%=outputurl("",rs("id"),"",ishtml)%><%else%>unew.asp?id=<%=rs("id")%><%end if%>" target="_blank" title="点击编辑文章"><%=rs("title")%></a></td>
<td width="" align="center"><%if rs("ispass")=0 then%>
<form id="frm<%=rs("id")%>" name="frm<%=rs("id")%>" action="usersave.asp" method="post" target="ifr1" onSubmit="return udelnew(this)" >
<input type="hidden" name="action" id="action" value="delnew"/>
<input type="hidden" name="id" id="id" value="<%=rs("id")%>"/>
<input type="submit" name="bt1" id="bt1" value="删除" class="trbt2"/>
<%'body6此行代码供安装插件用勿删改%>
</form>
<%end if%></td>
<td width="" align="center"><%if rs("ispass")=1 then%>
<span class="trfont3">已通过
<%else%>
<span class="trfont2a">待审
<%end if%>
</span></td>
<td width="" align="center"><%=rs("addtime")%></td>
</tr>
<%
rs.movenext
Loop
Else
%>
<div class="trmidword"> <a href="unew.asp">暂无文章，请先发表。</a></div>
<%
End If
rs.Close
%>
</table>
<%if ctr_page.page_count>1 then%>
<%= ctr_page.create_page("",ctr_page.page_count,pageno,"trpage")%>
<% end if%>
<%'body2此行代码供安装插件用勿删改%>
</div>
</div>
<div class="col-lg-3 <%=colbs3%> trrow1199 trlistright trovh" id="trlistright178888">
<%'body3此行代码供安装插件用勿删改%>
<!--#include file="../inc/umenu.asp"--> 
<!--#include file="../inc/rightinc.asp"-->
<%'body4此行代码供安装插件用勿删改%>
</div>
</div>
<!--#include file="../inc/bottomads.asp"-->
<div class="trpublicline clearfix "> </div>
<!--#include file="../inc/foot.asp"--> 
<!--#include file="../inc/ifr.asp"-->
</body>
</html>