<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../inc/charsetasp.asp"-->
<!DOCTYPE html>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<!--#include file="../core/Conn.asp"-->
<!--#include file="../core/fun/fun.asp"-->
<!--#include file="../core/fun/core.asp"-->
<!--#include file="../core/md5.asp"-->
<!--#include file="../skin/default/kp171014.asp"-->
<%'include此行代码供安装插件用勿删改%>
<%
Const funpath = "../"
Const dbpath = "../"
%>
<%'head此行代码供安装插件用勿删改%>
<!--#include file="../inc/headtaginc.asp"-->
<title>会员注册_<%=application(siternd & "55trsitetitle")%></title>
<!--#include file="../inc/phoneewmjs.asp"-->
<%'headend此行代码供安装插件用勿删改%>
</head>
<body>
<!--#include file="../inc/head.asp"-->
<div class="container trblock" id="trblock178888">
<div class="col-lg-9 <%=colbs9%> trrow1199 trmarb1199 trlist trovh " id="trleft178888">
<div class="trlisttitle1 "><span class="trcrumbs">当前位置： <a href="<%=funpath%>">网站首页</a> &gt;
      会员注册</span><span class="trmcrumbs">当前位置： <a href="<%=funpath%>">网站首页</a> &gt;
      会员注册</span> </div>
<div class="publicnr trpad1">
<%'body1此行代码供安装插件用勿删改%>
<%if application(siternd & "55tropenuserreg")=1 then%>
<div class="alert alert-warning trnotice1" role="alert"> <%=application(siternd & "55trregnotice")%> </div>
<form name="form1" id="form1" action="usersave.asp" method="post" onSubmit="return checkuseradd(this);" target="ifr1" class="form-horizontal">
<%'body5此行代码供安装插件用勿删改%>
<div class="form-group">
<label for="username" class="col-sm-2 <%=colbs2%> control-label">登录名</label>
<div class="col-sm-10 <%=colbs10%> ">
<input type="text" class="form-control" id="username" name="username" placeholder="登录名">
<span class="trfont2">5-20字符,不可输汉字,只包含数字,字母,下划线</span> </div>
</div>
<div class="form-group">
<label for="passwd1" class="col-sm-2 <%=colbs2%> control-label">密码</label>
<div class="col-sm-10 <%=colbs10%> ">
<input type="password" class="form-control" id="passwd1" name="passwd1" placeholder="密码">
<span class="trfont2">5-20字符,不可输汉字,只包含数字,字母,下划线</span> </div>
</div>
<div class="form-group">
<label for="passwd2" class="col-sm-2 <%=colbs2%> control-label">重复密码</label>
<div class="col-sm-10 <%=colbs10%> ">
<input type="password" class="form-control" id="passwd2" name="passwd2" placeholder="重复密码">
<span class="trfont2">5-20字符,不可输汉字,只包含数字,字母,下划线</span> </div>
</div>
<div class="form-group">
<label for="mail" class="col-sm-2 <%=colbs2%> control-label">邮箱</label>
<div class="col-sm-10 <%=colbs10%> ">
<input type="text" class="form-control" id="mail" name="mail" placeholder="邮箱">
</div>
</div>
<div class="form-group">
<label for="vcode" class="col-sm-2 <%=colbs2%> control-label">验证码</label>
<div class="col-sm-10 <%=colbs10%> ">
<div class="input-group">
<input type="text" class="form-control" id="vcode" name="vcode" placeholder="验证码" aria-describedby="basic-addon2">
<span class="input-group-addon" style="background:none;" id="basic-addon2"><img id="codeimg" src="../core/Code.asp?rd=<%=strrad%>" height="20" width="75" border="0" style="cursor:hand; " title="点击更换" onClick="javascript:this.src='<%=htmldns%>../core/Code.asp?rd='+randomString(10)+''" /></span> </div>
</div>
</div>
<%'body6此行代码供安装插件用勿删改%>
<div class="trdivbtn">
<button name="submit1" type="submit" id="submit1" class="btn btn-primary trbtn1" >提交注册</button>
</div>
<input name="action" type="hidden" id="action" value="reg"/>
<%'body7此行代码供安装插件用勿删改%>
</form>
<%else%>
<div class="alert alert-danger" role="alert">会员注册功能目前已关闭。 </div>
<%end if%>
<%'body2此行代码供安装插件用勿删改%>
</div>
</div>
<div class="col-lg-3 <%=colbs3%> trrow1199 trlistright trovh" id="trlistright178888">
<%'body3此行代码供安装插件用勿删改%>
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
