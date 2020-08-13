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
<!--#include file="../core/md5.asp"-->
<!--#include file="../skin/default/kp171014.asp"-->
<%'include此行代码供安装插件用勿删改%>
<%
Const funpath = "../"
Const dbpath = "../"
Set ctr_user = New tr_user
%>
<%'head此行代码供安装插件用勿删改%>
<!--#include file="../inc/headtaginc.asp"-->
<title>会员登陆_<%=application(siternd & "55trsitetitle")%></title>
<!--#include file="../inc/phoneewmjs.asp"-->
<%'headend此行代码供安装插件用勿删改%>
</head>
<body>
<!--#include file="../inc/head.asp"-->
<div class="container trblock" id="trblock178888">
<div class="col-lg-9 <%=colbs9%> trrow1199 trmarb1199 trlist trovh" id="trleft178888">
<div class="trlisttitle1 "><span class="trcrumbs">当前位置： <a href="<%=funpath%>">网站首页</a> &gt;
      会员登录</span><span class="trmcrumbs">当前位置： <a href="<%=funpath%>">网站首页</a> &gt;
      会员登录</span>
<%'call tr_crumbs(id,20)%>
</div>
<div class="publicnr trpad1">
<%'body1此行代码供安装插件用勿删改%>
<form name="form1" id="form1" action="login.asp" method="post" onSubmit="return checkuserlog(this);" class="form-horizontal">
<%'body5此行代码供安装插件用勿删改%>
<div class="form-group">
<label for="username" class="col-sm-2 <%=colbs2%> control-label">登录名</label>
<div class="col-sm-10 <%=colbs10%> ">
<input type="text" class="form-control" id="username" name="username" placeholder="登录名">
</div>
</div>
<div class="form-group">
<label for="passwd" class="col-sm-2 <%=colbs2%> control-label">密码</label>
<div class="col-sm-10 <%=colbs10%> ">
<input type="password" class="form-control" id="passwd" name="passwd" placeholder="密码">
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
<button name="submit1" type="submit" id="submit1" class="btn btn-primary trbtn1" >立即登陆</button>
</div>
<input name="action" type="hidden" id="action" value="login"/>
<%'body7此行代码供安装插件用勿删改%>
</form>
<div class="truserid" role="alert"> <a href="reg.asp">没有账号？注册</a> </div>
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
<%
act = request.Form("action")
Select Case act
  Case "login"
    Call ctr_user.truserlogin()
End Select
%>
<!--#include file="../inc/ifr.asp"-->
</body>
</html>
