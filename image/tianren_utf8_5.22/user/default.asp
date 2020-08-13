<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../inc/charsetasp.asp"-->
<!DOCTYPE html>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<!--#include file="../core/Conn.asp"-->
<!--#include file="../core/fun/fun.asp"-->
<!--#include file="../core/fun/core.asp"-->
<!--#include file="../core/class/tr_user.asp"-->
<!--#include file="../crinc/ulevstr.asp"-->
<!--#include file="../core/md5.asp"-->
<!--#include file="../skin/default/kp171014.asp"-->
<%'include此行代码供安装插件用勿删改%>
<%
Const funpath = "../"
Const dbpath = "../"
Set ctr_user = New tr_user
uname = ctr_user.truserislog(1)
safemd5 = trmd5(uname)
%>
<%'head此行代码供安装插件用勿删改%>
<!--#include file="../inc/headtaginc.asp"-->
<title>会员中心编辑个人资料_<%=application(siternd & "55trsitetitle")%></title>
<!--#include file="../inc/phoneewmjs.asp"-->
<%'headend此行代码供安装插件用勿删改%>
</head>
<body>
<!--#include file="../inc/head.asp"-->
<div class="container trblock" id="trblock178888">
<div class="col-lg-9 <%=colbs9%> trrow1199 trmarb1199 trlist trovh" id="trleft178888">
<div class="trlisttitle1 "><span class="trcrumbs">当前位置： <a href="<%=funpath%>">网站首页</a> &gt;
      会员个人中心 &gt; 编辑个人资料 </span><span class="trmcrumbs">当前位置： <a href="<%=funpath%>">网站首页</a> &gt;
 编辑个人资料 </span></div>
<div class="publicnr trpad1">
<%'body1此行代码供安装插件用勿删改%>
<div class="truseri1 clearfix">
<%'body5此行代码供安装插件用勿删改%>
<div class="col-lg-4 <%=colbs4%> truserblock"><i class="glyphicon glyphicon-user truserblocki"></i>登录名：
<div class="truserinfo"><%=ctr_user.username%></div>
</div>
<div class="col-lg-4 <%=colbs4%> truserblock"><i class="glyphicon glyphicon-usd truserblocki"></i>总积分：
<div class="truserinfo"><%=ctr_user.points%> 分</div>
</div>
<div class="col-lg-4 <%=colbs4%> truserblock"><i class="glyphicon glyphicon-education truserblocki"></i>会员等级：
<div class="truserinfo">
<%
spuserjfstr = Split(userjfstr, ",")
splevnamestr = Split(levnamestr, ",")
points = clng(ctr_user.points)
'response.write spuserjfstr(UBound(spuserjfstr))
'response.End()
For i = 0 To UBound(spuserjfstr)
  If points>= clng(spuserjfstr(i)) And points<clng(spuserjfstr(i + 1)) Then
    response.Write splevnamestr(i)
    Exit For
  ElseIf points>= clng(spuserjfstr(UBound(spuserjfstr))) Then
    response.Write splevnamestr(UBound(spuserjfstr))
    Exit For
  End If
Next
%>
</div>
</div>
<div class="col-lg-4 <%=colbs4%> truserblock"><i class="glyphicon glyphicon-list truserblocki"></i>发布文章：
<div class="truserinfo"><%=getfieldvalue("tr_article","count(id)"," and contrimd5='"&safemd5&"' and isdel=0 ",2,"")%> 篇</div>
</div>
<div class="col-lg-4 <%=colbs4%> truserblock"><i class="glyphicon glyphicon-text-height truserblocki"></i>发布文章状态：
<div class="truserinfo">
<%if ctr_user.openarticle=1 then%>
<span class="trfont3">已开启</span>
<%else%>
<span class="trfont2">已关闭</span>
<%end if%>
</div>
</div>
<div class="col-lg-4 <%=colbs4%> truserblock"><i class="glyphicon glyphicon-bookmark truserblocki"></i>发布文章赠送的积分：
<div class="truserinfo"><%=ctr_user.gavearticlepoints%> 分</div>
</div>
<div class="col-lg-4 <%=colbs4%> truserblock"><i class="glyphicon glyphicon-check truserblocki"></i>已连续签到：
<div class="truserinfo">
<%
nowtime = FormatDate(Now(), 1)
If DateDiff("d", ctr_user.lastsigntime, nowtime)>1 Or IsNull(ctr_user.lastsigntime) Then
%>
0
<%else%>
<%=ctr_user.signs%>
<%end if%>
天</div>
</div>
<div class="col-lg-4 <%=colbs4%> truserblock"><i class="glyphicon glyphicon-time truserblocki"></i>注册日期：
<div class="truserinfo"><%=FormatDate(ctr_user.regtime, 4) %></div>
</div>
<div class="col-lg-4 <%=colbs4%> truserblock"><i class="glyphicon glyphicon-asterisk truserblocki"></i>签到赠送的积分：
<div class="truserinfo"><%=ctr_user.gavesignpoints%> 分</div>
</div>
<%'body6此行代码供安装插件用勿删改%>
</div>
<div class="trseparateline">基本信息</div>
<form name="form1" id="form1" action="usersave.asp" method="post" onSubmit="return checkuseredit(this);" target="ifr1" class="form-horizontal">
<%'body7此行代码供安装插件用勿删改%>
<div class="form-group">
<label for="passwd0" class="col-sm-2 <%=colbs2%> control-label">原始密码</label>
<div class="col-sm-10 <%=colbs10%> ">
<input type="password" class="form-control" id="passwd0" name="passwd0" placeholder="原始密码">
<span class="trfont2">修改资料需要输入原始密码哦</span> </div>
</div>
<div class="form-group">
<label for="passwd1" class="col-sm-2 <%=colbs2%> control-label">新密码</label>
<div class="col-sm-10 <%=colbs10%> ">
<input type="password" class="form-control" id="passwd1" name="passwd1" placeholder="新密码">
<span class="trfont2">不修改请留空</span> </div>
</div>
<div class="form-group">
<label for="passwd2" class="col-sm-2 <%=colbs2%> control-label">重复新密码</label>
<div class="col-sm-10 <%=colbs10%> ">
<input type="password" class="form-control" id="passwd2" name="passwd2" placeholder="重复新密码">
<span class="trfont2">不修改请留空</span> </div>
</div>
<div class="form-group">
<label for="trmailstr" class="col-sm-2 <%=colbs2%> control-label">邮箱</label>
<div class="col-sm-10 <%=colbs10%> form-control-static"> <%=ctr_user.mail%></div>
</div>
<div class="trseparateline">以下为选填内容</div>
<div class="form-group">
<label for="qq" class="col-sm-2 <%=colbs2%> control-label">QQ号码</label>
<div class="col-sm-10 <%=colbs10%> ">
<input type="text" class="form-control" id="qq" name="qq" placeholder="QQ号码" value="<%=ctr_user.qq%>">
</div>
</div>
<div class="form-group">
<label for="nickname" class="col-sm-2 <%=colbs2%> control-label">昵称</label>
<div class="col-sm-10 <%=colbs10%> ">
<input type="text" class="form-control" id="nickname" name="nickname" placeholder="昵称" value="<%=ctr_user.nickname%>">
</div>
</div>
<div class="form-group">
<label for="homepage" class="col-sm-2 <%=colbs2%> control-label">主页</label>
<div class="col-sm-10 <%=colbs10%> ">
<input type="text" class="form-control" id="homepage" name="homepage" placeholder="昵称" value="<%=ctr_user.homepage%>">
</div>
</div>
<div class="form-group">
<label for="tel" class="col-sm-2 <%=colbs2%> control-label">电话号码</label>
<div class="col-sm-10 <%=colbs10%> ">
<input type="text" class="form-control" id="tel" name="tel" placeholder="电话号码" value="<%=ctr_user.tel%>">
</div>
</div>
<div class="form-group">
<label for="age" class="col-sm-2 <%=colbs2%> control-label">年龄</label>
<div class="col-sm-10 <%=colbs10%> ">
<input type="text" class="form-control" id="age" name="age" placeholder="年龄" value="<%=ctr_user.age%>">
</div>
</div>
<div class="form-group">
<label for="sex1" class="col-sm-2 <%=colbs2%> control-label">性别</label>
<div class="col-sm-10 <%=colbs10%> ">
<div class="radio">
<label>
<input type="radio" name="sex" id="sex1" value="1" <%if ctr_user.sex=1 then%> checked <%end if%>>
男 </label>
</div>
<div class="radio">
<label>
<input type="radio" name="sex" id="sex2" value="2" <%if ctr_user.sex=2 then%> checked <%end if%>>
女 </label>
</div>
</div>
</div>
<div class="form-group">
<label for="address" class="col-sm-2 <%=colbs2%> control-label">地址</label>
<div class="col-sm-10 <%=colbs10%> ">
<input type="text" class="form-control" id="address" name="address" placeholder="地址" value="<%=ctr_user.address%>">
</div>
</div>
<div class="form-group">
<label for="birthday" class="col-sm-2 <%=colbs2%> control-label">生日</label>
<div class="col-sm-10 <%=colbs10%> ">
<input type="text" class="form-control" id="birthday" name="birthday" placeholder="生日" value="<%=ctr_user.birthday%>">
</div>
</div>
<div class="form-group">
<label for="signature" class="col-sm-2 <%=colbs2%> control-label">签名</label>
<div class="col-sm-10 <%=colbs10%> ">
<textarea name="signature" class="form-control" id="signature" rows="3"> <%=ctr_user.signature%></textarea>
</div>
</div>
<div class="form-group">
<label for="intro" class="col-sm-2 <%=colbs2%> control-label">简介</label>
<div class="col-sm-10 <%=colbs10%> ">
<textarea name="intro" class="form-control" id="intro" rows="3"> <%=ctr_user.intro%></textarea>
</div>
</div>
<%'body8此行代码供安装插件用勿删改%>
<div class="trdivbtn">
<button name="submit1" type="submit" id="submit1" class="btn btn-primary trbtn1" >立即提交</button>
</div>
<input name="action" type="hidden" id="action" value="uedit"/>
<%'body9此行代码供安装插件用勿删改%>
</form>
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
