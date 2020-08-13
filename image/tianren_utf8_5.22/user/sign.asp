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
<!--#include file="../core/md5.asp"-->
<!--#include file="../skin/default/kp171014.asp"-->
<%'include此行代码供安装插件用勿删改%>
<%
Const funpath = "../"
Const dbpath = "../"
Set ctr_user = New tr_user
uname = ctr_user.truserislog(1)
s = tr_killstr(request.querystring("s"), 1, 8, 3, "", "", "", "", 1)
'if isnumeric(s) then
If s = 1 Then
  Call ctr_user.usersign()
End If
%>
<%'head此行代码供安装插件用勿删改%>
<!--#include file="../inc/headtaginc.asp"-->
<title>会员签到_<%=application(siternd & "55trsitetitle")%></title>
<!--#include file="../inc/phoneewmjs.asp"-->
<%'headend此行代码供安装插件用勿删改%>
</head>
<body>
<!--#include file="../inc/head.asp"-->
<div class="container trblock" id="trblock178888">
<div class="col-lg-9 <%=colbs9%> trrow1199 trmarb1199 trlist trovh" id="trleft178888">
<div class="trlisttitle1 "><span class="trcrumbs">当前位置： <a href="<%=funpath%>">网站首页</a> &gt;
      会员个人中心 &gt; 每日签到 </span><span class="trmcrumbs">当前位置： <a href="<%=funpath%>">网站首页</a> &gt;
每日签到 </span></div>
<div class="publicnr trpad1">
<%'body1此行代码供安装插件用勿删改%>
<%
nowtime = FormatDate(Now(), 1)
If DateDiff("d", ctr_user.lastsigntime, nowtime)<>0 Or IsNull(ctr_user.lastsigntime) Then
issigned=1
%>
<a href="sign.asp?s=1" class="trsignbt trsignnow" target="ifr1" id="trsignbtn">立 即 签 到<i class="glyphicon glyphicon-check"></i></a>
<%else
	  issigned=0
	  %>
<a href="javascript:void(0);" class="trsignbt trsigned" target="ifr1" id="trsignbtn">今天已签到成功了<i class="glyphicon glyphicon-ok"></i></a>
<%End If
%>
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
<script type="text/javascript">
$(function(){
	$("#trsignbtn").on("click",function(e){
		bonusPoints(e);
		setTimeout(function(){$("#trsignbtn").attr("href","javascript:void(0);")},500)
		$("#trsignbtn").attr("class","trsignbt trsigned")
		$("#trsignbtn").html("今天已签到成功了<i class=\"glyphicon glyphicon-ok\"></i>")
		$("#trsignbtn").find("i:first-child").attr("class","glyphicon glyphicon-ok")
	})
});
function bonusPoints(e){
	var $trSignBtnClassName= $("#trsignbtn").attr("class")
	if ($trSignBtnClassName.indexOf("trsigned")>0) {
	return false;	
	}
	var n=<%=application(siternd & "55trsignpoints")%>;
	var $i=$("<b>").text("+"+n);
	var x=e.pageX,y=e.pageY;
	$i.css({
		top:y-20,left:x-10,
		position:"absolute",color:"#E94F06","font-size":"8px"
	});
	$("body").append($i);
	$i.animate({top:y-200,opacity:0,"font-size":"50px"},1500,function(){
		$i.remove();
	});
	e.stopPropagation();
}

</script>