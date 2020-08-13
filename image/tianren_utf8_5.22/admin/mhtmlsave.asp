<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<%
Server.ScriptTimeOut = 36000
Const dbpath = "../"
Const funpath = "../"
%>
<!--#include file="admin_fun.asp"-->
<!--#include file="mhtmlinc.asp"-->
<%
if instr(","&adminlimitidrange&",",",2,")<1 then contrl_message "您没有访问权限，请尝试其他版块。", "", "", "", ""
%>
<script>
function fs(s,cs)
{
//alert (Math.ceil(s*100/cs)-100);
if (s<=cs){
document.getElementById('contents').style.backgroundPosition=(Math.ceil(s*100/cs)-100)*5+"px center"
document.getElementById('titles').value=Math.ceil(s*100/cs)+"%"
document.getElementById('counts').value=cs+" 个文件"
if (Math.ceil(cs-s/1.2)<=60){
document.getElementById('times').value=Math.ceil((cs-s)/1.2)+" 秒"
}
else{
document.getElementById('times').value=Math.ceil((cs-s)/72)+" 分钟"
}
if (Math.ceil(s*100/cs)>98){
	document.getElementById('times').value="生成完毕"
}
}
}
</script>
<title>生成静态页面<%=application(siternd & "55trsitename")%></title>
<link rel="shortcut icon" type="image/x-icon" href="../favicon.ico">
<link rel="bookmark" type="image/x-icon" href="../favicon.ico">
</head>

<body>
<div id="contents" style="height:30px; border:0px solid #FF0000; background:url(skin/default/img/progress.jpg) no-repeat -500px center #E9FFFF; line-height:30px; width:500px; font-size:12px" ><span style="margin-left:20px; color:#333;" >进度：
  <input type="text" name="titles" id="titles" value="0%" style="width:50px; border:0px;font-weight:bold; color:#333; background:none;"  />
  &nbsp;共：
  <input type="text" name="counts" id="counts" value="0" style="width:100px; border:0px; text-align:left; font-weight:bold; color:#333;  background:none;" />
  &nbsp; &nbsp; &nbsp;预计耗时：
  <input type="text" name="times" id="times" value="0" style="width:100px; border:0px; text-align:left; font-weight:bold; color:#333;  background:none;" />
  </span></div>
</body>
</html>
<%
action = request.Form("action")
If action<>"" And adminctrllimit = 1 Then Call contrl_message ("您只有查看权限，没有管理权限，请升级权限后继续操作。", "", "", "", "")
Select Case action
  Case "default"
    Call mdefaulthtml()
  Case "list"
    Call mlisthtml("")
  Case "show"
    Call mshowhtml("")
End Select
set ctr_admin=nothing
%>
