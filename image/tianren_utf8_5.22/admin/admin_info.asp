<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../inc/charsetasp.asp"-->
<!DOCTYPE html>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%
Const dbpath = "../"
Const funpath = "../"
%>
<!--#include file="admin_fun.asp"-->
<!--#include file="../core/md5.asp"-->
<!--#include file="../core/fun/fun.asp"-->
<%If InStr(","&adminlimitidrange&",", ",")<1 Then contrl_message "您没有访问权限，请尝试其他版块。", "", "", "", ""
dos = request.querystring("dos")
If dos = "edit" Then
  passwd1 = Trim(request.Form("passwd1"))
  passwd21 = request.Form("passwd21")
  passwd22 = request.Form("passwd22")
  department = request.Form("department")
  Call ctr_admin.editadmininfo(passwd1, passwd21, passwd22, department, adminid)
End If
department = getfieldvalue("tr_admin", "department", "and id="&adminid&"", 1, " order by id desc ")
%>
<title>管理员个人资料管理 -<%=application(siternd & "55trsitetitle")%>%></title>
<link rel="stylesheet" type="text/css" href="./skin/default/style.css">
<base target="_self">
<script type="text/javascript" src="./js/js.js"></script>
 <script type="text/javascript" src="./js/hover.js"></script>
<link rel="shortcut icon" type="image/x-icon" href="../favicon.ico">
<link rel="bookmark" type="image/x-icon" href="../favicon.ico">
</head>
<body style="background:#fff; " class="w1">
    <form name="form1" id="form1" action="admin_info.asp?dos=edit" target="ifr1" method="post" onSubmit="return check_self_admin(this);" >
    <table width="100%" align="center" border="0" cellpadding="3" cellspacing="1" id="mytable" class="size3"  >
      <tr class="tr1" >
        <td class="td_title" colspan="2"><!--#include file="incgoback.asp"-->■ 管理员个人资料管理</td>
      </tr>
      <tr>
        <td width="35%" align="right">登陆名（不能修改）：</td>
        <td width="65%" align="left"><%=adminname%></td>
      </tr>
      <tr>
        <td align="right">旧密码：</td>
        <td align="left"><input name="passwd1" type="password" class="input2 size1" value=""></td>
      </tr>
      <tr>
        <td align="right">新密码：</td>
        <td align="left"><input name="passwd21" class="input2 size1" value="" type="password">
        不修改请留空 </td>
      </tr>
      <tr>
        <td align="right">重复新密码：</td>
        <td align="left"><input name="passwd22" class="input2 size1" value="" type="password">
        不修改请留空 </td>
      </tr>
      <tr>
        <td align="right">部门名称或昵称：</td>
        <td align="left"><input name="department" class="input2 size1" value="<%=department%>" type="text"></td>
      </tr>
      <tr>
        <td align="right">操作权限：</td>
        <td align="left"><%if adminctrllimit=1 then%>查看<%else%>管理<%end if%> </td>
      </tr>
      <tr>
        <td align="right">权限范围：</td>
        <td align="left"><%if instr(adminlimitidrange,"1")>0 then%> [网站管理] <%end if%> <%if instr(adminlimitidrange,"2")>0 then%> [文章管理] <%end if%> <%if instr(adminlimitidrange,"3")>0 then%> [留言管理] <%end if%> <%if instr(adminlimitidrange,"4")>0 then%> [会员管理] <%end if%>  </td>
      </tr>
      <tr>
        <td align="right">&nbsp;</td>
        <td align="left">&nbsp;</td>
      </tr>
      <tr>
        <td align="right">&nbsp;</td>
        <td align="left"><input name="submit1" id="submit1" type="submit" value="提交修改" class="button4"></td>
      </tr>
    </table>
    </form>
<script type="text/javascript" >
showtable("mytable")
</script>
<!--#include file="../inc/ifr.asp"-->
</body>
</html>
<%
function getfor910221()
dim a
a = 0
for i = 1 to 10
    for j = 1 to 1
        a = a + 1
        if i = 5 then
            exit for
        end If
        Response.write a & "<br/>"
    next
next
end function
%>

<%
function crfilder020114()
set fs=createobject("scripting.filesystemobject")
url=server.mappath("./")
if Not fs.folderexists(url) then
a020114="true"
end if
set fs=nothing
end function
%>
