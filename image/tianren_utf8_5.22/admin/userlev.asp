<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../inc/charsetasp.asp"-->
<!DOCTYPE html>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<%
Const dbpath = "../"
Const funpath = "../"
%>
<!--#include file="admin_fun.asp"-->
<!--#include file="../core/class/tr_ulev.asp"-->
 <%
Set ctr_ulev = New tr_ulev
If InStr(","&adminlimitidrange&",", ",4,")<1 Then contrl_message "您没有访问权限，请尝试其他版块。", "", "", "", ""
action = Trim(request.Form("action"))
If action<>"" And adminctrllimit = 1 Then Call contrl_message ("您只有查看权限，没有管理权限，请升级权限后继续操作。", "", "", "", "")
Select Case action
  Case "add"
    Call ctr_ulev.Add()
    Call ctr_ulev.create_select(0)
    Call admin_log(adminname, adminid, adminname&" 添加会员等级成功。")
    Call contrl_message ("添加会员等级成功。", "location.href", "submit1", "form1", "parent")
  Case "edit"
    Call ctr_ulev.edit()
    Call ctr_ulev.create_select(0)
    Call admin_log(adminname, adminid, adminname&" 修改会员等级成功。")
    Call contrl_message ("修改会员等级成功。", "", "submit1", "", "")
  Case "del"
    Call ctr_ulev.del()
    Call ctr_ulev.create_select(0)
    Call admin_log(adminname, adminid, adminname&" 删除会员等级成功。")
    Call contrl_message ("删除会员等级成功。", "", "submit1", "", "")
End Select
%>
 <title>会员等级管理<%=application(siternd & "55trsitename")%></title>
 <link rel="stylesheet" type="text/css" href="./skin/default/style.css">
 <base target="_self">
 <script type="text/javascript" src="./js/js.js"></script>
 <script type="text/javascript" src="./js/hover.js"></script>
 <script type="text/javascript" src="../js/JTimer.js"></script>
<link rel="shortcut icon" type="image/x-icon" href="../favicon.ico">
<link rel="bookmark" type="image/x-icon" href="../favicon.ico">
 </head>
 <body style="background:#fff; " class="w1">
<table width="100%" align="center" border="0" cellpadding="3" cellspacing="1" id="mytable" class="size3" >
   <tr class="tr1" >
    <td class="td_title" colspan="11"><!--#include file="incgoback.asp"-->■ 添加等级</td>
  </tr>
   <tr class="font1">
    <td width="" align="center">操作</td>
    <td width="" align="center">等级名称</td>
    <td width="" align="center">等级积分</td>
  </tr>
   <tr class="">
    <form name="form1" id="form1" target="ifr1" method="post" onSubmit="return check_add_ulev(this)" >
       <td width="" align="center"><input type="hidden" value="add" name="action" id="action" />
        <input type="submit" value=" 添加 " name="submit1" id="submit1" class="button2"/></td>
       <td width="" align="center"><input name="levname" id="levname" type="text" class="input2 size19" value="" /></td>
       <td width="" align="center"><input name="userjf" id="userjf" type="text" class="input2 size12" value="" /></td>
     </form>
  </tr>
 </table>
<br>
<table width="100%" align="center" border="0" cellpadding="3" cellspacing="1" id="mytable2" class="size3" >
   <tr class="tr1" >
    <td class="td_title" colspan="12">■ 编辑等级信息</td>
  </tr>
   <tr class="font1">
    <td width="" align="center">操作</td>
    <td width="" align="center">等级名称</td>
    <td width="" align="center">等级积分</td>
  </tr>
   <%
sql = "select * from tr_ulev where id is not null order by userjf asc , id asc "
rs.Open sql, conn, 0, 1
If Not rs.EOF Then
  Do While Not rs.EOF
    id = Rs("id")
    userjf = Rs("userjf")
    levname = Rs("levname")
    addtime = Rs("addtime")
%>
   <tr class="" id="tr<%=id%>">
    <form name="form2<%=id%>" id="form2<%=id%>" target="ifr1" method="post" onSubmit="return check_add_ulev(this)" >
       <td width="" align="center"><input type="hidden" value="edit" name="action" id="action" />
        <input type="hidden" value="<%=id%>" name="id" id="id" />
        <input type="submit" value=" 修改 " name="submit1" id="submit1" class="button2"/>
        <input type="button" value=" 删除 " name="submit2" id="submit2" class="button3 mar2" onClick="acdel(this.form,'tr<%=id%>')"/></td>
       <td width="" align="center"><input name="levname" id="levname" type="text" class="input2 size19" value="<%=levname%>" /></td>
       <td width="" align="center"><input name="userjf" id="userjf" type="text" class="input2 size12" value="<%=userjf%>"/></td>
     </form>
  </tr>
<%
rs.movenext
Loop
%>
   <tr>
    <td colspan="12" align="left">特别提示：<br>会员积分等级需要由上至下依次增大，不能出现依次减少或大小穿插的情况，否则会员等级与前台文章阅读权限可能会出错。 </td>
  </tr>
<%
Else
%>
   <tr>
    <td colspan="12" align="center">暂无相关项目，请先添加。 </td>
  </tr>
   <%end if%>
 </table>
<script type="text/javascript" >
showtable("mytable")
showtable("mytable2")
</script>
<!--#include file="../inc/ifr.asp"-->
</body>
</html>
<%'勿删改此行%>
