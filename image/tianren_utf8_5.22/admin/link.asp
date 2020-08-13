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
<!--#include file="../core/class/tr_link.asp"-->
 <%
Set ctr_link = New tr_link
If InStr(","&adminlimitidrange&",", ",1,")<1 Then contrl_message "您没有访问权限，请尝试其他版块。", "", "", "", ""
action = Trim(request.Form("action"))
If action<>"" And adminctrllimit = 1 Then Call contrl_message ("您只有查看权限，没有管理权限，请升级权限后继续操作。", "", "", "", "")
Select Case action
  Case "add"
    Call ctr_link.Add()
    Call admin_log(adminname, adminid, adminname&" 添加友情链接成功。")
    Call contrl_message ("添加友情链接成功。", "location.href", "submit1", "form1", "parent")
  Case "edit"
    Call ctr_link.edit()
    Call admin_log(adminname, adminid, adminname&" 修改友情链接成功。")
    Call contrl_message ("修改友情链接成功。", "", "", "", "")
  Case "del"
    Call ctr_link.del()
    Call admin_log(adminname, adminid, adminname&" 删除友情链接成功。")
    Call contrl_message ("删除友情链接成功。", "", "", "", "")
End Select
%>
 <title>友情链接管理<%=application(siternd & "55trsitename")%></title>
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
		<td class="td_title" colspan="7"><!--#include file="incgoback.asp"-->■ 添加友情链接</td>
	</tr>
	<tr class="font1">
		<td width="" align="center">操作</td>
		<td width="60" align="center">状态</td>
		<td width="60" align="center">类型</td>
		<td width="" align="center">链接名称</td>
		<td width="" align="center">链接地址</td>
		<td width="" align="center">链接图片</td>
	</tr>
	<tr class="" valign="middle" >
		<form name="form1" id="form1" target="ifr1" method="post" onSubmit="return check_add_link(this)" >
			<td width="" align="center"><input type="hidden" value="add" name="action" id="action" />
				<input type="submit" value="添加" name="submit1" id="submit1" class="button2"/></td>
			<td width="" align="center"><label>
			  <input name="ispass" type="radio" id="ispass" value="1" checked="checked" />
			  开通</label>
			  <br />
			  <label>
			    <input name="ispass" id="ispass" type="radio" value="0" />
			    关闭</label></td>
			<td width="" align="center"><label>
			  <input name="types" type="radio" id="types" value="1" checked="checked" />
			  文字</label>
			  <br />
			  <label>
			    <input name="types" id="types" type="radio" value="2" />
			    图片</label></td>
			<td width="" align="center"><input name="title" id="title" type="text" class="input2 size2" value="" /></td>
			<td width="" align="center"><input name="homepage" type="text" class="input3 size1" id="homepage" value="" /></td>
			<td width="" align="center"><input name="picfile" id="picfile" type="text" class="input2 size4" value=""/></td>
		</form>
	</tr>
</table>
<br>
<table width="100%" align="center" border="0" cellpadding="3" cellspacing="1" id="mytable2" class="size3" >
	<tr class="tr1" >
		<td class="td_title" colspan="8">■ 编辑友情链接</td>
	</tr>
	<tr class="font1">
		<td width="" align="center">操作</td>
		<td width="" align="center">排序</td>
		<td width="60" align="center">状态</td>
		<td width="60" align="center">类型</td>
		<td align="center">链接名称</td>
		<td align="center">链接地址</td>
		<td width="" align="center">链接图片</td>
		<td width="" align="center">添加时间</td>
	</tr>
  <%
sql = "select * from tr_link where 1=1 and id is not null "
rs.Open sql, conn, 0, 1
If Not rs.EOF Then
  Do While Not rs.EOF

    id = Rs("id")
    orderid = Rs("orderid")
    types = Rs("types")
    title = Rs("title")
    homepage = Rs("homepage")
    'intro = Rs("intro")
    picfile = Rs("picfile")
    addtime = Rs("addtime")
    ispass = Rs("ispass")
%>
	<tr class="" id="tr<%=id%>" valign="middle" >
		<form name="form2<%=id%>" id="form2<%=id%>" target="ifr1" method="post" onSubmit="return check_add_link(this)" >
			<td width="" align="center"><input type="hidden" value="edit" name="action" id="action" />
        <input type="hidden" value="<%=id%>" name="id" id="id" />
        <input type="submit" value="修改" name="submit1" id="submit1" class="button2"/>
        <input type="button" value="删除" name="submit2" id="submit2" class="button3 mar2" onClick="acdel(this.form,'tr<%=id%>')" /></td>
			<td width="" align="center"><input name="orderid" id="orderid" type="text" class="input2 size6" value="<%=orderid%>" /></td>
			<td width="" align="center"><label>
					<input name="ispass" type="radio" id="ispass" value="1" <%if ispass=1 then%>checked="checked"<%end if%> />
					开通</label>
				<br />
				<label>
					<input name="ispass" id="ispass" type="radio" value="0" <%if ispass=0 then%>checked="checked"<%end if%> />
					关闭</label></td>
			<td width="" align="center"><label>
			  <input name="types" type="radio" id="types" value="1" <%if types=1 then%>checked="checked"<%end if%> />
					文字</label>
				<br />
				<label>
				  <input name="types" id="types" type="radio" value="2" <%if types=2 then%>checked="checked"<%end if%> />
			图片</label></td>
			<td width="" align="center"><input name="title" id="title" type="text" class="input2 size2" value="<%=title%>" /></td>
			<td width="" align="center"><input name="homepage" type="text" class="input3 size1" id="homepage" value="<%=homepage%>" /></td>
			<td width="" align="center"><input name="picfile" id="picfile" type="text" class="input2 size4" value="<%=picfile%>"/></td>
			<td width="" align="center"><%=addtime%></td>
		</form>
	</tr>
  <%
rs.movenext
Loop
Else
%>
<tr>
<td colspan="8" align="center">暂无相关项目，请先添加。
</td>
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