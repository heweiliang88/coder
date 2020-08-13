<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../inc/charsetasp.asp"-->
<!DOCTYPE html>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%
'版权声明：
'天人文章管理系统，以下简称“本程序、程序”可免费用于个人非商业用途，其他用途需付费购买版权，请按照如下约定执行：
'**本程序可免费用于个人非商业用途，除个人非商业外的用途均需要付费购买版权，并在官网登记域名作为正规版权记录，并可获得电子版版权证书(不是绑定域名，而是在官网数据库中登记域名旨在法律维权时避开已购买版权的域名)，侵权使用必将承担法律风险并追究高额赔偿
'**购买版权将获得稳定健全的升级服务、免费安装插件、模板、配套工具免费使用、官方优惠优先获得等特权。并可用于企业、政府、校园、单位、团体、个人、机构、部门、机关、公司及其他商业用途
'**购买版权后，除源代码中版权声明，后台左下角的“源程序：55tr.com”字样（这两种版权字样都不会在前台显示，只作为程序版权归属）以外所有代码都可以修改及删除，包括前台页面底部的官网链接及版权等代码都可修改。
'**购买版权后可用本程序为别人或自己提供建设网站，修改源码，二次开发的服务，但禁止发布或传播本程序及本程序的任何形态！禁止出售本程序及本程序的任何形态。说白了就是不能将程序出售或传播，否则侵权。（例如：小李与客户约定使用本程序制作某中学的官网，并收取某中学的网站制作服务费用1万元，并购买版权，此行为不侵权。小赵与客户约定使用本程序制作某小学的官网，并收取某小学的网站制作服务费用1万元，未购买版权，此行为侵权(因为未购买版权)，请终止！小孙使用本程序建立了个人漫画网站供用户浏览并获取广告收益，未购买版权，此行为不侵权。小张使用本程序制作为企业网站进行出售，并购买版权，此行为侵权（因为传播出售了程序），请终止！小王又将本程序制作为小说网站源码在网络免费发布传播提供下载，并购买版权，此行为侵权（因为传播出售了程序），请终止！）
'**不购买版权的个人非商业用途依然享受升级、安装插件、模板的权利，可修改源代码，但不得修改源代码中“官网链接”、“版权声明”、“应用中心”相关的代码，前台页面底部的“Copyright 2016 天人文章管理系统 版权所有，授权xxx使用 Powered by 55TR.COM”字样不得删除,否则请购买版权
'**不购买版权的除个人非商业外的用途请及时购买版权或终止使用本程序，否则侵权
'**计算机软件著作权登记证书登记号：2016SR204636，侵权必究，作者QQ81962480 欢迎洽谈基于此程序开发、修改业务。
'**版权声明会存在修改的偶然性，请以官网公布的对应程序版本版权声明为准！具体的版权声明、特权、及功能介绍，请浏览：http://www.55tr.com/gwappshow.asp?id=11
'**天人文章管理系统3.48之前的程序版本版权声明以此为准，之前的版权声明作废


Const dbpath = "../"
Const funpath = "../"
%>
<!--#include file="admin_fun.asp"-->
<!--#include file="../core/md5.asp"-->
<%
Set ctr_admin = New tr_admin
If InStr(","&adminlimitidrange&",", ",1,")<1 Then contrl_message "您没有访问权限，请尝试其他版块。", "", "", "", ""
action = Trim(request.Form("action"))
If action<>"" And adminctrllimit = 1 Then Call contrl_message ("您只有查看权限，没有管理权限，请升级权限后继续操作。", "", "", "", "")
Select Case action
  Case "add"
    Call ctr_admin.add_admin()
    Call admin_log(adminname, adminid, adminname&" 添加管理员成功。")
    Call contrl_message ("新增管理员成功。", "location.href", "submit1", "form1", "parent")
  Case "edit"
    Call ctr_admin.edit_admin()
    Call admin_log(adminname, adminid, adminname&" 修改管理员成功。")
    Call contrl_message ("管理员资料修改成功。", "", "submit1", "", "")
  Case "del"
    Call ctr_admin.del_admin(adminid)
    Call admin_log(adminname, adminid, adminname&" 删除管理员成功。")
    Call contrl_message ("管理员删除成功。", "", "submit1", "", "")
End Select
%>
<title>管理员管理<%=application(siternd & "55trsitename")%></title>
<link rel="stylesheet" type="text/css" href="./skin/default/style.css">
<base target="_self">
<script type="text/javascript" src="./js/js.js"></script>
<script type="text/javascript" src="./js/hover.js"></script>
<link rel="shortcut icon" type="image/x-icon" href="../favicon.ico">
<link rel="bookmark" type="image/x-icon" href="../favicon.ico">
</head>
<body style="background:#fff; " class="w1">
<table width="100%" align="center" border="0" cellpadding="3" cellspacing="1" id="mytable" class="size3" >
	<tr class="tr1" >
		<td class="td_title" colspan="17"><!--#include file="incgoback.asp"-->■ 添加管理员
</td>
	</tr>
	<tr class="font1">
		<td width="" align="center">操作</td>
		<td width="" align="center">登录名</td>
		<td width="" align="center">密码</td>
		<td width="" align="center">重复密码</td>
		<td width="" align="center">部门或昵称</td>
		<td width="" align="center">是否解冻</td>
		<td width="" align="center">操作权限</td>
		<td width="" align="center">权限范围</td>
	</tr>
	<tr>
		<form name="form1" id="form1" target="ifr1" method="post" onSubmit="return check_add_admin(this)" >
			<td width="" align="center"><input type="hidden" value="add" name="action" id="action">
				<input type="submit" value=" 添加 " name="submit1" id="submit1" class="button2"></td>
			<td width="" align="center"><input name="adminname" id="adminname" type="text" class="input2 size4" value=""></td>
			<td width="" align="center"><input name="passwd1" id="passwd1" type="password" class="input2 size4" value=""></td>
			<td width="" align="center"><input name="passwd2" id="passwd2" type="password" class="input2 size4" value=""></td>
			<td width="" align="center"><input name="department" id="department" type="text" class="input2 size4" value=""></td>
			<td width=""><label>
					<input name="isfree" type="radio" id="isfree" value="1" checked="checked">
					解冻</label>
				<br>
				<label>
					<input name="isfree" id="isfree" type="radio" value="0">
					冻结</label></td>
			<td width=""><label>
					<input name="ctrllimit" type="radio" id="ctrllimit" value="2" checked="checked">
					管理</label>
				<br>
				<label>
					<input name="ctrllimit" id="ctrllimit" type="radio" value="1">
					查看</label></td>
			<td width=""><label>
					<input name="limitidrange" id="limitidrange" type="checkbox" value="1">
					网站管理</label>
				<label>
					<input name="limitidrange" id="limitidrange" type="checkbox" value="2">
					文章管理</label>
				<br>
				<label>
					<input name="limitidrange" id="limitidrange" type="checkbox" value="5">
					采集管理</label>
				<label>
					<input name="limitidrange" id="limitidrange" type="checkbox" value="3">
					留言管理</label>
				<br>
				<label>
					<input name="limitidrange" id="limitidrange" type="checkbox" value="4">
					会员管理</label>					<input name="limitidrange" id="limitidrange" type="checkbox" value="6">
					应用管理</label></td>
		</form>
	</tr>
	<tr class="tr1" >
		<td class="td_title" colspan="17">■ 管理员管理</td>
	</tr>
	<tr class="font1">
		<td width="" align="center">操作</td>
		<td width="" align="center">登录名</td>
		<td width="" align="center">密码</td>
		<td width="" align="center">重复密码</td>
		<td width="" align="center">部门或昵称</td>
		<td width="" align="center">是否解冻</td>
		<td width="" align="center">操作权限</td>
		<td width="" align="center">权限范围</td>
	</tr>
	<%
sql = "select * from tr_admin where 1=1 and id is not null "
rs.Open sql, conn, 0, 1
Do While Not rs.EOF
  id = Rs("id")
  adminname = Rs("adminname")
  isfree = Rs("isfree")
  ctrllimit = Rs("ctrllimit")
  limitidrange = Rs("limitidrange")
  department = Rs("department")
%>
	<tr id="tr<%=id%>">
		<form action="" target="ifr1" name="form2<%=id%>" id="form2<%=id%>" method="post" onSubmit="return check_edit_admin(this)" >
			<td width="" align="center"><input type="hidden" value="edit" name="action" id="action">
				<input type="hidden" value="<%=id%>" name="id" id="id">
				<input type="submit" value=" 修改 " name="submit1" id="submit1" class="button2">
				<br>
				<input type="button" value=" 删除 " name="submit2" id="submit2" class="button3 mar2" onClick="acdel(this.form,'tr<%=id%>')"></td>
			<td width="" align="center"><input name="adminname" type="text" class="input2 size4" id="adminname" value="<%=adminname%>" readonly></td>
			<td width="" align="center"><input name="passwd1" id="passwd1" type="password" class="input2 size4" value=""></td>
			<td width="" align="center"><input name="passwd2" id="passwd2" type="password" class="input2 size4" value=""></td>
			<td width="" align="center"><input name="department" id="department" type="text" class="input2 size4" value="<%=department%>"></td>
			<td width=""><label>
					<input name="isfree" type="radio" id="isfree" value="1" <%if isfree=1 then%> checked="checked"<%end if%>>
					解冻</label>
				<br>
				<label>
					<input name="isfree" id="isfree" type="radio" value="0" <%if isfree=0 then%> checked="checked"<%end if%>>
					冻结</label></td>
			<td width=""><label>
					<input name="ctrllimit" type="radio" id="ctrllimit" value="2" <%if ctrllimit=2 then%> checked="checked"<%end if%>>
					管理</label>
				<br>
				<label>
					<input name="ctrllimit" id="ctrllimit" type="radio" value="1" <%if ctrllimit=1 then%> checked="checked"<%end if%>>
					查看</label></td>
			<td width=""><label>
					<input name="limitidrange" id="limitidrange" type="checkbox" value="1" <%if instr(limitidrange,"1")>0 then%> checked="checked"<%end if%>>
					网站管理</label>
				<label>
					<input name="limitidrange" id="limitidrange" type="checkbox" value="2" <%if instr(limitidrange,"2")>0 then%> checked="checked"<%end if%>>
					文章管理</label>
				<br>
				<label>
					<input name="limitidrange" id="limitidrange" type="checkbox" value="5" <%if instr(limitidrange,"5")>0 then%> checked="checked"<%end if%>>
					采集管理</label>
				<label>
					<input name="limitidrange" id="limitidrange" type="checkbox" value="3" <%if instr(limitidrange,"3")>0 then%> checked="checked"<%end if%>>
					留言管理</label>
				<br>
				<label>
					<input name="limitidrange" id="limitidrange" type="checkbox" value="4" <%if instr(limitidrange,"4")>0 then%> checked="checked"<%end if%>>
					会员管理</label>					<input name="limitidrange" id="limitidrange" type="checkbox" value="6" <%if instr(limitidrange,"6")>0 then%> checked="checked"<%end if%>>
					应用管理</label></td>
		</form>
	</tr>
	<%
rs.movenext
Loop
%>
	<tr>
		<td align="right">&nbsp;</td>
		<td colspan="15">&nbsp;</td>
	</tr>
</table>
</form>
<script type="text/javascript" >
showtable("mytable")
</script>
<!--#include file="../inc/ifr.asp"-->
</body>
</html>
<%'勿删改%>

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
