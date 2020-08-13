<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../inc/charsetasp.asp"-->
<!DOCTYPE html>
<html>
 <head>
 <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
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
<!--#include file="../core/class/tr_column.asp"-->
 <%
Set ctr_column = New tr_column
If InStr(","&adminlimitidrange&",", ",2,")<1 Then contrl_message "您没有访问权限，请尝试其他版块。", "", "", "", ""
columnid = Trim(request.querystring("columnid"))
If columnid = "" Then columnid = 0
If columnid>0 Then
  queueid = getfieldvalue("tr_column", "queueid", " and id="&columnid&"", 1, " order by id desc ")
Else
  queueid = ",0,"
End If
action = Trim(request.Form("action"))
If action<>"" And adminctrllimit = 1 Then Call contrl_message ("您只有查看权限，没有管理权限，请升级权限后继续操作。", "", "", "", "")
Select Case action
  Case "add"
    Call ctr_column.Add()
    Call ctr_column.cr()
    Call admin_log(adminname, adminid, adminname&" 添加栏目成功。")
    Call contrl_message ("添加栏目成功。", "location.href", "submit1", "form1", "parent")
  Case "edit"
    Call ctr_column.edit()
    Call ctr_column.cr()
    Call admin_log(adminname, adminid, adminname&" 修改栏目成功。")
    Call contrl_message ("修改栏目成功。", "", "submit1", "", "")
  Case "del"
    Call ctr_column.del()
    Call ctr_column.cr()
    Call admin_log(adminname, adminid, adminname&" 删除栏目成功。")
    Call contrl_message ("删除栏目成功。", "", "submit1", "", "")
End Select
%>
 <title>网站栏目管理<%=application(siternd & "55trsitename")%></title>
 <link rel="stylesheet" type="text/css" href="./skin/default/style.css">
 <base target="_self">
 <script type="text/javascript" src="./js/js.js"></script>
 <script type="text/javascript" src="./js/hover.js"></script>
 <script type="text/javascript" src="../js/jtimer.js"></script>
 <link rel="shortcut icon" type="image/x-icon" href="../favicon.ico">
<link rel="bookmark" type="image/x-icon" href="../favicon.ico">
</head>
 <body style="background:#fff; " class="w1">
<table width="100%" align="center" border="0" cellpadding="3" cellspacing="1" id="mytable" class="size3" >
<tr class="tr1" >
		<td class="td_title" colspan="13"><!--#include file="incgoback.asp"-->■ 添加栏目 &nbsp;&nbsp;|&nbsp;&nbsp;&nbsp;<%call getcolumn_url(columnid)%></td>
	</tr>
	<tr class="font1">
		<td width="" align="center">操作</td>
		<td width="60" align="center">状态</td>
		<td width="60" align="center">导航</td>
		<td width="60" align="center">首页</td>
		<td width="60" align="center">类型</td>
		<td width="60" align="center">图文模块</td>
		<td width="" align="center">栏目名称</td>
		<td width="" align="center">每页容量</td>
		<td width="" align="center">跳转链接</td>
		<td width="" align="center">关键词</td>
		<td width="" align="center">描述</td>
	</tr>
	<tr class="">
		<form name="form1" id="form1" target="ifr1" method="post" onSubmit="return check_add_column(this)" >
			<td width="" align="center" valign="middle" ><input type="hidden" value="add" name="action" id="action" /><input type="hidden" value="<%=columnid%>" name="columnid" id="columnid" /><input type="hidden" value="<%=queueid%>" name="queueid" id="queueid" />
				<input type="submit" value="添加" name="submit1" id="submit1" class="button2"/></td>
			<td width="" align="center"><label>
			  <input name="isopen" type="radio" id="isopen" value="1" checked="checked" />
			  开通</label>
			  <br />
			  <label>
			    <input name="isopen" id="isopen" type="radio" value="0" />
			    关闭</label></td>
			<td width="" align="center"><label>
			  <input name="isnav" type="radio" id="isnav" value="1" checked="checked" />
			  是</label>
			  <br />
			  <label>
			    <input name="isnav" id="isnav" type="radio" value="0" />
			    否</label></td>
			<td width="" align="center"><label>
			  <input name="isindex" type="radio" id="isindex" value="1" checked="checked" />
			  是</label>
			  <br />
			  <label>
			    <input name="isindex" id="isindex" type="radio" value="0" />
			    否</label></td>
<td width="" align="center"><label>
			  <input name="types" type="radio" id="types" value="1" checked="checked" />
			  文字</label>
			  <br />
			  <label>
			    <input name="types" id="types" type="radio" value="2" />
			    图片</label></td>
<td width="" align="center"><label>
			  <input name="isimgtext" type="radio" id="isimgtext" value="1" checked="checked" />
			  是</label>
			  <br />
			  <label>
			    <input name="isimgtext" id="isimgtext" type="radio" value="0" />
			    否</label></td>
			<td width="" align="center"><textarea name="colname" class="input3 size16" id="colname"></textarea></td>
			<td width="" align="center"><input name="tpagesize" id="tpagesize" type="text" class="input2 size6" value="20" /></td>
			<td width="" align="center"><textarea name="jumpurl" class="input3 size16" id="jumpurl"></textarea></td>
			<td width="" align="center"><textarea name="keywords" class="input3 size11" id="keywords"></textarea></td>
			<td width="" align="center"><textarea name="descriptions" class="input3 size11" id="descriptions"></textarea></td>
	  </form>
	</tr>
</table>
<br>
<table width="100%" align="center" border="0" cellpadding="3" cellspacing="1" id="mytable2" class="size3" >
	<tr class="tr1" >
		<td class="td_title" colspan="13">■ 编辑栏目信息</td>
	</tr>
	<tr class="font1">
		<td width="" align="center">操作</td>
		<td width="" align="center">排序</td>
		<td width="60" align="center">状态</td>
		<td width="45" align="center">导航</td>
		<td width="45" align="center">首页</td>
		<td width="60" align="center">类型</td>
		<td width="60" align="center">图文模块</td>
		<td width="40" align="center">下级栏目</td>
		<td width="" align="center">栏目名称</td>
		<td width="" align="center">每页容量</td>
		<td width="" align="center">跳转链接</td>
		<td width="" align="center">关键词</td>
		<td width="" align="center">描述</td>
	</tr>
  <%
andstr = " and columnid="&columnid&" "
sql = "select * from tr_column where 1=1 "&andstr&" and id is not null order by orderid asc , id asc "
rs.Open sql, conn, 0, 1
If Not rs.EOF Then
  Do While Not rs.EOF
    id = Rs("id")
    colname = Rs("colname")
    keywords = Rs("keywords")
    descriptions = Rs("descriptions")
    intro = Rs("intro")
    types = Rs("types")
    addtime = Rs("addtime")
    orderid = Rs("orderid")
    queueid = Rs("queueid")
    columnid = Rs("columnid")
    jumpurl = Rs("jumpurl")
    isnav = Rs("isnav")
    isindex = Rs("isindex")
    islist = Rs("islist")
    isshow = Rs("isshow")
    isopen = Rs("isopen")
    tpagesize = Rs("tpagesize")
	isimgtext = rs("isimgtext")
%>
	<tr class="" id="tr<%=id%>" valign="middle" >
		<form name="form2<%=id%>" id="form2<%=id%>" target="ifr1" method="post" onSubmit="return check_add_column(this)" >
			<td width="" align="center"><input type="hidden" value="edit" name="action" id="action" />
        <input type="hidden" value="<%=id%>" name="id" id="id" /><input type="hidden" value="<%=columnid%>" name="columnid" id="columnid" /><input type="hidden" value="<%=queueid%>" name="queueid" id="queueid" />
        <input type="submit" value="修改" name="submit1" id="submit1" class="button2"/>
        <input type="button" value="删除" name="submit2" id="submit2" class="button3 mar2" onClick="acdel(this.form,'tr<%=id%>')"/></td>
			<td width="" align="center"><input name="orderid" id="orderid" type="text" class="input2 size13" value="<%=orderid%>" /><br>ID:<%=rs("id")%></td>
			<td width="" align="center"><label>
					<input name="isopen" type="radio" id="isopen" value="1" <%if isopen=1 then%>checked="checked"<%end if%> />
					开通</label>
				<br>
				<label>
					<input name="isopen" id="isopen" type="radio" value="0" <%if isopen=0 then%>checked="checked"<%end if%> />
					关闭</label></td>
			<td width="" align="center"><label>
			  <input name="isnav" type="radio" id="isnav" value="1" <%if isnav=1 then%>checked="checked"<%end if%> />
					是</label>
				<br />
				<label>
				  <input name="isnav" id="isnav" type="radio" value="0" <%if isnav=0 then%>checked="checked"<%end if%> />
					否</label></td>
			<td width="" align="center"><label>
			  <input name="isindex" type="radio" id="isindex" value="1" <%if isindex=1 then%>checked="checked"<%end if%> />
					是</label>
				<br />
				<label>
				  <input name="isindex" id="isindex" type="radio" value="0" <%if isindex=0 then%>checked="checked"<%end if%> />
					否</label></td>
    <td width="" align="center"><label>
			  <input name="types" type="radio" id="types" value="1" <%if types=1 then%>checked="checked"<%end if%> />
					文字</label>
				<br />
				<label>
				  <input name="types" id="types" type="radio" value="2" <%if types=2 then%>checked="checked"<%end if%> />
			图片</label></td>
			<td width="" align="center"><label>
			  <input name="isimgtext" type="radio" id="isimgtext" value="1" <%if isimgtext=1 then%>checked="checked"<%end if%> />
					是</label>
				<br />
				<label>
				  <input name="isimgtext" id="isimgtext" type="radio" value="0" <%if isimgtext=0 then%>checked="checked"<%end if%> />
					否</label></td>
			<td width="" align="center"><%= getnextcolumnname(id)%></td>
			<td width="" align="center"><textarea name="colname" class="input3 size16" id="colname"><%=colname%></textarea></td>
			<td width="" align="center"><input name="tpagesize" type="text" class="input3 size6" id="tpagesize" value="<%=tpagesize%>" /></td>
			<td width="" align="center"><textarea name="jumpurl" class="input3 size16" id="jumpurl"><%=jumpurl%></textarea></td>
			<td width="" align="center"><textarea name="keywords" class="input3 size11" id="keywords"><%=keywords%></textarea></td>
			<td width="" align="center"><textarea name="descriptions" class="input3 size11" id="descriptions"><%=descriptions%></textarea></td>
		</form>
	</tr>
  <%
rs.movenext
Loop
Else
%>
<tr>
<td colspan="13" align="center">暂无相关项目，请先添加。
</td>
</tr>	
<%end if%>
<tr>
<td colspan="13" align="left"><p>小提示：<br>  
  1、当栏目存在下级栏目或栏目下存在文章时需要先删除下级栏目与删除或转移文章方能删除该栏目。</p>
  <p>2、想在前台导航栏与首页栏目板块中调整栏目的顺序，可使用上面的排序功能，排序的数字越大排名越靠后。<br>
  </p>
</td>
</tr>	
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
