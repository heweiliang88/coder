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
<!--#include file="../core/class/tr_ads.asp"-->
 <%
Set ctr_ads = New tr_ads
If InStr(","&adminlimitidrange&",", ",1,")<1 Then contrl_message "您没有访问权限，请尝试其他版块。", "", "", "", ""
action = Trim(request.Form("action"))
If action<>"" And adminctrllimit = 1 Then Call contrl_message ("您只有查看权限，没有管理权限，请升级权限后继续操作。", "", "", "", "")
Select Case action
  Case "add"
    Call ctr_ads.Add()
    Call ctr_ads.createads()
    Call admin_log(adminname, adminid, adminname&" 添加广告成功。")
    Call contrl_message ("添加广告成功。", "location.href", "submit1", "form1", "parent")
  Case "edit"
    Call ctr_ads.edit()
    Call ctr_ads.createads()
    Call admin_log(adminname, adminid, adminname&" 修改广告成功。")
    Call contrl_message ("修改广告成功。", "", "submit1", "", "")
  Case "del"
    Call ctr_ads.del()
    Call ctr_ads.createads()
    Call admin_log(adminname, adminid, adminname&" 删除广告成功。")
    Call contrl_message ("删除广告成功。", "", "submit1", "", "")
End Select
%>
 <title>网站广告管理<%=application(siternd & "55trsitename")%></title>
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
		<td class="td_title" colspan="7"><!--#include file="incgoback.asp"-->■ 添加网站广告</td>
	</tr>
	<tr class="font1">
		<td width="" align="center">操作</td>
		<td width="" align="center">状态</td>
		<td width="" align="center">广告名称</td>
		<td width="" align="center">广告代码</td>
		<td width="" align="center">到期时间</td>
	</tr>
	<tr class="" valign="middle" >
		<form name="form1" id="form1" target="ifr1" method="post" onSubmit="return check_add_ads(this)" >
			<td width="" align="center"><input type="hidden" value="add" name="action" id="action">
				<input type="submit" value="添加" name="submit1" id="submit1" class="button2"></td>
			<td width="" align="center"><label>
					<input name="ispass" type="radio" id="ispass" value="1" checked="checked">
					开通</label>
				<br>
				<label>
					<input name="ispass" id="ispass" type="radio" value="0">
					关闭</label></td>
			<td width="" align="center"><input name="title" id="title" type="text" class="input2 size2" value=""></td>
			<td width="" align="center"><textarea name="content" class="input3 size14" id="content"></textarea></td>
			<td width="" align="center"><input name="endtime" id="endtime" type="text" class="input2 size4" value="" onClick="JTC.setday({startDay: '<%=FormatDate(now(), 2)%>'})"></td>
		</form>
	</tr>
</table>
<br>
<table width="100%" align="center" border="0" cellpadding="3" cellspacing="1" id="mytable2" class="size3" >
	<tr class="tr1" >
		<td class="td_title" colspan="7">■ 编辑网站广告</td>
	</tr>
	<tr class="font1">
		<td width="" align="center">操作</td>
		<td width="" align="center">排序</td>
		<td width="50" align="center">状态</td>
		<td width="" align="center">广告名称</td>
		<td width="" align="center">广告代码</td>
		<td width="" align="center">到期时间</td>
		<td width="" align="center">调用代码</td>
	</tr>
  <%
sql = "select * from tr_ads where 1=1 and id is not null order by orderid desc ,callid desc "
rs.Open sql, conn, 0, 1
If Not rs.EOF Then
  Do While Not rs.EOF
    id = Rs("id")
    orderid = Rs("orderid")
    'addtime=Rs("addtime")
    'types=Rs("types")
    content = Rs("content")
    'clicks=Rs("clicks")
    callid = Rs("callid")
    title = Rs("title")
    ispass = Rs("ispass")
    endtime = Rs("endtime")
%>
	<tr class="" id="tr<%=id%>" valign="middle" >
		<form name="form2<%=id%>" id="form2<%=id%>" target="ifr1" method="post" onSubmit="return check_add_ads(this)" >
			<td width="" align="center"><input type="hidden" value="edit" name="action" id="action">
        <input type="hidden" value="<%=id%>" name="id" id="id">
        <input type="submit" value="修改" name="submit1" id="submit1" class="button2">
        <br>
        <input type="submit" value="删除" name="submit2" id="submit2" class="button3 mar2"  onclick="acdel(this.form,'tr<%=id%>')"></td>
			<td width="" align="center"><input name="orderid" id="orderid" type="text" class="input2 size6" value="<%=orderid%>"></td>
			<td width="" align="center"><label>
					<input name="ispass" type="radio" id="ispass" value="1" <%if ispass=1 then%>checked="checked"<%end if%>>
					开通</label>
				<br>
				<label>
					<input name="ispass" id="ispass" type="radio" value="0" <%if ispass=0 then%>checked="checked"<%end if%>>
					关闭</label></td>
			<td width="" align="center"><input name="title" id="title" type="text" class="input2 size2" value="<%=title%>"></td>
			<td width="" align="center"><textarea name="content" class="input3 size14" id="content"><%=content%></textarea></td>
			<td width="" align="center"><input name="endtime" id="endtime" type="text" class="input2 size4" value="<%=endtime%>" onClick="JTC.setday({startDay: '<%=endtime%>'})"></td>
			<td width="" align="center"><div class="input3 size20" style="word-wrap:break-word; word-break:break-all;" onClick="selectText911030('dydmdiv911030<%=id%>')" id="dydmdiv911030<%=id%>"><%=Server.HTMLEncode("<script type=""text/javascript""> _55tr_com('tr"&callid&"')</script>")%></div>
</td>
		</form>
	</tr>
  <%
rs.movenext
Loop
%>
<tr>
<td colspan="7" align="left"><p>小提示：上面的广告位在没有修改前台模板时请不要删除，因为上面的每个广告位都与前台的广告位一一对应，如果删除那么前台广告就无法关闭或正常显示。</p>
<p>“调用代码”是程序自动生成的，无法修改，是用来给有html基础的用户在前台添加或修改广告位时在前台调用用的。</p>
<p>天人文章管理系统设有站内图片及链接dns功能，直白说就是，各个投放广告的页面路径不同，文件层级不同。</p>
  <p>单纯的使用相对路径无法保证每个页面都能够正常显示站内的相对路径广告图片与打开站内的相对路径广告链接。现在，只需要在链接的href=""中的第一个引号后面添加trgoddns即可解决相对路径的问题。</p>
  <p>例如：&lt;a href=&quot;show.asp?id=1&quot;&gt;文章&lt;/a&gt;，改为&lt;a href=&quot;trgoddnsshow.asp?id=1&quot;&gt;文章&lt;/a&gt;即可。</p>
  <p>同理，对于图片类型的广告位同样如此，例如：&lt;img src=&quot;upfile/image/202002/pic.jpg&quot;/&gt;，改为&lt;img src=&quot;trgoddnsupfile/image/202002/pic.jpg&quot;/&gt;即可。</p>
  <p>同样，对于图片链接&lt;a href=&quot;show.asp?id=1&quot;&gt;&lt;img src=&quot;upfile/image/202002/pic.jpg&quot;/&gt;&lt;/a&gt;，也可以改为&lt;a href=&quot;trgoddnsshow.asp?id=1&quot;&gt;&lt;img src=&quot;trgoddnsupfile/image/202002/pic.jpg&quot;/&gt;&lt;/a&gt;</p>
  <p>注意，此设置只针对相对路径的站内图片与链接，对于外部网站的绝对路径无需这样设置。 </p></td>
</tr>	
<%
Else
%>
<tr>
<td colspan="7" align="center">暂无相关项目，请先添加。
</td>
</tr>	
	<%end if%>
</table>
<script type="text/javascript" >
showtable("mytable")
showtable("mytable2")
</script>
<script type="text/javascript">
function selectText911030(containerid) {
if (document.selection) {
var range = document.body.createTextRange();
range.moveToElementText(document.getElementById(containerid));
range.select();
} else if (window.getSelection) {
var range = document.createRange();
range.selectNode(document.getElementById(containerid));
window.getSelection().removeAllRanges();
window.getSelection().addRange(range);
}
}
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
