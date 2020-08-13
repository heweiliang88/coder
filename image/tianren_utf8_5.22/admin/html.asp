<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../inc/charsetasp.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
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
if instr(","&adminlimitidrange&",",",2,")<1 then contrl_message "您没有访问权限，请尝试其他版块。", "", "", "", ""
If adminctrllimit = 1 Then Call contrl_message ("您只有查看权限，没有管理权限，请升级权限后继续操作。", "", "", "", "")

%>
<title>生成静态页<%=application(siternd & "55trsitename")%></title>
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
		<td class="td_title" colspan="4">■ 整站静态化管理 &nbsp;&nbsp;| <a href="site_set.asp" target="_self"  style="color:#00F;">切换为动态版</a> | <a href="site_set.asp" target="_self" style="color:#00F;">切换为静态版</a> | <a href="site_set.asp" target="_self" style="color:#00F;">切换为伪静态版</a></td>
	</tr>
	<%if application(siternd & "55trurltype")<>2 then%>
	<tr colspan="4">
		<td> 当前网站状态不是静态版，无法生成静态，请到 "站点设置-- 动态、静态、伪静态"中切换为静态，<a href="site_set.asp" target="_self" style="color:#ff0000;">点此直达切换为静态版</a></td>
	</tr>
	<%else%>
	<tr class="font1">
		<td width="" align="center">操作</td>
		<td width="" align="center">选项</td>
		<td width="60" align="center">进度</td>
		<td width="60" align="center">&nbsp;</td>
	</tr>
	<tr class="">
		<form name="form1" id="form1" action="htmlsave.asp" target="ifr1" method="post" onSubmit="" >
			<td width="" align="center"><input type="hidden" value="default" name="action" id="action" />
				<input type="submit" value=" 生成首页 " name="submit1" id="submit1" class="button2"/></td>
			<td width="" align="center">&nbsp;</td>
			<td width="60" align="center"><iframe id="ifr1" name="ifr1" style="width:500px; height:30px; border:1px solid #0CF; padding:0; background:#E9FFFF;" scrolling="no" frameborder="0"   marginwidth="0" marginheight="0" ></iframe></td>
			<td width="60" align="center">&nbsp;</td>
		</form>
	</tr>
	<form name="form2" id="form2" action="htmlsave.asp" target="ifr2" method="post" onSubmit="" >
		<tr class="">
			<td width="" rowspan="4" align="center"><input type="hidden" value="list" name="action" id="action" />
				<input type="submit" value=" 生成列表页 " name="submit1" id="submit1" class="button2"/></td>
			<td width="" align="center"><span class="font5">可按shfit或ctrl键进行多选<br>
				全部留空为生成全部</span><br />
				<select name="listgroup" id="listgroup" size="10" multiple="multiple">
					
					<!--#include file="../crinc/column_select.asp"-->
					
				</select></td>
			<td width="60" rowspan="4" align="center"><iframe id="ifr2" name="ifr2" style="width:500px; height:30px; border:1px solid #0CF; padding:0; background:#E9FFFF;" scrolling="no" frameborder="0"   marginwidth="0" marginheight="0" ></iframe></td>
			<td width="60" rowspan="4" align="center">&nbsp;</td>
		</tr>
		<tr class="">
			<td align="center">ID范围，从
				<input name="id1" type="text" class="size12 input2" id="id1" maxlength="10" >
				至
				<input type="text" id="id2" name="id2" class="size12 input2" maxlength="10" ></td>
		</tr>
		<tr class="">
			<td align="center"><input type="text" id="days" name="days" class="size12 input2" maxlength="5" />
				天内更新的</td>
	</form>
		</tr>
	
	<tr class="">
		<td align="center">&nbsp;</td>
	</tr>
	<form name="form2" id="form2" action="htmlsave.asp" target="ifr3" method="post" onSubmit="" >
		<tr class="">
			<td width="" rowspan="4" align="center"><input type="hidden" value="show" name="action" id="action" />
				<input type="submit" value=" 生成内容/说明页 " name="submit1" id="submit1" class="button2"/></td>
			<td width="" align="center"><span class="font5">可按shfit或ctrl键进行多选<br>
				不选即为生成全部</span><br />
				<select name="listgroup" id="listgroup" size="10" multiple="multiple">
					
					<!--#include file="../crinc/column_select.asp"-->
					
				</select></td>
			<td width="60" rowspan="4" align="center"><iframe id="ifr3" name="ifr3" style="width:500px; height:30px; border:1px solid #0CF; padding:0; background:#E9FFFF;" scrolling="no" frameborder="0"   marginwidth="0" marginheight="0" ></iframe></td>
			<td width="60" rowspan="4" align="center">&nbsp;</td>
		</tr>
		<tr class="">
			<td align="center">ID范围，从
				<input type="text" id="id1" name="id1" class="size12 input2" maxlength="10" />
				至
				<input type="text" id="id2" name="id2" class="size12 input2" maxlength="10" /></td>
		</tr>
		<tr class="">
			<td align="center"><input type="text" id="days" name="days" class="size12 input2" maxlength="5" />
				天内更新的</td>
		</tr>
	</form>
	<tr class="">
		<td align="left"><br></td>
	</tr>
	<tr class="">
		<td colspan="4" align="left"><p>小提示：<br>
				1、本生成静态功能为服务器端后台运行，是指在浏览器中选择好要生成的页面，点击生成静态按钮，服务器会把所有指定的静态页面生成完毕，即使关闭浏览器也不影响静态生成。<br>
				2、这种方式有自身的优势，同时也需要管理员在页面较多的时候确认好所生成的页面，以避免指令发出想终止却要等待执行完成的情况。<br>
				3、文章每次发布之后都会自动生成内容页的静态，所以只需要再手动生成首页与文章所属的列表页静态即可，最好等每次的文章全部发表完毕后再生成首页和列表页。<br>
				4、当你的文章是采集的时候，就需要首页、内容页、列表页都生成的。<br>
				5、ID范围填写小技巧：只生成1个ID时可只在左边或右边填写该ID,想生成某个ID后面的所有静态时可在左边填写该ID，在右边填写一个很大的数字即可，就是说右边的ID不需要太准确，只要大于现有最大ID即可。 <br>
				6、首次开启静态版功能时需要整站生成一次静态。</td>
	</tr>
	<%end if%>
</table>
<br>
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
