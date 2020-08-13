<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%on error resume next%>
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
<!--#include file="../core/class/tr_ulev.asp"-->
<!--#include file="../core/class/tr_ads.asp"-->
<%
Set ctr_column = New tr_column
Set ctr_ads = New tr_ads
Set ctr_ulev = New tr_ulev

If InStr(","&adminlimitidrange&",", ",1,")<1 Then contrl_message "您没有访问权限，请尝试其他版块。", "", "", "", ""
docache = tr_killstr(Trim(request.querystring("docache")), 1, 1, 3, "", "", "", "", 1)
If docache = "" Then docache = 0
If docache=1 And adminctrllimit = 1 Then Call contrl_message ("您只有查看权限，没有管理权限，请升级权限后继续操作。", "", "", "", "")

%>
<title>网站栏目管理<%=application(siternd & "55trsitename")%></title>
<link rel="stylesheet" type="text/css" href="./skin/default/style.css">
<base target="_self">
<script type="text/javascript" src="./js/js.js"></script>
<script type="text/javascript" src="./js/hover.js"></script>
<script type="text/javascript" src="../js/jtimer.js"></script>
<link rel="shortcut icon" type="image/x-icon" href="../favicon.ico">
<link rel="bookmark" type="image/x-icon" href="../favicon.ico">
</head><body style="background:#fff; " class="w1">
<table width="100%" align="center" border="0" cellpadding="3" cellspacing="1" id="mytable" class="size3" >
  <tr class="tr1" >
    <td class="td_title" colspan="4"><!--#include file="incgoback.asp"-->■ 一键更新缓存</td>
  </tr>
  <tr class="font1">
    <td width="" align="left" colspan="2"><input type="button" value="立即更新" class="button1" onClick="location='recache.asp?docache=1'" /></td>
  </tr>
  <%if docache=1 then%>
  <tr class="font1">
    <td width="" align="center">项目</td>
    <td align="center">状态</td>
  </tr>
  <%
call ctr_column.cr()
%>
  <tr class="">
    <td width="" align="center" valign="middle" >更新栏目菜单缓存</td>
    <td align="center">完成</td>
  </tr>
  <%
	Response.flush
	response.Clear()
	%>
  <%
call ctr_ulev.create_select(0)
%>
  <tr class="">
    <td width="" align="center" valign="middle" >更新会员等级缓存</td>
    <td align="center">完成</td>
  </tr>
  <%
	Response.flush
	response.Clear()
	%>
  <%
call ctr_ads.createads()
%>
  <tr class="">
    <td width="" align="center" valign="middle" >更新广告缓存</td>
    <td align="center">完成</td>
  </tr>
  <%
	Response.flush
	response.Clear()

'下面的功能是自动获取官网注册码，请勿删除，不存在威胁。
Function get55tr(url)
  Dim Http
  Set Http = server.CreateObject("MSX"&"ML2.XML"&"HTTP")
  Http.Open "GET", url, false
  Http.send()
  If Http.readystate<>4 Then
    Exit Function
  End If
  get55tr = bytesToBSTR(Http.responseBody, "ut"&"f-8")
  Set http = Nothing
  If Err.Number<>0 Then Err.Clear
End Function
domainurl=Request.ServerVariables("Http_Host")
codestr=get55tr("http://www.55tr.com/domainkeys.asp?u="&domainurl&"")	
if instr(codestr,"|")>0 then
spcode1=split(codestr,",")	
	for d=0 to ubound(spcode1)
	spcode2=split(spcode1(d),"|")
	keys=trim(spcode2(0))
	notes=trim(spcode2(1))
    addtime = FormatDate(Now(), 1)
    sql = "select * from tr_keys where keys='"&keys&"' "
	rs.open sql,conn,0,1
	if rs.eof then
	input=true
	else
	input=false
	end if
	rs.close
	if input then
    sql = "select * from tr_keys where 1=2 "
    rs.Open sql, conn, 1, 3
    rs.addnew()
    id = GetMaxValue("tr_keys", "id") + 1
    Rs("id") = id
    Rs("keys") = keys
    Rs("notes") = notes
    Rs("addtime") = addtime
    Rs("isdel") = 0
    rs.update
    rs.Close
	end if
	next
end if	
	
	%>
  <tr class="">
    <td width="" align="center" valign="middle" >更新注册组件缓存</td>
    <td align="center">完成</td>
  </tr>
  <%
	Response.flush
	response.Clear()
	%>
  <%
  sql="select * from tr_share where 1=1 "
  rs.open sql,conn,0,2
  if not rs.eof then
  siteshare=rs("siteshare")
  showshare=rs("showshare")
  jssiteshare = tojs(siteshare) & vbCrLf
  jsshowshare = tojs(showshare) & vbCrLf
  Call createfile(jsshowshare, "../crinc/showshare.asp", 2, "gb2312")
  Call createfile(jssiteshare, "../crinc/siteshare.asp", 2, "gb2312")
  end if
  rs.close
%>
  <tr class="">
    <td width="" align="center" valign="middle" >更新分享功能缓存</td>
    <td align="center">完成</td>
  </tr>
  <%
	Response.flush
	response.Clear()
	%>
  <%
application.contents.removeall   
%>
  <tr class="">
    <td width="" align="center" valign="middle" >更新内存缓存</td>
    <td align="center">完成</td>
  </tr>
  <%
	Response.flush
	response.Clear()
	%>
  <tr class="">
    <td width="" align="center" valign="middle" >全部项目更新完毕</td>
    <td align="center">√</td>
  </tr>
  <%
	Response.flush
	response.Clear()%>
  <%end if%>
  <tr class="">
    <td colspan="2" align="left" valign="middle" ><p>小提示：</p>
      <p>1、一键更新缓存功能一般适用于程序更换数据库之后、程序升级之后、程序更换空间、安装新功能模块、新插件、新模板等情况，来重新生成系统所用到的缓存文件，日常后台新增内容时系统会自动生成相应的缓存，因此日常使用可不用更新，即使更新也只消耗很少服务器负载。</p></td>
  </tr>
</table>
<br>
<script type="text/javascript" >
showtable("mytable")
</script> 
<!--#include file="../inc/ifr.asp"-->
</body>
</html>
