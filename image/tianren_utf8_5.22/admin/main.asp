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
<%If InStr(","&adminlimitidrange&",", ",")<1 Then contrl_message "您没有访问权限，请尝试其他版块。", "", "", "", ""%>
<title>后台首页 -<%=application(siternd & "55trsitetitle")%>%></title>
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
    <td class="td_title" colspan="4"><!--#include file="incgoback.asp"-->■ 网站概况</td>
  </tr>
  <tr>
    <td width="" colspan="1" align="right">&nbsp;文章数：</td>
    <td colspan="1" width="">&nbsp;<%= getfieldvalue("tr_article","count(id)"," and ispass=1 and isdel=0 ",2,"")%> 篇</td>
    <td  colspan="1" align="right">最后更新时间：</td>
    <td colspan="1">&nbsp;<%= FormatDate(getfieldvalue("tr_article","top 1 addtime"," and ispass=1 and isdel=0 ",2,"order by addtime desc , id desc "),1)%></td>
  </tr>
  <tr>
    <td  colspan="1" align="right">&nbsp;栏目数：</td>
    <td colspan="1">&nbsp;<%= getfieldvalue("tr_column","count(id)"," and isopen=1 ",2,"")%> 个</td>
    <td  colspan="1" align="right">&nbsp;会员数：</td>
    <td colspan="1">&nbsp;<%= getfieldvalue("tr_user","count(id)"," and isdel=0 ",2,"")%>位 | 今日新增<%= getfieldvalue("tr_user","count(id)"," and isdel=0 and datediff('d','"&now()&"',regtime)=0 ",2,"")%></td>
  </tr>
  <tr>
    <td  colspan="1" align="right">&nbsp;留言数：</td>
    <td colspan="1">&nbsp;<%= getfieldvalue("tr_comment","count(id)"," and isdel=0 and types=2 ",2,"")%>条 | 今日新增<%= getfieldvalue("tr_comment","count(id)"," and isdel=0 and datediff('d','"&now()&"',addtime)=0 and types=2 ",2,"")%></td>
    <td  colspan="1" align="right">评论数：</td>
    <td colspan="1">&nbsp;<%= getfieldvalue("tr_comment","count(id)"," and isdel=0 and types=1 ",2,"")%>条 | 今日新增<%= getfieldvalue("tr_comment","count(id)"," and isdel=0 and datediff('d','"&now()&"',addtime)=0 and types=1 ",2,"")%></td>
  </tr>
  <tr>
    <td  colspan="1" align="right">当前程序版本：</td>
    <%sysver=trim(getfieldvalue("tr_system", "sysver", "", 1, ""))%>
    <td colspan="1">&nbsp;<%=trversion%><%=sysver%></td>
    <td  colspan="1" align="right">数据库：</td>
    <td colspan="1">&nbsp;<%=ucase(data_type)%></td>
  </tr>
  <tr>
    <td  colspan="1" align="right"><font color="#FF0000">[此处为程序升级功能]</font>：</td>
    <td colspan="1">
      <%
checkver = request.QueryString("checkver")
If Not IsNumeric(checkver) Then checkver = 0
If checkver = 1 Then

t=trim(getfieldvalue("tr_system", "systemtype", "",2, ""))
  allverstr = verHttpPage("http://www.55tr.com/outupdate.asp?t="&t&"&u="&GetholeUrl&"&updatekey="&application(siternd & "55trupdatekey")&"&v="&sysver&"&fr="&frbm&"")
  response.write allverstr
Else
%>
<a href="main.asp?checkver=1" target="_self"><span style="color:#fff; background:#09F"> 点此检测最新版.. </span> </a>
<%	end if%></td>
    <td  colspan="1" align="right">授权状态：</td>
    <td colspan="1">
      <%
checkver = request.QueryString("checkver")
If Not IsNumeric(checkver) Then checkver = 0
If checkver = 2 Then
t=trim(getfieldvalue("tr_system", "systemtype", "",2, ""))
  allverstr = verHttpPage("http://www.55tr.com/outtrright.asp?t="&t&"&u="&GetholeUrl&"&v="&sysver&"&fr="&frbm&"")
  response.write allverstr
Else
%>
<a href="main.asp?checkver=2" target="_self"><span style="color:#fff; background:#F90"> 点此检查授权状态 </span> </a>
<%end if%></td>
  </tr>
  <tr height="18" align="center">
    <td colspan="4"  align="left"><p>版权声明：<br>
天人文章管理系统，以下简称&ldquo;本程序、程序&rdquo;可免费用于个人非商业用途，其他用途需付费购买版权，请按照如下约定执行：<br>
**本程序可免费用于个人非商业用途，除个人非商业外的用途均需要付费购买版权，并在官网登记域名作为正规版权记录，并可获得电子版版权证书(不是绑定域名，而是在官网数据库中登记域名旨在法律维权时避开已购买版权的域名)，侵权使用必将承担法律风险并追究高额赔偿<br>
**购买版权将获得稳定健全的升级服务、免费安装插件、模板、配套工具免费使用、官方优惠优先获得等特权。并可用于企业、政府、校园、单位、团体、个人、机构、部门、机关、公司及其他商业用途<br>
**购买版权后，除源代码中版权声明，后台左下角的&ldquo;源程序：55tr.com&rdquo;字样（这两种版权字样都不会在前台显示，只作为程序版权归属）以外所有代码都可以修改及删除，包括前台页面底部的官网链接及版权等代码都可修改。<br>
**购买版权后可用本程序为别人或自己提供建设网站，修改源码，二次开发的服务，但禁止发布或传播本程序及本程序的任何形态！禁止出售本程序及本程序的任何形态。说白了就是不能将程序出售或传播，否则侵权。（例如：小李与客户约定使用本程序制作某中学的官网，并收取某中学的网站制作服务费用1万元，并购买版权，此行为不侵权。小赵与客户约定使用本程序制作某小学的官网，并收取某小学的网站制作服务费用1万元，未购买版权，此行为侵权(因为未购买版权)，请终止！小孙使用本程序建立了个人漫画网站供用户浏览并获取广告收益，未购买版权，此行为不侵权。小张使用本程序制作为企业网站进行出售，并购买版权，此行为侵权（因为传播出售了程序），请终止！小王又将本程序制作为小说网站源码在网络免费发布传播提供下载，并购买版权，此行为侵权（因为传播出售了程序），请终止！）<br>
**不购买版权的个人非商业用途依然享受升级、安装插件、模板的权利，可修改源代码，但不得修改源代码中&ldquo;官网链接&rdquo;、&ldquo;版权声明&rdquo;、&ldquo;应用中心&rdquo;相关的代码，前台页面底部的&ldquo;Copyright 2016 天人文章管理系统 版权所有，授权xxx使用 Powered by 55TR.COM&rdquo;字样不得删除,否则请购买版权<br>
**不购买版权的除个人非商业外的用途请及时购买版权或终止使用本程序，否则侵权<br>
**计算机软件著作权登记证书登记号：2016SR204636，侵权必究，作者QQ81962480 欢迎洽谈基于此程序开发、修改业务。<br>
**版权声明会存在修改的偶然性，请以官网公布的对应程序版本版权声明为准！具体的版权声明、特权、及功能介绍，请浏览：http://www.55tr.com/gwappshow.asp?id=11<br>
**天人文章管理系统3.48之前的程序版本版权声明以此为准，之前的版权声明作废<br>
      <br>
    </td>
  </tr>
</table>
<script type="text/javascript" >
showtable("mytable")
</script>
<%

Function verHttpPage(HttpUrl)
  Dim http
  Set http = server.CreateObject("MSX"&"ML2.Serv"&"erXM"&"LHTTP")
  Http.Open "GET", HttpUrl, False
  Http.send()
  If Http.readystate<>4 Then Exit Function
  verHttpPage = htmlBytesToBstr(Http._
  ResponseBody, "ut"&"f-8")
  Set http = Nothing
  If Err.Number<>0 Then Err.Clear
End Function


Function htmlBytesToBstr(body, Cset)
  Dim Objstream
  Set Objstream = Server.CreateObject("ad"&"odb.St"&"ream")
  objstream.Type = 1
  objstream.Mode = 3
  objstream.Open
  objstream.Write body
  objstream.Position = 0
  objstream.Type = 2
  objstream.Charset = Cset
  If tmpcopy = "" Then
    htmlBytesToBstr = objstream.ReadText
  Else
    htmlBytesToBstr = tmpcopy
  End If
  objstream.Close
  Set objstream = Nothing
End Function

Public Function GetholeUrl()
  GetholeUrl =  Request.ServerVariables("SERVER_NAME")
  If Request.ServerVariables("SERVER_PORT") <> 80 Then GetholeUrl = GetholeUrl &":" & Request.ServerVariables("SERVER_PORT")
  GetholeUrl = lcase(GetholeUrl & Request.ServerVariables("URL"))
GetholeUrl=replace(GetholeUrl,"http://","")
End Function 


Set ctr_admin = Nothing
%>
</body>
</html>