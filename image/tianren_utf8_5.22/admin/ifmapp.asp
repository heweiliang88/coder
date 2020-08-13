<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../inc/charsetasp.asp"-->
<!DOCTYPE html>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%Server.ScriptTimeOut = 36000
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
<%
'If InStr(","&adminlimitidrange&",", ",2,")<1 Then contrl_message "您没有访问权限，请尝试其他版块。", "", "", "", ""
sql="select systemtype,sysver from tr_system where 1=1 "
rs.open sql,conn,0,1
if not rs.eof then
t=rs("systemtype")
v=rs("sysver")
end if
rs.close

't=tr_killstr(request.querystring("t"), 1, 16, 1, "", "", "", "", 1)
lx=tr_killstr(request.querystring("lx"), 1, 16, 1, "", "", "", "", 1)


Public Function GetholeUrl()
  GetholeUrl =  Request.ServerVariables("SERVER_NAME")
  If Request.ServerVariables("SERVER_PORT") <> 80 Then GetholeUrl = GetholeUrl &":" & Request.ServerVariables("SERVER_PORT")
  GetholeUrl = lcase(GetholeUrl & Request.ServerVariables("URL"))
'  If Trim(Request.QueryString) <>"" Then GetholeUrl = GetholeUrl &"?" & Trim(Request.QueryString)
GetholeUrl=replace(GetholeUrl,"http://","")
End Function 

%>
<title><%=application(siternd & "55trsitename")%></title>
<link rel="stylesheet" type="text/css" href="./skin/default/style.css">
<link rel="stylesheet" href="../kindeditor_4.1.10/themes/default/default.css" />
<base target="_self">
<script type="text/javascript" src="./js/js.js"></script>
<link rel="shortcut icon" type="image/x-icon" href="../favicon.ico">
<link rel="bookmark" type="image/x-icon" href="../favicon.ico">
<script>
function changeFrameHeight(ifrid){
    var ifm= document.getElementById(ifrid); 
	if (document.documentElement.clientHeight<880){
    ifm.height=document.documentElement.clientHeight;
	}
	else{
		ifm.height=880
	}
}
</script>
</head><body style="background:#fff; " class="w1">
<table width="100%" align="center" border="0" cellpadding="3" cellspacing="1" id="mytable" class="size3" >
  <tr class="tr1" >
    <td class="td_title" colspan="8"><!--#include file="incgoback.asp"-->■
      获取应用</td>
  </tr>
  <tr class="">
    <td align="right"><iframe id="iframepage" src="https://www.55tr.com/ifmapplist.asp?u=<%=GetholeUrl%>&lx=<%=lx%>&t=<%=t%>&v=<%=v%>&fr=<%=frbm%>&cn1=<%=cn1%>&cn2=<%=cn2%>&cn3=<%=cn3%>&cn4=<%=cn4%>&cn5=<%=cn5%>" style="width:100%; border:none; background:#fff;" onload="changeFrameHeight('iframepage')"  ></iframe></td>
  </tr>
</table>
<!--#include file="../inc/ifr.asp"--> 
<script>
window.onresize=function(){  
     changeFrameHeight('iframepage');  
} </script>
</body>
</html>

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
