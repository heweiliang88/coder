<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../inc/charsetasp.asp"-->
<!--#include file="../core/Conn.asp"-->
<!--#include file="../core/fun/fun.asp"-->
<!--#include file="../core/fun/core.asp"-->
<!--#include file="../core/class/tr_user.asp"-->
<!--#include file="../core/class/tr_comment.asp"-->
<!--#include file="../core/md5.asp"-->
<!DOCTYPE html>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>操作接收页</title>
</head>
<body>
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



Call checkoutside()
Const funpath = "../"
Const dbpath = "../"
Set ctr_user = New tr_user
Set ctr_comment = New tr_comment
act = request.Form("action")
Select Case act
  Case "reg"
    Call ctr_user.userreg()
  Case "login"
    Call ctr_user.truserlogin()
  Case "uedit"
    Call ctr_user.useredit()
  Case "addnew"
    If application(siternd & "55tropenarticle") = 1 Then Call ctr_user.ueditnews("add")
  Case "editnew"
    Call ctr_user.ueditnews("edit")
  Case "delnew"
    Call ctr_user.udelnew()
  Case "guest"
    If application(siternd & "55tropenguestbook") = 1 Then Call ctr_comment.addcommentguest(2)
  Case "comment"
    If application(siternd & "55tropencomment") = 1 Then Call ctr_comment.addcommentguest(1)
End Select
Set ctr_user = Nothing
Set ctr_comment = Nothing
%>
</body>
</html>
