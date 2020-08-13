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
  Case "edit"
    Call createshare()
    Call admin_log(adminname, adminid, adminname&" 编辑网站分享功能成功。")
    Call contrl_message ("编辑网站分享功能成功。", "", "", "", "")
End Select
  sql="select * from tr_share where 1=1 "
  rs.open sql,conn,0,1
  if not rs.eof then
  stshare=rs("siteshare")
  shshare=rs("showshare")
  end if
  rs.close

%>
<title>网站广告管理<%=application(siternd & "55trsitename")%></title>
<link rel="stylesheet" type="text/css" href="./skin/default/style.css">
<base target="_self">
<script type="text/javascript" src="./js/js.js"></script>
<script type="text/javascript" src="./js/hover.js"></script>
<link rel="shortcut icon" type="image/x-icon" href="../favicon.ico">
<link rel="bookmark" type="image/x-icon" href="../favicon.ico">
</head><body style="background:#fff; " class="w1">
<table width="100%" align="center" border="0" cellpadding="3" cellspacing="1" id="mytable" class="size3" >
  <tr class="tr1" >
    <td class="td_title" colspan="7"><!--#include file="incgoback.asp"-->■ 网站分享功能</td>
  </tr>
  <tr class="font1">
    <td width="" align="center">操作</td>
    <td width="" align="center">整站显示分享按钮</td>
    <td width="" align="center">内容页文章底部分享按钮</td>
  </tr>
  <tr class="" valign="middle" >
    <form name="form1" id="form1" target="ifr1" method="post" >
      <td width="" align="center"><input type="hidden" value="edit" name="action" id="action">
        <input type="submit" value="编辑" name="submit1" id="submit1" class="button2"></td>
      <td width="" align="center"><textarea name="siteshare" class="input3 size14" id="siteshare"><%=stshare%></textarea></td>
      <td width="" align="center"><textarea name="showshare" class="input3 size14" id="showshare"><%=shshare%></textarea></td>
    </form>
  </tr>
  <tr>
    <td colspan="7" align="left"><p>此功能为管理现在流行的分享按钮功能，可添加两处位置的按钮，一处是全站所有前台页面都显示的按钮，适合于放置页面左侧或右侧悬浮的分享按钮，另一处是内容页文章底部的分享按钮，适合于放置一排按钮式的分享按钮。<br>
        推荐使用的分享按钮代码获取网站：<a href="http://share.baidu.com/" target="_blank">百度分享</a> | <a href="http://www.jiathis.com/" target="_blank">JiaThis分享</a> | <a href="http://www.bshare.cn/" target="_blank">BShare分享</a> <br>
        小提示：<br>
        1、不建议两处分享位置都使用同一个网站的分享按钮，因为在某些情况下会出现冲突的问题，如果出现不能显示的情况，请更换其他网站的分享代码 <br>
        2、首次使用此功能需要整站生成静态一次之后即可生效，以后即使改动分享代码也不再需要生成静态（仅在静态版情况下，其他动态版，伪静态版不需要生成）</p></td>
  </tr>
</table>
<br>
<script type="text/javascript" >
showtable("mytable")
</script> 
<%
  Public Sub createshare()
  siteshare=request.form("siteshare")
  showshare=request.form("showshare")
  sql="select * from tr_share where 1=1 "
  rs.open sql,conn,0,2
  if rs.eof then
  rs.addnew()
  rs("id")=1
  end if
  rs("addtime")=now()
  rs("siteshare")=siteshare
  rs("showshare")=showshare
  rs.update
  rs.close
      jssiteshare = tojs(siteshare) & vbCrLf
      jsshowshare = tojs(showshare) & vbCrLf
    Call createfile(jsshowshare, "../crinc/showshare.asp", 2, "utf-8")
    Call createfile(jssiteshare, "../crinc/siteshare.asp", 2, "utf-8")
  End Sub
%>
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