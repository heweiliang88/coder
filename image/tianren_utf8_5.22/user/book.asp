<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../inc/charsetasp.asp"-->
<!DOCTYPE html>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<!--#include file="../core/Conn.asp"-->
<!--#include file="../core/fun/fun.asp"-->
<!--#include file="../core/fun/core.asp"-->
<!--#include file="../core/class/tr_page.asp"-->
<!--#include file="../core/class/tr_user.asp"-->
<!--#include file="../core/class/tr_comment.asp"-->
<!--#include file="../skin/default/kp171014.asp"-->
<%'include此行代码供安装插件用勿删改%>
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



Const funpath = "../"
Const dbpath = "../"
Set ctr_page = New tr_page
Set ctr_user = New tr_user
Set ctr_comment = New tr_comment
%>
<%pageno = tr_killstr(request.querystring("pageno"), 1, 8, 1, "", "", "", "", 1)
%>
<%'head此行代码供安装插件用勿删改%>
<!--#include file="../inc/headtaginc.asp"-->
<title>反馈留言_<%=application(siternd & "55trsitetitle")%></title>
<link rel="shortcut icon" type="image/x-icon" href="<%=funpath%>favicon.ico" >
<link rel="Bookmark" type="image/x-icon" href="<%=funpath%>favicon.ico">
<!--#include file="../inc/phoneewmjs.asp"-->
<%'headend此行代码供安装插件用勿删改%>
</head>
<body>
<!--#include file="../inc/head.asp"-->
<div class="container trblock" id="trblock178888">
<div class=" col-lg-9 <%=colbs9%> trrow1199 trmarb1199 trlist trovh" id="trleft178888">
<div class="trlisttitle1 "><span class="trcrumbs">当前位置： <a href="<%=funpath%>">网站首页</a> &gt;
      会员个人中心 &gt; 反馈留言 </span><span class="trmcrumbs">当前位置： <a href="<%=funpath%>">网站首页</a> &gt; 反馈留言 </span></div>
<div class="publicnr trpad1 trmar2">
<%'body1此行代码供安装插件用勿删改%>
<%if application(siternd & "55tropenguestbook")=1 then%>
<div class="alert alert-warning alert-dismissible trnotice1" role="alert">
<button type="button" class="close" data-dismiss="alert" aria-label="Close"><span aria-hidden="true">&times;</span></button>
<%=application(siternd & "55trguestnotice")%> </div>
<form id="form1" name="form1" method="post" action="usersave.asp" target="ifr1" onSubmit="return checkguest(this)">
<%'body5此行代码供安装插件用勿删改%>
<table class="trsendtb" >
<tr>
<td align="center" colspan="11"><textarea id="content" name="content" class="trgcontent trinput6"></textarea></td>
</tr>
<%'body6此行代码供安装插件用勿删改%>
<tr>
<td align="left" nowrap> 昵称<span class="trfont2"> * </span></td>
<td align="left" ><input type="text" value="" name="username" id="username" class=" trinput3" /></td>
<td align="left" nowrap>email</td>
<td align="left" ><input type="text" value="" name="mail" id="mail" class=" trinput3" /></td>
<td align="left" nowrap> 主页</td>
<td align="left" ><input type="text" value="" name="homepage" id="homepage" class=" trinput3" /></td>
<td align="left" nowrap > 验证码</td>
<td align="left" ><input name="vcode" type="text" class=" trinput7" id="vcode" value="" maxlength="4" /></td>
<td align="left" ><img id="codeimg" src="../core/Code.asp?rd=" style="cursor:hand; " title="点击更换" onClick="javascript:this.src='../core/Code.asp?rd='+randomString(10)+''" height="24" border="0" width="70"></td>
<td align="right" ><input type="submit" value="立即提交" name="submit1" id="submit1" class="trbt3 trmar1" />
<input type="hidden" value="guest" name="action" /></td>
</tr>
<%'body7此行代码供安装插件用勿删改%>
</table>
<%'body8此行代码供安装插件用勿删改%>
</form>
<table align="center" border="0" cellpadding="3" cellspacing="1" id="mytable" class="trtable2" >
<%
sql = "select id,username,content,addtime,homepage,asuser,answer,astime from tr_comment where types=2 and ispass=1 and isdel=0 order by addtime desc , id desc  "
page_size = 20
pagei = 0
n = (pageno -1) * page_size
Call ctr_page.create_rs(sql, rs, conn, page_size, pageno)
'	rs.open sql,conn,0,1
If Not rs.EOF Then
  Do While Not rs.EOF
    ' and pagei<page_size
    If pagei>= page_size Then Exit Do
    pagei = pagei + 1
    n = n + 1
%>
<tr class="">
<td width=""><div class="trcontents trfl trbookbt">
<p class="trp1"> <span class="uname">
<%if trim(rs("homepage"))<>"" then%>
<%if isweburl(rs("homepage")) then%>
<a href="#"><%=rs("username")%></a>
<%else%>
<%=rs("username")%>
<%end if%>
<%else%>
<%=rs("username")%>
<%end if%>
</span> 说： <%=rs("content")%> </p>
<p class="trp2"><%=FormatDate(rs("addtime"), 1) %> </p>
<%if len(rs("answer"))>1 then%>
<p class="trp3">管理员 回复：<%=rs("answer")%></p>
<p class="trp2"><%=FormatDate(rs("astime"), 1) %> </p>
<%end if%>
<%'body9此行代码供安装插件用勿删改%>
</div></td>
</tr>
<%
rs.movenext
Loop
Else
%>
<tr>
<td>暂无留言，请先发表。</td>
</tr>
<%
End If
rs.Close
%>
</table>
<%if ctr_page.page_count>1 then%>
<%= ctr_page.create_page("",ctr_page.page_count,pageno,"trpage")%>
<% end if%>
<%else%>
<div class="trunotice trfontu">留言功能目前已关闭。 </div>
<%end if%>
<%'body2此行代码供安装插件用勿删改%>
</div>
</div>
<div class="col-lg-3 <%=colbs3%> trrow1199 trlistright trovh" id="trlistright178888">
<%'body3此行代码供安装插件用勿删改%>
<!--#include file="../inc/rightinc.asp"-->
<%'body4此行代码供安装插件用勿删改%>
</div>
</div>
<!--#include file="../inc/bottomads.asp"-->
<div class="trpublicline clearfix "> </div>
<!--#include file="../inc/foot.asp"--> 
<!--#include file="../inc/ifr.asp"-->
</body>
</html>
