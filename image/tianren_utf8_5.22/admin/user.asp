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
<!--#include file="../core/class/tr_user.asp"-->
<!--#include file="../core/md5.asp"-->
<!--#include file="../crinc/ulevstr.asp"-->
 <%
Set ctr_user = New tr_user
If InStr(","&adminlimitidrange&",", ",4,")<1 Then contrl_message "您没有访问权限，请尝试其他版块。", "", "", "", ""
id = Trim(request.querystring("id"))
If IsNumeric(id) Then
  Call ctr_user.getfieldv(id)
  actionval = "edit"
Else
  actionval = "add"
  id = 0
End If
action = Trim(request.Form("action"))
If action<>"" And adminctrllimit = 1 Then Call contrl_message ("您只有查看权限，没有管理权限，请升级权限后继续操作。", "", "", "", "")
Select Case action
  Case "add"
    Call ctr_user.adminadduser()
    Call admin_log(adminname, adminid, adminname&" 添加会员成功。")
    Call contrl_message ("添加会员成功。", "", "submit1", "form1", "")
  Case "edit"
    Call ctr_user.adminedit()
    Call admin_log(adminname, adminid, adminname&" 修改会员成功。")
    Call contrl_message ("修改会员成功。", "", "submit1", "", "")
  Case "del"
    Call ctr_user.del()
    Call admin_log(adminname, adminid, adminname&" 删除会员成功。")
    Call contrl_message ("删除会员成功。", "", "submit1", "", "")
End Select
%>
 <title><%if actionval="add" then%>添加会员 <%elseif actionval="edit" then%>编辑会员ID：<%=id%><%end if%><%=application(siternd & "55trsitename")%></title>
 <link rel="stylesheet" type="text/css" href="./skin/default/style.css">
 <base target="_self">
 <script type="text/javascript" src="./js/js.js"></script>
 <script type="text/javascript" src="./js/hover.js"></script>
 <script type="text/javascript" src="../js/JTimer.js"></script>
<link rel="shortcut icon" type="image/x-icon" href="../favicon.ico">
<link rel="bookmark" type="image/x-icon" href="../favicon.ico">
 </head>
 <body style="background:#fff; " class="w1">
<form name="form1" id="form1" action="" method="post" onSubmit="return <%if actionval="add" then%>checkadminadduser<%else%>checkadminedituser<%end if%>(this); " target="ifr1">
<table width="100%" align="center" border="0" cellpadding="3" cellspacing="1" id="mytable" class="size3" >
   <tr class="tr1" >
    <td class="td_title" colspan="13"><!--#include file="incgoback.asp"-->■ <%if actionval="add" then%>添加会员 <%elseif actionval="edit" then%>编辑会员<span style="color:#F00; margin-left:20px;">ID：<%=id%></span><%end if%></td>
  </tr>
  <tr>
    <td align="right" width="120">登录名：</td>
    <td colspan="3"><input name="username" type="text" id="username" size="20" maxlength="50" class="input2 size2" value="<%=ctr_user.username%>"  <%if actionval="edit" then%>
 readonly="readonly"<%end if%>/><span class="font5"> *</span></td>
    <td></td>
  </tr>
    <tr>
    <td align="right">密码：</td>
    <td colspan="3"><input name="passwd1" type="text" id="passwd1" size="20" maxlength="50" class="input2 size2" /><span class="font5"> *</span></td>
    <td></td>
  </tr>
  <tr>
    <td align="right">重复密码：</td>
    <td colspan="3"><input name="passwd2" type="text" id="passwd2" size="20" maxlength="50" class="input2 size2" /><span class="font5"> *</span></td>
    <td></td>
  </tr>
  <tr>
    <td align="right">邮箱：</td>
    <td colspan="3"><input name="mail" type="text" id="mail" size="20" maxlength="50" class="input2 size2"  value="<%=ctr_user.mail%>"/><span class="font5"> * </span></td>
    <td></td>
  </tr>
  <%if actionval="edit" then%>
  <tr>
    <td align="right" width="120">会员等级：</td>
    <td><%
spuserjfstr = Split(userjfstr, ",")
splevnamestr = Split(levnamestr, ",")
points = clng(ctr_user.points)
For i = 0 To UBound(spuserjfstr)
  If points>= clng(spuserjfstr(i)) And points<clng(spuserjfstr(i + 1)) Then
    response.Write splevnamestr(i)
    Exit For
  ElseIf points>= clng(spuserjfstr(UBound(spuserjfstr))) Then
    response.Write splevnamestr(UBound(spuserjfstr))
    Exit For
  End If
Next
%></td>
    <td align="right">&nbsp;</td>
    <td>&nbsp;</td>
    <td></td>
  </tr>
  <tr>
    <td align="right" width="120">注册时间：</td>
    <td><%=ctr_user.regtime%></td>
    <td align="right">登陆总次数：</td>
    <td><%=ctr_user.logcount%></td>
    <td></td>
  </tr>
  <tr>
    <td align="right" width="120">最后登录时间：</td>
    <td><%=ctr_user.lastlogintime%></td>
    <td align="right">注册ip：</td>
    <td><%=ctr_user.regip%></td>
    <td></td>
  </tr>
  <tr>
    <td width="120" align="right">解冻：</td>
    <td><label for="isfree">
      <input name="isfree" type="checkbox" id="isfree" value="1" <%if ctr_user.isfree=1 then%>checked="checked" <%end if%> />
      解冻（不勾选即为冻结）</label></td>
    <td align="right">发表文章总数：</td>
    <td><%safemd5 = trmd5(ctr_user.username)%><%=getfieldvalue("tr_article","count(id)"," and contrimd5='"&safemd5&"' and isdel=0 ",2,"")%></td>
    <td></td>
  </tr>
  <%end if%>
  <tr>
    <td align="right">总积分：</td>
    <td><input name="points" type="text" id="points" size="20" maxlength="50" class="input2 size2" value="<%=ctr_user.points%>" /></td>
    <td align="right">开启发布文章：</td>
    <td><label for="openarticle">
      <input name="openarticle" type="checkbox" id="openarticle" value="1" <%if ctr_user.openarticle=1 then%>checked="checked" <%end if%>/>
      开启（不勾选即为关闭）</label></td>
    <td></td>
  </tr>
  <tr>
    <td width="" align="right">已连续签到：</td>
    <td><%
nowtime = FormatDate(Now(), 1)
If DateDiff("d", ctr_user.lastsigntime, nowtime)>1 Or IsNull(ctr_user.lastsigntime) Then
%>
      0
      <%else%>
      <%=ctr_user.signs%>
      <%end if%>
      天</td>
    <td align="right">发布文章所得积分：</td>
    <td><input name="gavearticlepoints" type="text" id="gavearticlepoints" size="20" maxlength="50" class="input2 size2" value="<%=ctr_user.gavearticlepoints%>" /></td>
    <td></td>
  </tr>
  <tr>
    <td align="right" width="">&nbsp;</td>
    <td>&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td>&nbsp;</td>
    <td></td>
  </tr>
  <tr>
    <td align="right">以下为选填内容&nbsp;&nbsp;</td>
    <td colspan="3"></td>
    <td></td>
  </tr>
  <tr>
    <td align="right">QQ号码：</td>
    <td colspan="3"><input name="qq" type="text" id="qq" size="20" maxlength="50" class="input2 size2" value="<%=ctr_user.qq%>" /></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td align="right">昵称：</td>
    <td colspan="3"><input name="nickname" type="text" id="nickname" size="20" maxlength="50" class="input2 size2" value="<%=ctr_user.nickname%>" /></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td align="right">主页：</td>
    <td colspan="3"><input name="homepage" type="text" id="homepage" size="20" maxlength="50" class="input2 size2" value="<%=ctr_user.homepage%>" /></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td align="right">电话号码：</td>
    <td colspan="3"><input name="tel" type="text" id="tel" size="20" maxlength="50" class="input2 size2" value="<%=ctr_user.tel%>" /></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td align="right">年龄：</td>
    <td colspan="3"><input name="age" type="text" id="age" size="20" maxlength="50" class="input2 size2" value="<%=ctr_user.age%>" /></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td align="right">性别：</td>
    <td colspan="3">&nbsp;&nbsp;<label for="sex1"><input name="sex" type="radio" id="sex1" value="1"<%if ctr_user.sex=1 then%>checked<%end if%>/> 男</label>&nbsp; <label for="sex2"><input name="sex" type="radio" id="sex2" value="2"<%if ctr_user.sex=2 then%>checked<%end if%>/> 女</label></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td align="right">地址：</td>
    <td colspan="3"><input name="address" type="text" id="address" size="20" maxlength="50" class="input2 size4" value="<%=ctr_user.address%>"/></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td align="right">生日：</td>
    <td colspan="3"><input name="birthday" type="text" id="birthday" size="20" maxlength="50" class="input2 size2" value="<%=ctr_user.birthday%>"onclick="JTC.setday({startDay: '1990-06-06'})" readonly /></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td align="right">签名：</td>
    <td colspan="3"><textarea name="signature" cols="20" class="input3 size9" id="signature"><%=ctr_user.signature%></textarea></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td align="right">简介：</td>
    <td colspan="3"><textarea name="intro" cols="20" class="input3 size9" id="intro"><%=ctr_user.intro%></textarea></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td align="right">&nbsp;</td>
    <td colspan="3">&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td align="right">&nbsp;</td>
    <td colspan="3"><input name="submit1" type="submit" id="submit1" class="button1" value="立即提交" /><input name="action" type="hidden" id="action" value="<%=actionval%>"/><input name="id" type="hidden" id="id" value="<%=id%>"/></td>
    <td>&nbsp;</td>
  </tr>
</table>
</form>
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