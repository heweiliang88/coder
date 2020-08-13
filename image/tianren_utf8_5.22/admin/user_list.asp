<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../inc/charsetasp.asp"-->
<!DOCTYPE html>
<html>
 <head>
 <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<%
Const dbpath = "../"
Const funpath = "../"
%>
<!--#include file="admin_fun.asp"-->
<!--#include file="../core/class/tr_user.asp"-->
<!--#include file="../core/class/tr_page.asp"-->
<!--#include file="../core/md5.asp"-->
 <%
Set ctr_user = New tr_user
Set ctr_page = New tr_page
If InStr(","&adminlimitidrange&",", ",4,")<1 Then contrl_message "您没有访问权限，请尝试其他版块。", "", "", "", ""
isdel = request.querystring("isdel")
If isdel = 1 Then
  delstr = " and isdel=1 "
Else
  delstr = " and isdel=0 "
End If
propertys = request.querystring("property")
  Select Case propertys
    Case "isfree"
      andpropertys = " and isfree=1 "&delstr
    Case "isfreeze"
      andpropertys = " and isfree=0 "&delstr
    Case "opencomment"
      andpropertys = " and opencomment=1 "&delstr
    Case "closecomment"
      andpropertys = " and opencomment=0 "&delstr
    Case "openarticle"
      andpropertys = " and openarticle=1 "&delstr
    Case "closearticle"
      andpropertys = " and openarticle=0 "&delstr
    Case "sex1"
      andpropertys = " and sex=1 "&delstr
    Case "sex2"
      andpropertys = " and sex=2 "&delstr
    Case Else
      andpropertys = delstr
  End Select
  pageno = CInt(request.querystring("pageno"))
  If Not IsNumeric(pageno) Then pageno = 1
  If pageno<1 Then pageno = 1
  searchaction = request.querystring("searchaction")
  If searchaction = "search" Then
    words = request.querystring("words")
    fieldstr = request.querystring("fields")
    If fieldstr<>"" Then Str = " and "&fieldstr&" like '%"&words&"%' "
  End If
  action = Trim(request.Form("action"))
  If action<>"" And adminctrllimit = 1 Then Call contrl_message ("您只有查看权限，没有管理权限，请升级权限后继续操作。", "", "", "", "")
  id = Trim(Request.Form("id"))
  Select Case action
    Case "del"
      Call changevalue("tr_user", "isdel=1", "id in("&id&")")
      Call admin_log(adminname, adminid, adminname&" 删除会员至回收站成功，ID:"&id&"。")
      Call contrl_message ("删除会员至回收站成功。",  "location.href", "", "", "parent")
    Case "dodel"
      Call delline("tr_user", "id in("&id&")")
      Call admin_log(adminname, adminid, adminname&" 彻底删除会员成功，ID:"&id&"。")
      Call contrl_message ("彻底删除会员成功。",  "location.href", "", "", "parent")
    Case "isfree"
      Call changevalue("tr_user", "isfree=1", "id in("&id&")")
      Call admin_log(adminname, adminid, adminname&" 解冻会员成功，ID:"&id&"。")
      Call contrl_message ("解冻会员成功。", "", "", "", "")
    Case "opencomment"
      Call changevalue("tr_user", "opencomment=1", "id in("&id&")")
      Call admin_log(adminname, adminid, adminname&" 开通留言成功，ID:"&id&"。")
      Call contrl_message ("开通留言成功。", "", "", "", "")
    Case "openarticle"
      Call changevalue("tr_user", "openarticle=1", "id in("&id&")")
      Call admin_log(adminname, adminid, adminname&" 开通发文章成功，ID:"&id&"。")
      Call contrl_message ("开通发文章成功。", "", "", "", "")
    Case "unisdel"
      Call changevalue("tr_user", "isdel=0", "id in("&id&")")
      Call admin_log(adminname, adminid, adminname&" 从回收站恢复会员成功，ID:"&id&"。")
      Call contrl_message ("从回收站恢复会员成功。", "location.href", "", "", "parent")
    Case "unisfree"
      Call changevalue("tr_user", "isfree=0", "id in("&id&")")
      Call admin_log(adminname, adminid, adminname&" 冻结会员成功，ID:"&id&"。")
      Call contrl_message ("冻结会员成功。", "", "", "", "")
    Case "unopencomment"
      Call changevalue("tr_user", "opencomment=0", "id in("&id&")")
      Call admin_log(adminname, adminid, adminname&" 关闭留言功能成功，ID:"&id&"。")
      Call contrl_message ("关闭留言功能成功。", "", "", "", "")
    Case "unopenarticle"
      Call changevalue("tr_user", "openarticle=0", "id in("&id&")")
      Call admin_log(adminname, adminid, adminname&" 关闭发文章功能成功，ID:"&id&"。")
      Call contrl_message ("关闭发文章功能成功。", "", "", "", "")
  End Select
%>
 <title>会员列表<%=application(siternd & "55trsitename")%></title>
 <link rel="stylesheet" type="text/css" href="./skin/default/style.css">
 <base target="_self">
 <script type="text/javascript" src="./js/js.js"></script>
 <script type="text/javascript" src="./js/hover.js"></script>
 <script type="text/javascript" src="../js/JTimer.js"></script>
 <link rel="shortcut icon" type="image/x-icon" href="../favicon.ico">
<link rel="bookmark" type="image/x-icon" href="../favicon.ico">
</head>
 <body style="background:#fff; " class="w1">
<table width="100%" align="center" border="0" cellpadding="3" cellspacing="1" id="mytable2" class="size3" >
	<tr class="tr1" >
		<td class="td_title" colspan="12"><!--#include file="incgoback.asp"-->■ <%if isdel=1 then%>回收站<%end if%>会员列表 <%if isdel<>1 then%>| <a href="user_list.asp">全部</a> <a href="user_list.asp?property=isfreeze">冻结</a> <a href="user_list.asp?property=isfree">解冻</a> <a href="user_list.asp?property=opencomment">开启留言评论</a> <a href="user_list.asp?property=closecomment">关闭留言评论</a> <a href="user_list.asp?property=openarticle">开启发文章</a> <a href="user_list.asp?property=closearticle">关闭发文章</a> <a href="user_list.asp?property=sex1">男</a> <a href="user_list.asp?property=sex2">女</a> <%end if%></td>
	</tr>
	<tr>
		<td colspan="12"><form target="_self" name="form1" id="form1" method="get" >
				<select name="fields" id="fields" class="size10">
					<option value="username" <%if fieldstr="username" then%>selected<%end if%> >登录名
					<option value="nickname" <%if fieldstr="nickname" then%>selected<%end if%> >昵称
					<option value="qq" <%if fieldstr="qq" then%>selected<%end if%> >QQ
					<option value="tel" <%if fieldstr="tel" then%>selected<%end if%> >电话

				</select>
				<input type="text" name="words" id="words" value="<%=words%>" class="input2 size2"/>
				<input type="hidden" value="search" name="searchaction" id="searchaction" />
				<input type="submit" value=" 搜索 " name="submit1" id="submit1" class="button5"/>
			</form></td>
	</tr>
	<tr class="font1">
		<td width="" align="center" onClick="checkAll()" onMouseOver="style.backgroundColor='#FFBD66'" onMouseOut="style.backgroundColor=''"><label>
				<input type="checkbox" value="1" name="selectall" id="selectall" />
			</label>
			全选</td>
		<td width="" align="center">序号</td>
		<td width="" align="center">会员登陆名</td>
		<td width="" align="center">昵称</td>
		<td align="center">最后登陆时间</td>
		<td width="" align="center">登陆次数</td>
		<td align="center">积分</td>
		<td width="" align="center">是否解冻</td>
		<td width="" align="center">文章数</td>
		<td width="" align="center">发布文章</td>
	</tr>
	<form name="form2" id="form2" target="ifr1" method="post" onSubmit="return checkchose(this)">
		<%
page_size = 20
pagei = 0
'n=0
n = (pageno -1) * page_size
sql = "select * from tr_user where 1=1 "&Str&andpropertys&" and id is not null order by id desc "
Call ctr_page.create_rs(sql, rs, conn, page_size, pageno)
If Not rs.EOF And Not rs.bof Then
  Do While Not rs.EOF And pagei<page_size
    pagei = pagei + 1
    n = n + 1
    id = Rs("id")
    username = Rs("username")
    passwd = Rs("passwd")
    nickname = Rs("nickname")
    qq = Rs("qq")
    homepage = Rs("homepage")
    tel = Rs("tel")
    regtime = Rs("regtime")
    lastlogintime = Rs("lastlogintime")
    regip = Rs("regip")
    lastloginip = Rs("lastloginip")
    age = Rs("age")
    mail = Rs("mail")
    sex = Rs("sex")
    address = Rs("address")
    zip = Rs("zip")
    birthday = Rs("birthday")
    points = Rs("points")
    signature = Rs("signature")
    intro = Rs("intro")
    photo = Rs("photo")
    headpic = Rs("headpic")
    isfree = Rs("isfree")
    gavecommentpoints = Rs("gavecommentpoints")
    gavearticlepoints = Rs("gavearticlepoints")
    ismailactivated = Rs("ismailactivated")
    isdel = Rs("isdel")
    articlenum = Rs("articlenum")
    commentnum = Rs("commentnum")
    mailcode = Rs("mailcode")
    opencomment = Rs("opencomment")
    openarticle = Rs("openarticle")
    openuserreg = Rs("openuserreg")
    openuserlogin = Rs("openuserlogin")
    edittime = Rs("edittime")
%>
		<tr class="" id="tr<%=Rs("Id")%>">
			<td width="" align="center" onClick="checkOne('id<%=id%>')"onmouseover="style.backgroundColor='#FFBD66'" onMouseOut="style.backgroundColor=''"><input type="checkbox" value="<%=id%>" name="id" id="id<%=id%>" onClick="checkOne('id<%=id%>')" />
				ID:<%=id%></td>
			<td width="" align="center"><%=n%></td>
			<td width="" align="center"><a href="user.asp?id=<%=id%>"><%=username%></a></td>
			<td width="" align="center"><%=nickname%></td>
			<td width="" align="center"><%=lastlogintime%></td>
			<td width="" align="center"><%=logcount%>
	    </td>
			<td width="" align="left"><%=points%></td>
			<td width="" align="center"><%if isfree=1 then%>
				解冻
				<%else%>
				<font color="#FF0000">冻结</font>
  <%end if%></td>
			<td width="" align="center"><%safemd5 = trmd5(username)%><%=getfieldvalue("tr_article","count(id)"," and contrimd5='"&safemd5&"' and isdel=0 ",2,"")%></td>
			<td width="" align="center"><%if openarticle=1 then%>
				开通
				<%else%>
				<font color="#FF0000">关闭</font>
    <%end if%></td>
		</tr>
		<%
rs.movenext
Loop
%>
		<tr>
			<td width="" align="center" onClick="checkAll()" onMouseOver="style.backgroundColor='#FFBD66'" onMouseOut="style.backgroundColor=''"><label>
					<input type="checkbox" value="1" name="selectall2" id="selectall2"/>
				</label>
				全选</td>
			<td align="center"><input type="hidden" value="del" name="action" id="action" />
				<input type="submit" value="删除" name="submit1" id="submit1" class="button6 " onClick="changev('<%if isdel=1 then%>dodel<%else%>del<%end if%>')"/></td>
			<td colspan="10" align="left"><%if isdel=1 then%>
            <input type="submit" value="恢复" name="submit1" id="submit1" class="button5" onClick="changev('unisdel')" />
            <%end if%><input type="submit" value="解冻" name="submit1" id="submit1" class="button5" onClick="changev('isfree')" />
				<input type="submit" value="开留言" name="submit1" id="submit1" class="button5 " onClick="changev('opencomment')" />
				<input type="submit" value="开发文章" name="submit1" id="submit1" class="button5 " onClick="changev('openarticle')" />
				<input type="submit" value="冻结" name="submit1" id="submit1" class="button6 " onClick="changev('unisfree')" />
				<input type="submit" value="关留言" name="submit1" id="submit1" class="button6 " onClick="changev('unopencomment')" />
	      <input type="submit" value="关发文章" name="submit1" id="submit1" class="button6 " onClick="changev('unopenarticle')" /><%'end if%>
				</td>
		</tr>
	</form>
	<tr>
		<%if ctr_page.page_count>1 then%><td colspan="13" align="left"><%= ctr_page.create_page("",ctr_page.page_count,pageno,"trpage")%> </td><% end if%>
	</tr>
	<%
Else
%>
	<tr>
		<td colspan="12" align="center">暂无相关项目<%if propertys<>"isdel" then%>，请先添加。 <%end if%></td>
  </tr>
	<%end if%>
</table>
<script type="text/javascript" >
showtable("mytable2")
function changev(value){
document.getElementById("action").value	=value
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