<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../inc/charsetasp.asp"-->
<!DOCTYPE html>
<html>
 <head>
 <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%
Const dbpath = "../"
Const funpath = "../"
%>
<!--#include file="admin_fun.asp"-->
<!--#include file="../core/class/tr_article.asp"-->
<!--#include file="../core/class/tr_page.asp"-->
 <%
Set ctr_article = New tr_article
Set ctr_page = New tr_page
If InStr(","&adminlimitidrange&",", ",2,")<1 Then contrl_message "您没有访问权限，请尝试其他版块。", "", "", "", ""
isdel = request.querystring("isdel")
If isdel = 1 Then
  delstr = " and isdel=1 "
Else
  delstr = " and isdel=0 "
End If
propertys = request.querystring("property")
  Select Case propertys
    Case "isfocus"
      andpropertys = " and isfocus=1 "&delstr
    Case "isroll"
      andpropertys = " and isroll=1 "&delstr
    Case "isindextop"
      andpropertys = " and isindextop=1 "&delstr
    Case "islisttop"
      andpropertys = " and islisttop=1 "&delstr
    Case "iscommend"
      andpropertys = " and iscommend=1 "&delstr
    Case "isrollimg"
      andpropertys = " and isrollimg=1 "&delstr
    Case "picfiles"
      andpropertys = " and picfiles<>'' "&delstr
    Case "ispass"
      andpropertys = " and ispass=1 "&delstr
    Case "isclose"
      andpropertys = " and ispass=0 "&delstr
    Case "iscontribute"
      andpropertys = " and iscontribute=1 "&delstr
    Case "isabout"
      andpropertys = " and isabout=1 "&delstr
      'case "isdel"
      'andpropertys=" and isdel=1 "
    Case Else
      andpropertys = delstr
  End Select
  pageno = CInt(request.querystring("pageno"))
  If Not IsNumeric(pageno) Then pageno = 1
  If pageno<1 Then pageno = 1
  searchaction = request.querystring("searchaction")
  If searchaction = "search" Then
    columnid = request.querystring("column")
    words = request.querystring("words")
    types = request.querystring("types")
    If columnid<>"" Then str1 = " and queueid like '%,"&columnid&",%' "
    If Not IsNumeric (columnid) Then columnid = 0
    If types<>"" And words<>"" Then str2 = " and "&types&" like '%"&words&"%' "
    Str = str1&str2
  End If
  action = Trim(request.Form("action"))
  id = Trim(Request.Form("id"))
  If action<>"" And adminctrllimit = 1 Then Call contrl_message ("您只有查看权限，没有管理权限，请升级权限后继续操作。", "", "", "", "")
  Select Case action
    Case "del"
      Call changevalue("tr_article", "isdel=1", "id in("&id&")")
      Call admin_log(adminname, adminid, adminname&" 删除文章至回收站成功，ID:"&id&"。")
      Call contrl_message ("删除文章至回收站成功。", "location.href", "", "", "parent")
    Case "dodel"
      Call delline("tr_article", "id in("&id&")")
      Call admin_log(adminname, adminid, adminname&" 彻底删除文章成功，ID:"&id&"。")
      Call contrl_message ("彻底删除文章成功。", "location.href", "", "", "parent")
    Case "isfocus"
      Call changevalue("tr_article", "isfocus=1", "id in("&id&")")
      Call admin_log(adminname, adminid, adminname&" 设置焦点图成功，ID:"&id&"。")
      Call contrl_message ("设置焦点图成功。", "", "", "", "")
    Case "isroll"
      Call changevalue("tr_article", "isroll=1", "id in("&id&")")
      Call admin_log(adminname, adminid, adminname&" 设置滚动文字成功，ID:"&id&"。")
      Call contrl_message ("设置滚动文字成功。", "", "", "", "")
    Case "isindextop"
      Call changevalue("tr_article", "isindextop=1", "id in("&id&")")
      Call admin_log(adminname, adminid, adminname&" 设置首页置顶成功，ID:"&id&"。")
      Call contrl_message ("设置首页置顶成功。", "", "", "", "")
    Case "islisttop"
      Call changevalue("tr_article", "islisttop=1", "id in("&id&")")
      Call admin_log(adminname, adminid, adminname&" 设置列表页置顶成功，ID:"&id&"。")
      Call contrl_message ("设置列表页置顶成功。", "", "", "", "")
    Case "iscommend"
      Call changevalue("tr_article", "iscommend=1", "id in("&id&")")
      Call admin_log(adminname, adminid, adminname&" 设置推荐成功，ID:"&id&"。")
      Call contrl_message ("设置推荐成功。", "", "", "", "")
    Case "isrollimg"
      Call changevalue("tr_article", "isrollimg=1", "id in("&id&")")
      Call admin_log(adminname, adminid, adminname&" 设置图片滚动成功，ID:"&id&"。")
      Call contrl_message ("设置图片滚动成功。", "", "", "", "")
    Case "ispass"
      ' id = trim(Request.Form("id"))
      Call changevalue("tr_article", "ispass=1", "id in("&id&")")
      sql = "select * from tr_article where id in("&id&") and isgavepoints=0 and iscontribute=1 "
      rs.Open sql, conn, 1, 3
      If Not rs.EOF Then
        Do While Not rs.EOF
          rs("isgavepoints") = 1
          rs("ispass") = 1
          Call changevalue("tr_user", "gavearticlepoints=iif(isNull(gavearticlepoints),0,gavearticlepoints)+"&application(siternd & "55trpublisharticlepoints")&",points=iif(isNull(points),0,points)+"&application(siternd & "55trpublisharticlepoints")&"", "safemd5 ='"&rs("contrimd5")&"'")
          rs.update
          rs.movenext
        Loop
      End If
      rs.Close
      Call admin_log(adminname, adminid, adminname&" 开通文章成功，ID:"&id&"。")
      Call contrl_message ("开通文章成功。", "", "", "", "")
    Case "unisdel"
      Call changevalue("tr_article", "isdel=0", "id in("&id&")")
      Call admin_log(adminname, adminid, adminname&" 从回收站恢复文章成功，ID:"&id&"。")
      Call contrl_message ("从回收站恢复文章成功。", "location.href", "", "", "parent")
    Case "unfocus"
      Call changevalue("tr_article", "isfocus=0", "id in("&id&")")
      Call admin_log(adminname, adminid, adminname&" 取消焦点图成功，ID:"&id&"。")
      Call contrl_message ("取消焦点图成功。", "", "", "", "")
    Case "unroll"
      Call changevalue("tr_article", "isroll=0", "id in("&id&")")
      Call admin_log(adminname, adminid, adminname&" 取消滚动文字成功，ID:"&id&"。")
      Call contrl_message ("取消滚动文字成功。", "", "", "", "")
    Case "unindextop"
      Call changevalue("tr_article", "isindextop=0", "id in("&id&")")
      Call admin_log(adminname, adminid, adminname&" 取消首页置顶成功，ID:"&id&"。")
      Call contrl_message ("取消首页置顶成功。", "", "", "", "")
    Case "unlisttop"
      Call changevalue("tr_article", "islisttop=0", "id in("&id&")")
      Call admin_log(adminname, adminid, adminname&" 取消列表页置顶成功，ID:"&id&"。")
      Call contrl_message ("取消列表页置顶成功。", "", "", "", "")
    Case "uncommend"
      Call changevalue("tr_article", "iscommend=0", "id in("&id&")")
      Call admin_log(adminname, adminid, adminname&" 取消推荐成功，ID:"&id&"。")
      Call contrl_message ("取消推荐成功。", "", "", "", "")
    Case "unrollimg"
      Call changevalue("tr_article", "isrollimg=0", "id in("&id&")")
      Call admin_log(adminname, adminid, adminname&" 取消滚动图片成功，ID:"&id&"。")
      Call contrl_message ("取消滚动图片成功。", "", "", "", "")
    Case "unpass"
      Call changevalue("tr_article", "ispass=0", "id in("&id&")")
      Call admin_log(adminname, adminid, adminname&" 关闭文章成功，ID:"&id&"。")
      Call contrl_message ("关闭文章成功。", "", "", "", "")

  End Select
%>
 <title>文章列表<%=application(siternd & "55trsitename")%></title>
 <link rel="stylesheet" type="text/css" href="./skin/default/style.css">
 <base target="_self">
 <script type="text/javascript" src="./js/js.js"></script>
 <script type="text/javascript" src="./js/hover.js"></script>
 <script type="text/javascript" src="../js/jtimer.js"></script>
<link rel="shortcut icon" type="image/x-icon" href="../favicon.ico">
<link rel="bookmark" type="image/x-icon" href="../favicon.ico">
 </head>
<body style="background:#ffffff; " class="w1">
<table width="100%" align="center" border="0" cellpadding="3" cellspacing="1" id="mytable2" class="size3" >
  <tr class="tr1" >
		<td class="td_title" colspan="10"><!--#include file="incgoback.asp"-->■ <%if isdel=1 then%>回收站<%end if%>文章列表 <%if isdel<>1 then%>| <a href="article_list.asp">全部</a> <a href="article_list.asp?property=isfocus">焦点图</a> <a href="article_list.asp?property=isroll">文字滚动</a> <a href="article_list.asp?property=isindextop">首页置顶</a> <a href="article_list.asp?property=islisttop">列表置顶</a> <a href="article_list.asp?property=iscommend">推荐</a> <a href="article_list.asp?property=isrollimg">图片滚动</a> <a href="article_list.asp?property=picfiles">图片</a> <a href="article_list.asp?property=ispass">已通过</a> <a href="article_list.asp?property=isclose">未通过</a> <a href="article_list.asp?property=iscontribute">投稿 </a> <a href="article_list.asp?property=isabout">关于 </a><%end if%></td>
  </tr>
	<tr>
		<td colspan="10" valign="middle"><form target="_self" name="form1" id="form1" method="get" >
				<select name="types" id="types" class="size10">
					<option value="title" <%if types="title" then%>selected<%end if%> >标题
					<option value="content" <%if types="content" then%>selected<%end if%> >内容
				</select>
				<select name="column" id="column" class="size10">
					<option value="" >选择栏目
					<!--#include file="../crinc/column_select.asp"-->
				</select>
				<input type="text" name="words" id="words" value="<%=words%>" class="input2 size2">
				<input type="hidden" value="search" name="searchaction" id="searchaction">
				<input type="submit" value=" 搜索 " name="submit1" id="submit1" class="button4">
			</form></td>
  </tr>
	<tr class="font1">
		<td width="" align="center" onClick="checkAll()" onMouseOver="style.backgroundColor='#FFBD66'" onMouseOut="style.backgroundColor=''"><label>
				<input type="checkbox" value="1" name="selectall" id="selectall">
			</label>
			全选</td>
		<td width="" align="center">序号</td>
		<td width="200" align="center">标题</td>
		<td width="" align="center">状态</td>
		<td align="center">所属栏目</td>
		<td width="" align="center">属性</td>
		<td width="" align="center">点击数</td>
		<td width="" align="center">添加时间</td>
	</tr>
	<form name="form2" id="form2" target="ifr1" method="post" onSubmit="return checkchose(this)">
		<%
page_size = 20
pagei = 0
n = (pageno -1) * page_size
sql = "select * from tr_article where 1=1 "&Str&andpropertys&" and id is not null order by id desc "
Call ctr_page.create_rs(sql, rs, conn, page_size, pageno)
If Not rs.EOF And Not rs.bof Then
  Do While Not rs.EOF And pagei<page_size
    pagei = pagei + 1
    n = n + 1
    id = Rs("id")
    channelid = Rs("channelid")
    title = Rs("title")
    titlecolor = Rs("titlecolor")
    content = Rs("content")
    keywords = Rs("keywords")
    keywords2 = Rs("keywords2")
    descriptions = Rs("descriptions")
    addtime = Rs("addtime")
    clicks = Rs("clicks")
    commentnum = Rs("commentnum")
    jumpurl = Rs("jumpurl")
    updown = Rs("updown")
    author = Rs("author")
    sources = Rs("sources")
    inputer = Rs("inputer")
    tags = Rs("tags")
    minuspoints = Rs("minuspoints")
    limittype = Rs("limittype")
    limitpoints = Rs("limitpoints")
    limitmore = Rs("limitmore")
    columnid = Rs("columnid")
    queueid = Rs("queueid")
    columnname = Rs("columnname")
    isfocus = Rs("isfocus")
    isroll = Rs("isroll")
    isindextop = Rs("isindextop")
    islisttop = Rs("islisttop")
    iscommend = Rs("iscommend")
    isrollimg = Rs("isrollimg")
    picfiles = Rs("picfiles")
    ispass = Rs("ispass")
    isdel = Rs("isdel")
    isabout = Rs("isabout")
    iscontribute = Rs("iscontribute")
    contributor = Rs("contributor")
%>
		<tr class="" id="tr<%=Rs("Id")%>">
			<td width="" align="center" onClick="checkOne('id<%=id%>')"onmouseover="style.backgroundColor='#FFBD66'" onMouseOut="style.backgroundColor=''"><input type="checkbox" value="<%=id%>" name="id" id="id<%=id%>" onClick="checkOne('id<%=id%>')">
				ID:<%=id%></td>
			<td width="" align="center"><%=n%></td>
			<td width="400" align="left" style="word-break:break-all;"><a href="article.asp?id=<%=id%>"><%=title%></a></td>
			<td width="" align="center"><%if ispass=1 then%>
				通过
				<%else%>
				<font color="#FF0000">未通过</font>
				<%end if%></td>
			<td width="" align="center"><%=getfieldvalue("tr_column", "colname", "and id="&columnid&"", 1, "")%></td>
			<td width="" align="center"><%if isfocus=1 then%>
				焦
				<%end if%>
  <% if isroll=1 then%>
				滚
				<%end if%>
  <% if isindextop=1 then%>
				首顶
				<%end if%>
  <% if islisttop=1 then%>
				列顶
				<%end if%>
  <% if iscommend=1 then%>
				荐
				<%end if%>
  <% if isrollimg=1 then%>
				滚图
				<%end if%>
  <% if len(picfiles)>3 then%>
				图
				<%end if%>
  <% if iscontribute=1 then%>
				稿
				<%end if%>
                  <% if isabout=1 then%>
				关于
				<%end if%></td>
			<td width="" align="center"><%=clicks%></td>
			<td width="" align="center"><%=FormatDate(addtime, 11) %></td>
	  </tr>
  <%
rs.movenext
Loop
%>
		<tr>
			<td width="" align="center" onClick="checkAll()" onMouseOver="style.backgroundColor='#FFBD66'" onMouseOut="style.backgroundColor=''"><label>
					<input type="checkbox" value="1" name="selectall2" id="selectall2">
				</label>
				全选</td>
			<td align="center"><input type="hidden" value="del" name="action" id="action">
				<input type="submit" value="删除" name="submit1" id="submit1" class="button6 " onClick="changev('<%if isdel=1 then%>dodel<%else%>del<%end if%>')"></td>
			<td colspan="6" align="left"><%if isdel=1 then%>
            <input type="submit" value="恢复" name="submit1" id="submit1" class="button5" onClick="changev('unisdel')">
            <%end if%>
            <input type="submit" value="焦" name="submit1" id="submit1" class="button5" onClick="changev('isfocus')">
				<input type="submit" value="滚" name="submit1" id="submit1" class="button5 " onClick="changev('isroll')">
				<input type="submit" value="首顶" name="submit1" id="submit1" class="button5 " onClick="changev('isindextop')">
				<input type="submit" value="列顶" name="submit1" id="submit1" class="button5 " onClick="changev('islisttop')">
				<input type="submit" value="荐" name="submit1" id="submit1" class="button5 " onClick="changev('iscommend')">
				<input type="submit" value="图滚" name="submit1" id="submit1" class="button5 " onClick="changev('isrollimg')">
				<input type="submit" value="开通" name="submit1" id="submit1" class="button5 " onClick="changev('ispass')">
				<input type="submit" value="否焦" name="submit1" id="submit1" class="button6 " onClick="changev('unfocus')">
				<input type="submit" value="否滚" name="submit1" id="submit1" class="button6 " onClick="changev('unroll')">
				<input type="submit" value="否首顶" name="submit1" id="submit1" class="button6 " onClick="changev('unindextop')">
				<input type="submit" value="否列顶" name="submit1" id="submit1" class="button6 " onClick="changev('unlisttop')">
				<input type="submit" value="否荐" name="submit1" id="submit1" class="button6 " onClick="changev('uncommend')">
				<input type="submit" value="否图滚" name="submit1" id="submit1" class="button6 " onClick="changev('unrollimg')">
    <input type="submit" value="关闭" name="submit1" id="submit1" class="button6 " onClick="changev('unpass')">
		  </td>
		</tr>
	</form>
	<tr>
		<%if ctr_page.page_count>1 then%><td colspan="10" align="left"><%= ctr_page.create_page("",ctr_page.page_count,pageno,"trpage")%> </td><% end if%>
	</tr>
	<%
Else
%>
	<tr>
		<td colspan="10" align="center">暂无相关项目<%if propertys<>"isdel" then%>，请先添加。 <%end if%></td>
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
