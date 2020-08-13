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
<!--#include file="../core/class/tr_comment.asp"-->
<!--#include file="../core/class/tr_page.asp"-->
<%
Set ctr_comment = New tr_comment
Set ctr_page = New tr_page
isdel = request.querystring("isdel")
If isdel = 1 Then
  delstr = " and isdel=1 "
Else
  delstr = " and isdel=0 "
End If
iscom = request.querystring("iscom")
If IsNumeric(iscom) Then
  If iscom<>1 Then iscom = 0
Else
  iscom = 0
End If
arid = request.querystring("arid")
If IsNumeric(arid) And arid<>"" Then
  aridstr = " and ownarticle="&arid&" "
  aridurl = "&arid="&arid&""
Else
  aridstr = ""
  aridurl = ""
End If
If iscom = 1 Then
  iscomstr = " and types=1 "
  iscomurl = "&iscom=1"
  If InStr(","&adminlimitidrange&",", ",2,")<1 Then contrl_message "您没有访问权限，请尝试其他版块。", "", "", "", ""
Else
  iscomstr = " and types=2 "
  iscomurl = "&iscom=0"
  If InStr(","&adminlimitidrange&",", ",3,")<1 Then contrl_message "您没有访问权限，请尝试其他版块。", "", "", "", ""
End If
propertys = request.querystring("property")
  Select Case propertys
    Case "ispass"
      andpropertys = " and ispass=1 "&delstr&iscomstr&aridstr
    Case "isclose"
      andpropertys = " and ispass=0 "&delstr&iscomstr&aridstr
    Case "isreply"
      andpropertys = " and len(answer)>0 "&delstr&iscomstr&aridstr
    Case "isnotreply"
      andpropertys = " and (len(answer)=0 or answer is null) "&delstr&iscomstr&aridstr
    Case Else
      andpropertys = delstr&iscomstr&aridstr
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
    Case "answer"
      Call ctr_comment.admanswer()
      Call admin_log(adminname, adminid, adminname&" 回复留言/评论成功。，ID:"&id&"。")
      Call contrl_message_cbox ("回复留言/评论成功。", "", "id"&id&"", "", "")
    Case "del"
      Call changevalue("tr_comment", "isdel=1", "id in("&id&")")
      Call admin_log(adminname, adminid, adminname&" 删除留言/评论至回收站成功。，ID:"&id&"。")
      Call contrl_message ("删除留言/评论至回收站成功。",  "location.href", "", "", "parent")
    Case "dodel"
      Call delline("tr_comment", "id in("&id&")")
      Call admin_log(adminname, adminid, adminname&" 彻底删除留言/评论成功。，ID:"&id&"。")
      Call contrl_message ("彻底删除留言/评论成功。",  "location.href", "", "", "parent")
    Case "ispass"
      Call changevalue("tr_comment", "ispass=1", "id in("&id&")")
      Call admin_log(adminname, adminid, adminname&" 通过留言/评论成功。，ID:"&id&"。")
      Call contrl_message ("通过留言/评论成功。", "", "", "", "")
    Case "unisdel"
      Call changevalue("tr_comment", "isdel=0", "id in("&id&")")
      Call admin_log(adminname, adminid, adminname&" 从回收站恢复留言/评论成功。，ID:"&id&"。")
      Call contrl_message ("从回收站恢复留言/评论成功。",  "location.href", "", "", "parent")
    Case "unispass"
      Call changevalue("tr_comment", "ispass=0", "id in("&id&")")
      Call admin_log(adminname, adminid, adminname&"不通过留言/评论成功。，ID:"&id&"。")
      Call contrl_message ("不通过留言/评论成功。", "", "", "", "")
  End Select
%>
<title>评论/留言列表<%=application(siternd & "55trsitename")%></title>
<link rel="stylesheet" type="text/css" href="./skin/default/style.css">
<base target="_self">
<script type="text/javascript" src="./js/js.js"></script>
<script type="text/javascript" src="./js/hover.js"></script>
<script type="text/javascript" src="../js/jtimer.js"></script>
<link rel="shortcut icon" type="image/x-icon" href="../favicon.ico">
<link rel="bookmark" type="image/x-icon" href="../favicon.ico">
</head>
<body style="background:#fff; " class="w1">
<table width="100%" align="center" border="0" cellpadding="3" cellspacing="1" id="mytable2" class="size3" >
  <tr class="tr1" >
    <td class="td_title" colspan="7"><!--#include file="incgoback.asp"-->■
      <%if isdel=1 then%>
      回收站
      <%end if%>
      <%if iscom=1 then%>评论<%else%>留言<%end if%>管理
      <%if isdel<>1 then%>
      | <a href="comment_list.asp?55tr.com=1<%=iscomurl&aridurl%>">全部</a> <a href="comment_list.asp?property=ispass<%=iscomurl&aridurl%>">已通过</a> <a href="comment_list.asp?property=isclose<%=iscomurl&aridurl%>">未通过</a> <a href="comment_list.asp?property=isreply<%=iscomurl&aridurl%>">已回复</a> <a href="comment_list.asp?property=isnotreply<%=iscomurl&aridurl%>">未回复</a>
      <%end if%></td>
  </tr>
  <tr>
    <td colspan="7"><form target="_self" name="form1" id="form1" method="get" >
        <select name="fields" id="fields" class="size10">
          <option value="content" <%if fieldstr="content" then%>selected<%end if%> >内容
          <option value="username" <%if fieldstr="username" then%>selected<%end if%> >昵称
          <option value="mail" <%if fieldstr="mail" then%>selected<%end if%> >邮箱
          <option value="tel" <%if fieldstr="homepage" then%>selected<%end if%> >主页
        </select>
        <input type="text" name="words" id="words" value="<%=words%>" class="input2 size2">
        <input type="hidden" value="search" name="searchaction" id="searchaction">
        <input type="hidden" value="<%=iscom%>" name="iscom" id="iscom">
        <input type="submit" value=" 搜索 " name="submit1" id="submit1" class="button4">
      </form></td>
  </tr>
  <tr class="font1">
    <td width="" align="center" onClick="checkAll()" onMouseOver="style.backgroundColor='#FFBD66'" onMouseOut="style.backgroundColor=''"><label>
        <input type="checkbox" value="1" name="selectall" id="selectall">
      </label>
      全选</td>
    <td width="" align="center">序号</td>
    <td width="390" align="center">内容</td>
    <%if iscom=1 then%><td width="" align="center">文章ID</td><%end if%>
    <td width="50" align="center">已审核</td>
    <td width="250" align="center">回复内容</td>
    <td width="" align="center">操作</td>
  </tr>
  <form name="form2" id="form2" target="ifr1" method="post" onSubmit="return checkchose(this)">
    <%
page_size = 20
pagei = 0
'n=0
n = (pageno -1) * page_size
sql = "select * from tr_comment where 1=1 "&Str&andpropertys&" and id is not null order by id desc "
Call ctr_page.create_rs(sql, rs, conn, page_size, pageno)
If Not rs.EOF And Not rs.bof Then
  Do While Not rs.EOF And pagei<page_size
    pagei = pagei + 1
    n = n + 1
    id = Rs("id")
    title = Rs("title")
    content = Rs("content")
    ownarticle = Rs("ownarticle")
    owncolumn = Rs("owncolumn")
    owncolumnqueue = Rs("owncolumnqueue")
    addtime = Rs("addtime")
    types = Rs("types")
    username = Rs("username")
    mail = Rs("mail")
    homepage = Rs("homepage")
    qq = Rs("qq")
    tel = Rs("tel")
    ip = Rs("ip")
    ispass = Rs("ispass")
    isdel = Rs("isdel")
    answer = Rs("answer")
    astime = Rs("astime")
    asuser = Rs("asuser")
    istop = Rs("istop")
    iscommand = Rs("iscommand")
    Files = Rs("files")
    picfiles = Rs("picfiles")
    headpic = Rs("headpic")
%>
    <tr class="" id="tr<%=Rs("Id")%>">
      <td width="" align="center" onClick="checkOne('id<%=id%>')"onmouseover="style.backgroundColor='#FFBD66'" onMouseOut="style.backgroundColor=''"><input type="checkbox" value="<%=id%>" name="id" id="id<%=id%>" onClick="checkOne('id<%=id%>')">
        ID:<%=id%></td>
      <td width="" align="center"><%=n%></td>
      <td width="" align="left">昵称：<%=username%> | 邮箱：<%=mail%><br>
        主页：<%=homepage%> | IP：<%=ip%><br>
        时间：<%=FormatDate(addtime, 11) %><br>
        内容：<%=content%></td>
      <%if iscom=1 then%><td width="" align="center"><%=ownarticle%></td><%end if%>
      <td width="" align="center"><%if ispass=1 then%>
        通过
        <%else%>
        <font color="#FF0000">未通过</font>
        <%end if%></td>
      <td width="" align="left"><%if asuser<>"" then%>
        回复人：<%=asuser%><br>
        <%end if%>
        <%if astime<>"" then%>
        时 间 ：<%=FormatDate(astime, 11)%><br>
        <%end if%>
        <textarea name="answer<%=id%>" class="input3 size15" ><%if answer<>"" then%><%=CHTMLEncode_fan(answer)%><%end if%>
</textarea></td>
      <td width="" align="center"><input type="submit" class="button4" value="回复" onMouseDown="checkcheckbox('id<%=id%>','id')"></td>
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
      <td align="center"><input type="hidden" value="answer" name="action" id="action">
        <input type="submit" value="删除" name="submit1" id="submit1" class="button6 " onClick="changev('<%if isdel=1 then%>dodel<%else%>del<%end if%>')"></td>
      <td colspan="5" align="left"><%if isdel=1 then%>
        <input type="submit" value="恢复" name="submit1" id="submit1" class="button5" onClick="changev('unisdel')">
        <%end if%>
        <input type="submit" value="通过" name="submit1" id="submit1" class="button5" onClick="changev('ispass')">
        <input type="submit" value="不通过" name="submit1" id="submit1" class="button6 " onClick="changev('unispass')">
      </td>
    </tr>
  </form>
  <tr>
    <%if ctr_page.page_count>1 then%>
    <td colspan="7" align="left"><%= ctr_page.create_page("",ctr_page.page_count,pageno,"trpage")%></td>
    <% end if%>
  </tr>
  <%
Else
%>
  <tr>
    <td colspan="7" align="center">暂无相关项目
      <%if propertys<>"isdel" then%>
      ，请先添加。
      <%end if%></td>
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
