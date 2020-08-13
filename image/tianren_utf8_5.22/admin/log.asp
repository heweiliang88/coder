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
<!--#include file="../core/class/tr_page.asp"-->
 <%
Set ctr_page = New tr_page
If InStr(","&adminlimitidrange&",", ",1,")<1 Then contrl_message "您没有访问权限，请尝试其他版块。", "", "", "", ""
pageno = CInt(request.querystring("pageno"))
If pageno<1 Then pageno = 1
action = Trim(request.Form("action"))
If action<>"" And adminctrllimit = 1 Then Call contrl_message ("您只有查看权限，没有管理权限，请升级权限后继续操作。", "", "", "", "")
Select Case action
  Case "del"
    id = Trim(Request.Form("id"))
    Call del()
    Call admin_log(adminname, adminid, adminname&" 删除后台操作日志成功，ID:"&id&"。")
    Call contrl_message ("删除后台操作日志成功。", "location.href", "submit1", "", "parent")
End Select
Private Sub del()
  id = Trim(Request.Form("id"))
  sql = "delete from tr_log where id in ("&id&")"
  conn.Execute(sql)
End Sub
%>
 <title>后台操作日志<%=application(siternd & "55trsitename")%></title>
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
    <td class="td_title" colspan="7"><!--#include file="incgoback.asp"-->■ 后台操作日志管理</td>
  </tr>
   <tr class="font1">
    <td width="60" align="center" onClick="checkAll()" onMouseOver="style.backgroundColor='#FFBD66'" onMouseOut="style.backgroundColor=''"><label>
        <input type="checkbox" value="1" name="selectall" id="selectall" />
      </label>
       全选</td>
    <td width="" align="center">序号</td>
    <td width="450" align="center">内容</td>
    <td width="" align="center">操作人</td>
    <td align="center">操作人ID</td>
    <td align="center">IP地址</td>
    <td width="90" align="center">添加时间</td>
  </tr>
   <form name="form2" id="form2" target="ifr1" method="post" onSubmit="return checkchose(this)">
    <%
page_size = 30
pagei = 0
n = 0
sql = "select * from tr_log where 1=1 and id is not null order by id desc "
Call ctr_page.create_rs(sql, rs, conn, page_size, pageno)
If Not rs.EOF And Not rs.bof Then
  Do While Not rs.EOF And pagei<page_size
    pagei = pagei + 1
    n = n + 1
    id = Rs("id")
    addtime = Rs("addtime")
    content = Rs("content")
    ip = Rs("ip")
    username = Rs("username")
    userid = Rs("userid")
    isdel = Rs("isdel")
%>
    <tr class="" id="tr<%=Rs("Id")%>">
       <td width="" align="center" onClick="checkOne('id<%=id%>')"onmouseover="style.backgroundColor='#FFBD66'" onMouseOut="style.backgroundColor=''"><input type="checkbox" value="<%=id%>" name="id" id="id<%=id%>" onClick="checkOne('id<%=id%>')" /></td>
       <td width="" align="center"><%=n%></td>
       <td width="" align="left"><%=content%></td>
       <td width="" align="center"><%=username%></td>
       <td width="" align="center"><%=userid%></td>
       <td width="" align="left"><%=ip%></td>
       <td width="" align="center"><%=FormatDate(addtime, 11) %></td>
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
        <input type="submit" value=" 删除 " name="submit1" id="submit1" class="button3 "/></td>
       <td colspan="5" align="left"><%if ctr_page.page_count>1 then response.write ctr_page.create_page("log.asp?pageno=",ctr_page.page_count,pageno,"trpage") end if%></td>
     </tr>
  </form>
   <%
Else
%>
   <tr>
    <td colspan="7" align="center">暂无相关项目，请先添加。 </td>
  </tr>
   <%end if%>
 </table>
<script type="text/javascript" >
showtable("mytable2")
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
