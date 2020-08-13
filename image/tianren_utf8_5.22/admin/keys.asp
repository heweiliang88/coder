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
<!--#include file="../core/class/tr_keys.asp"-->
<!--#include file="../core/class/tr_page.asp"-->
<%
Set ctr_keys = New tr_keys
Set ctr_page = New tr_page
If InStr(","&adminlimitidrange&",", ",")<1 Then contrl_message "您没有访问权限，请尝试其他版块。", "", "", "", ""
isdel=request.QueryString("isdel")
if isnumeric(isdel) then
if cint(isdel)<>1 then isdel=0
else
isdel=0
end if
isdelstr=" and isdel="&isdel&""

  pageno = CInt(request.querystring("pageno"))
  If Not IsNumeric(pageno) Then pageno = 1
  If pageno<1 Then pageno = 1

action = Trim(request.Form("action"))
id = Trim(request.Form("id"))

Select Case action
  Case "add"
    Call ctr_keys.Add()
    Call admin_log(adminname, adminid, adminname&" 添加注册码成功。")
    Call contrl_message ("添加注册码成功。", "location.href", "submit1", "form1", "parent")
  Case "edit"
    Call ctr_keys.edit()
    Call admin_log(adminname, adminid, adminname&" 修改注册码成功。")
    Call contrl_message ("修改注册码成功。", "", "submit1", "", "")
  Case "del"
      Call delline("tr_keys", "id ="&id&"")
    Call admin_log(adminname, adminid, adminname&" 删除注册码成功。")
    Call contrl_message ("删除注册码成功。", "", "submit1", "", "")
  Case "dodel"
      Call delline("tr_keys", "id ="&id&"")
    Call admin_log(adminname, adminid, adminname&" 彻底删除注册码成功。")
    Call contrl_message ("彻底删除注册码成功。", "", "submit1", "", "")
  Case "unisdel"
      Call changevalue("tr_keys", "isdel=0", "id ="&id&"")
    Call admin_log(adminname, adminid, adminname&" 从回收站恢复注册码成功。")
    Call contrl_message ("从回收站恢复注册码成功。",  "location.href", "submit1", "", "parent")
End Select
%>
 <title>注册码管理<%=application(siternd & "55trsitename")%></title>
 <link rel="stylesheet" type="text/css" href="./skin/default/style.css">
 <base target="_self">
 <script type="text/javascript" src="./js/js.js"></script>
 <script type="text/javascript" src="./js/hover.js"></script>
 <script type="text/javascript" src="../js/JTimer.js"></script>
<link rel="shortcut icon" type="image/x-icon" href="../favicon.ico">
<link rel="bookmark" type="image/x-icon" href="../favicon.ico">
 </head>
 <body style="background:#fff; " class="w1">
<table width="100%" align="center" border="0" cellpadding="3" cellspacing="1" id="mytable" class="size3" >
   <tr class="tr1" >
    <td class="td_title" colspan="11"><!--#include file="incgoback.asp"-->■ 添加注册码</td>
  </tr>
   <tr class="font1">
    <td width="" align="center">操作</td>
    <td width="" align="center">注册码</td>
    <td width="" align="center">备注信息</td>
  </tr>
   <tr class="">
    <form name="form1" id="form1" target="ifr1" method="post" onSubmit="return check_add_keys(this)" >
       <td width="" align="center"><input type="hidden" value="add" name="action" id="action" />
        <input type="submit" value=" 添加 " name="submit1" id="submit1" class="button2"/></td>
       <td width="" align="center"><input name="keys" type="text" class="input2 size19" id="keys" value="" maxlength="40" /></td>
       <td width="" align="center"><input name="notes" id="notes" type="text" class="input2 size19" value="" /></td>
     </form>
  </tr>
 </table>
<br>
<table width="100%" align="center" border="0" cellpadding="3" cellspacing="1" id="mytable2" class="size3" >
   <tr class="tr1" >
    <td class="td_title" colspan="12">■ <%if isdel=1 then%>回收站 <%end if%>编辑注册码信息</td>
  </tr>
   <tr class="font1">
    <td width="" align="center">操作</td>
    <td width="" align="center">注册码</td>
    <td width="" align="center">备注信息</td>
  </tr>
   <%
page_size = 20
pagei = 0
n = (pageno -1) * page_size
sql = "select * from tr_keys where id is not null "&isdelstr&" order by addtime asc , id asc "
Call ctr_page.create_rs(sql, rs, conn, page_size, pageno)
If Not rs.EOF And Not rs.bof Then
  Do While Not rs.EOF And pagei<page_size
    pagei = pagei + 1
    id = Rs("id")
    addtime = Rs("addtime")
    keys = Rs("keys")
    notes = Rs("notes")
%>
   <tr class="" id="tr<%=id%>">
    <form name="form2<%=id%>" id="form2<%=id%>" target="ifr1" method="post" onSubmit="return check_add_keys(this)" >
       <td width="" align="center"><input type="hidden" value="edit" name="action" id="action<%=id%>" />
        <input type="hidden" value="<%=id%>" name="id" id="id" />
        <%if isdel=0 then%><input type="submit" value=" 修改 " name="submit1" id="submit1" class="button2"/> <input type="button" value=" 删除 " name="submit2" id="submit2" class="button3 mar2" onClick="acdel(this.form,'tr<%=id%>')"/><%else%><input type="submit" value="恢复" name="submit1" id="submit1" class="button2" onClick="changev('action<%=id%>','unisdel')"> <input type="button" value=" 彻底删除 " name="submit2" id="submit2" class="button3 mar2" onClick="acdodel(this.form,'tr<%=id%>')"/><%end if%></td>
       <td width="" align="center"><input name="keys" type="text" class="input2 size19" id="keys" value="<%=keys%>" maxlength="40" /></td>
       <td width="" align="center"><input name="notes" id="notes" type="text" class="input2 size19" value="<%=notes%>"/></td>
     </form>
  </tr>
<%
rs.movenext
Loop
%>
<%
Else
%>
   <tr>
    <td colspan="12" align="center">暂无相关项目，请先添加。 </td>
  </tr>
   <%end if%>
	<tr>
		<%if ctr_page.page_count>1 then%><td colspan="10" align="left"><%= ctr_page.create_page("",ctr_page.page_count,pageno,"trpage")%> </td><% end if%>
	</tr>
   <tr>
    <td colspan="12" align="left"><p>特别提示：<br>
1、注册码用于激活从官网获取到的功能模块、插件、模板，官网会长期更新各种免费功能模块、模板、插件欢迎关注 。<br>
2、只要正确输入注册码即可，备注信息是为了让管理员分清注册码的归属可随意填写，系统会自动检索到注册码。<br>
3、每个不同的功能模块或插件都对应不同的注册码，注册码是根据你的网站一级域名生成的。<br>
    </p></td>
  </tr>
 </table>
<script type="text/javascript" >
showtable("mytable")
showtable("mytable2")
function changev(acid,value){
document.getElementById(acid).value	=value
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
