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
 <%
If InStr(","&adminlimitidrange&",", ",1,")<1 Then contrl_message "您没有访问权限，请尝试其他版块。", "", "", "", ""
action = Trim(request.Form("action"))
If action<>"" And adminctrllimit = 1 Then Call contrl_message ("您只有查看权限，没有管理权限，请升级权限后继续操作。", "", "", "", "")
fileurl = server.mappath(dburl)
folderurl = Left(fileurl, InStrRev(fileurl, "\"))&"db_backup\"
Select Case action
  Case "backup"
    db_backupname = tr_killstr(Replace(request.Form ("db_backupname"), ".asp", ""), 1, 20, 2, "", "submit1", "", "", 2)
    Call CopyNewFile(fileurl, folderurl&db_backupname&".asp")
    Call admin_log(adminname, adminid, adminname&" 备份网站数据库"&folderurl&db_backupname&"成功。")
    Call contrl_message (" 备份网站数据库"&folderurl&db_backupname&"成功。", "location.href", "submit1", "", "parent")
  Case "restore"
    restore_file = tr_killstr(Replace(request.Form ("restore_file"), ".asp", ""), 1, 20, 2, "", "submit2", "", "", 2)
    Call CopyNewFile(folderurl&restore_file&".asp", fileurl)
    Call admin_log(adminname, adminid, adminname&" 还原网站数据库"&folderurl&restore_file&"成功。")
    Call contrl_message (" 还原网站数据库"&folderurl&restore_file&"成功。", "", "submit2", "", "")
End Select
%>
 <title>数据库管理<%=application(siternd & "55trsitename")%></title>
 <link rel="stylesheet" type="text/css" href="./skin/default/style.css">
 <base target="_self">
 <script type="text/javascript" src="./js/js.js"></script>
  <script type="text/javascript" src="./js/hover.js"></script>
<link rel="shortcut icon" type="image/x-icon" href="../favicon.ico">
<link rel="bookmark" type="image/x-icon" href="../favicon.ico">
</head>
 <body style="background:#fff; " class="w1">
<table width="100%" align="center" border="0" cellpadding="3" cellspacing="1" id="mytable" class="size3" >
   <tr class="tr1" >
    <td class="td_title" colspan="5"><!--#include file="incgoback.asp"-->■ 备份数据库</td>
  </tr>
   <tr>
    <td width="30%" align="right">现用数据库：</td>
    <td colspan="3" width="99%"><%=fileurl%> 数据库大小：<%= GetfileSize(fileurl) %></td>
  </tr>
   <form target="ifr1" method="post" name="form1" id="form1" onSubmit="return check_dbbackup(this)" >
    <tr>
       <td align="right">备份数据库命名：</td>
       <td colspan="3"><input name="db_backupname" type="text" class="input2 size1" id="db_backupname" value="<%=FormatDate(now(), 14) &gen_key(3)&".asp"%>" readonly /></td>
     </tr>
    <tr>
       <td align="right">&nbsp;</td>
       <td colspan="3"><font color="#FF0000">请及时关注您的空间剩余情况以免出现无法备份数据库的问题！<br>
当数据库较大时备份数据库可能会需要数分钟或数十分钟时间。</font>&nbsp;</td>
     </tr>
    <tr>
       <td align="right">&nbsp;</td>
       <td colspan="3"><input type="hidden" value="backup" name="action" id="action" />
        <input type="submit" value=" 立即备份 " name="submit1" id="submit1" class="button1"/></td>
     </tr>
    <tr>
       <td align="right">&nbsp;</td>
       <td colspan="3">&nbsp;</td>
     </tr>
  </form>
   <tr class="tr1" >
    <td class="td_title" colspan="5">■ 还原数据库</td>
  </tr>
   <form target="ifr1" method="post" name="form2" id="form2" onSubmit="return check_dbrestore(this)" >
    <tr>
       <td align="right">选择要还原的数据库：<br><font color="#FF0000">还原前请务必备份现用数据库！</font><br>删除多余的数据库请使用ftp或远程桌面连接。
</td>
       <td colspan="3"><select name="restore_file" size="10">
<%
ishavebackup = false
Set fs = CreateObject("Scripting.FileSystemObject")
Set f = fs.GetFolder(folderurl)'文件夹的名称
Set fc = f.Files
For Each f1 in fc
  response.Write "<option value="&f1.Name&">"&f1.Name&"&nbsp;&nbsp;"& GetfileSize(folderurl&f1.Name)
  response.Write "<br>"
  ishavebackup = true
Next
If ishavebackup = false Then
  response.Write "<option value="""">目前没有可还原的数据库文件，请先进行数据库的备份"
End If
%>
         </select></td>
     </tr>
    <tr>
       <td align="right">&nbsp;</td>
       <td colspan="3"><input type="hidden" value="restore" name="action" id="action" />
        <input type="submit" value=" 立即还原 " name="submit2" id="submit2" class="button7"/></td>
     </tr>
  </form>
   <tr>
    <td align="right">&nbsp;</td>
    <td colspan="3">&nbsp;</td>
  </tr>
 </table>
<script type="text/javascript" >
showtable("mytable")
</script>
<!--#include file="../inc/ifr.asp"-->
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
