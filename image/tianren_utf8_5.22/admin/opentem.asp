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
<%
If InStr(","&adminlimitidrange&",", ",")<1 Then contrl_message "您没有访问权限，请尝试其他版块。", "", "", "", ""
id = tr_killstr(request.querystring("id"), 1, 8, 3, "", "", "", "", 2)
  If id<>0 And adminctrllimit = 1 Then Call contrl_message ("您只有查看权限，没有管理权限，请升级权限后继续操作。", "", "", "", "")
sql="select * from tr_myapp where id="&id&" "
rs.open sql,conn,1,3
if not rs.eof then
insurl=rs("insurl")
appid=rs("appid")
rs("isopen")=1
rs.update
end if
rs.close

sql="update tr_myapp set isopen=0 where id<>"&id&" and apptype=2 "
conn.execute(sql)
if insurl<>"" then
response.redirect("uncompress.asp?inspath="&insurl&"&appid="&appid&"")
end if
%>
</head>
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