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
<!--#include file="../core/md5.asp"-->
<title>操作接收页</title>
<link rel="shortcut icon" type="image/x-icon" href="../favicon.ico">
<link rel="bookmark" type="image/x-icon" href="../favicon.ico">
</head>
<body>
<%
Call admin_log(adminname, adminid, adminname&" 退出登录。")
Response.Cookies("55trcms")("id") = ""
Response.Cookies(admincookieskey)("md5name") = ""
Response.Cookies(admincookieskey)("passwd") = ""
session("tr_cms") = ""
Set ctr_admin = Nothing
Call contrl_message("", "./", "", "", "top")
%>
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