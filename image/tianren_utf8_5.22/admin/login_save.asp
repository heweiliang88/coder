<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../inc/charsetasp.asp"-->
<!DOCTYPE html>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<%
Call checkoutside()
Const dbpath = "../"
%>
<!--#include file="../core/conn.asp"-->
<!--#include file="../core/class/tr_admin.asp"-->
<!--#include file="../core/fun/fun.asp"-->
<!--#include file="../core/md5.asp"-->
<title>操作接收页</title>
</head>
<body>
<%
Call checkoutside()
ad_name = LCase(tr_killstr(request.Form("ad_name"), 5, 30, 2, "./", "", "", "top", 2))
ad_psw=request.Form("ad_psw")
if trim(ad_psw)="" then Call contrl_message("密码不能为空，请输入", "./", "", "", "top")
ad_psw=left(ad_psw,30)
ver_code = tr_killstr(request.Form("ver_code"), 1, 4, 2, "./", "", "", "top", 2)
Set ctr_admin = New tr_admin
Call ctr_admin.tr_admin_login (ad_name, ad_psw, ver_code)
Set ctr_admin = Nothing
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