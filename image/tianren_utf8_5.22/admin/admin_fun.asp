<!--#include file="../core/Conn.asp"-->
<!--#include file="../core/class/tr_admin.asp"-->
<!--#include file="../core/fun/fun.asp"-->
<%
Set ctr_admin = New tr_admin

checkadminloginvalue = ctr_admin.checkadminlogin
spcheckadminlogin = Split(checkadminloginvalue, "|")
adminname = spcheckadminlogin(0)
adminctrllimit = spcheckadminlogin(1)
adminlimitidrange = spcheckadminlogin(2)
adminid = spcheckadminlogin(3)
admincookieskey = spcheckadminlogin(4)
%>
