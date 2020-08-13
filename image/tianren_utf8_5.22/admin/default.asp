<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../inc/charsetasp.asp"-->
<!--#include file="../core/fun/fun.asp"-->
<%
strs = gen_key(16)
If session("tr_cms")<>"55trcms_8u73gdyf982" Then
  response.redirect ("login.asp?r="&strs&"")
Else
  response.redirect ("ad_index.html?r="&strs&"")
End If
%>
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
