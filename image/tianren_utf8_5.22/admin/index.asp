<!--#include file="default.asp"-->
<%
'本文件为扩展首页，不会影响你的网站收录与其他任何方面，只为更广泛的适应不同服务器的默认首页设置，请勿删除。
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
