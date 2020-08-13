<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../inc/charsetasp.asp"-->
<%
Const dbpath = "../"
Const funpath = "../"
%>
<!--#include file="admin_fun.asp"-->
<%
turl=GetholeUrl()
if turl<>"" then turl=replace(turl,"http://","")
if turl<>"" then turl=replace(turl,"https://","")
'appid=103
appid=request.querystring("appid")
if not isnumeric(appid) then appid=103
app_have_managepage=true
app_have_install=true
sql = "select top 1 * from tr_myapp where appid="&appid&" order by id desc "
rs.open sql,conn,0,1
if not rs.eof then
managepage = trim(Rs("managepage"))
if managepage<>"" then
spmanagepage=split(managepage,",")
if instr(spmanagepage(0),"|")>0 then
spmnsi=split(spmanagepage(0),"|")
appurl=spmnsi(0)
response.Redirect(appurl)
else
app_have_managepage=false

end if
else
app_have_managepage=false
end if
else
app_have_install=false
end if
rs.close

if not app_have_managepage then
toappid=request.querystring("toappid")
if isnumeric(toappid) then 
sql = "select top 1 * from tr_myapp where appid="&toappid&" order by id desc "
rs.open sql,conn,0,1
if not rs.eof then
managepage = trim(Rs("managepage"))
if managepage<>"" then
spmanagepage=split(managepage,",")
if instr(spmanagepage(0),"|")>0 then
spmnsi=split(spmanagepage(0),"|")
appurl=spmnsi(0)
response.Redirect(appurl)
end if
else
app_have_managepage=false
end if
else
'app_have_install=false
end if
rs.close
end if
end if
%>

<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>高级功能</title>
<%
if app_have_managepage=false then
%>
<div class="ckgw"><a href="http://www.55tr.com/ifmappshow.asp?id=<%=appid%>&u=<%=turl%>&t=9&v=5.05&#xxjs"> 请点此查看该应用的控制方法</a></div>
<%end if%>
<%
if app_have_install=false then
%>
<meta http-equiv="Refresh" content="0;url=http://www.55tr.com/ifmappshow.asp?id=<%=appid%>&u=<%=turl%>&t=9&v=5.05&fr=&cn1=&cn2=&cn3=&cn4=&cn5=" />
<%end if%>
<style>
.ckgw{width:100%; height:50px; text-align:center; margin-top:30px; }
.ckgw a{display:block;width:400px; padding:20px;  margin:auto; background:#09F; border-radius:5px;color:#fff;}
</style>
</head>
<body>
</body>
</html>
<%
Public Function GetholeUrl()
  GetholeUrl =  Request.ServerVariables("SERVER_NAME")
  If Request.ServerVariables("SERVER_PORT") <> 80 Then GetholeUrl = GetholeUrl &":" & Request.ServerVariables("SERVER_PORT")
  GetholeUrl = lcase(GetholeUrl & Request.ServerVariables("URL"))
GetholeUrl=replace(GetholeUrl,"http://","")
End Function 

%>