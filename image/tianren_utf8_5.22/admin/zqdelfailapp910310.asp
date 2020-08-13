<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../inc/charsetasp.asp"-->
<!DOCTYPE html>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%
Const dbpath = "../"
Const funpath = "../"
charsetstr = "utf-8"
%>
<!--#include file="admin_fun.asp"-->
<%
If InStr(","&adminlimitidrange&",", ",")<1 Then contrl_message "您没有访问权限，请尝试其他版块。", "", "", "", ""
dodelfailapp = trim(request.querystring("dodelfailapp"))
if not isnumeric(dodelfailapp) then dodelfailapp=0
If id<>"" And adminctrllimit = 1 Then Call contrl_message ("您只有查看权限，没有管理权限，请升级权限后继续操作。", "", "", "", "")

if dodelfailapp=1 then
if data_type="access" then
sql="select * from tr_myapp where insactive=0 and addtime < #"&FormatDate(DateAdd("n", -30, now()),1)&"# "
elseif data_type="mssql" then
sql="select * from tr_myapp where insactive=0 and addtime < '"&FormatDate(DateAdd("n", -30, now()),1)&"' "
end if


rs.Open sql, conn, 0, 1
If Not rs.EOF Then
do while not rs.eof
  id = rs("id")
  insurl = rs("insurl")
  isopen = rs("isopen")
  insactive = rs("insactive")
  unstr = rs("unstr")
  Folders = rs("folders")
  apptype = rs("apptype")

if apptype=2 then'为模板时
  if insactive=0 then
  call deltemfun170205(id)
    Call delline("tr_myapp", "id="&id&"")
	response.write "失效应用卸载成功<br>"
'    Call contrl_message ("ID为："&id&" 的应用删除成功。", "parent.location.href", "", "", "parent")
  else
  If isopen =1 Then
'    Call contrl_message ("ID为："&id&" 的应用正在使用，请切换到其他应用再删除。", "", "", "", "")
  Else
call deltemfun170205(id)
    Call delline("tr_myapp", "id="&id&"")
    Call delfileapp("./app"&insurl)
	response.write "失效应用卸载成功<br>"
'    Call contrl_message ("ID为："&id&" 的应用删除成功。", "parent.location.href", "", "", "parent")
  End If
  End If

else'不为模板时
sql810725="delete from tr_myapp where id="&id&" "'810725
conn.execute(sql810725)
call deltemfun170205(id)
	response.write "失效应用卸载成功<br>"
end if

rs.movenext
loop
else
	response.write "失效应用卸载成功<br>"

End If
rs.Close

end if


function deltemfun170205(id)
if trim(Folders)="" or isnull(trim(Folders)) then exit function
spfolders = Split(Folders, ",")
ubfolders = UBound(spfolders)
For nf = 0 To ubfolders
  spurl = Split(spfolders(nf), "|")
  if ubound(spurl)>1 then
  if spurl(2)<>"" then
  charsetstr=spurl(2)
  end if
  end if
  If InStr(spfolders(nf), "|del")>0 Then
    If spurl(0)<>"" Then
      Call delfileapp(spurl(0))
    End If
  Else
    If spurl(0)<>"" Then
      unfile = ReadFromTextFileun (spurl(0), charsetstr)
      If InStr(unfile, unstr)>0 Then
        patrnstrasp = "<"&Chr(37)&"'"&unstr&Chr(37)&">[\s\S]*?<"&Chr(37)&"'"&unstr&Chr(37)&">"
        instrpatrnstrasp = "<"&Chr(37)&"'"&unstr&Chr(37)&">"
        patrnstrjs = "\/\/"&unstr&"[\s\S]*?\/\/"&unstr&""
        instrpatrnstrjs = "//"&unstr&""
        patrnstrcss = "\/\*"&unstr&"\*\/[\s\S]*?\/\*"&unstr&"\*\/"
        instrpatrnstrcss = "/*"&unstr&"*/"
        If InStr(unfile, instrpatrnstrasp)>0 Then
          unfile = Replacestrun(unfile, patrnstrasp, "")
        ElseIf InStr(unfile, instrpatrnstrjs)>0 Then
          unfile = Replacestrun(unfile, patrnstrjs, "")
        ElseIf InStr(unfile, instrpatrnstrcss)>0 Then
          unfile = Replacestrun(unfile, patrnstrcss, "")
        End If
        Call createfile(unfile, spurl(0), 2, charsetstr)
      End If
    End If
  End If
Next
end function
%>
</head>
<body>
<a href="zqdelfailapp910310.asp?dodelfailapp=1" style="display:block;width:300px; height:40px; line-height:40px; margin:10px;color:#FFF;background:#09F;font-size:14px;text-align:center; text-decoration:none;" target="_self">立即卸载全部失效应用</a>


</body>
</html>