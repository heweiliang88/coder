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
appid = tr_killstr(request.querystring("appid"), 1, 8, 3, "", "", "", "", 2)
if not isnumeric(appid) then response.End()
If id<>"" And adminctrllimit = 1 Then Call contrl_message ("您只有查看权限，没有管理权限，请升级权限后继续操作。", "", "", "", "")

sql = "select * from tr_myapp where appid="&appid&" "
rs.Open sql, conn, 0, 1
If Not rs.EOF Then

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
'    Call contrl_message ("ID为："&id&" 的应用删除成功。", "parent.location.href", "", "", "parent")
  else
  If isopen =1 Then
'    Call contrl_message ("ID为："&id&" 的应用正在使用，请切换到其他应用再删除。", "", "", "", "")
  Else
call deltemfun170205(id)
    Call delline("tr_myapp", "id="&id&"")
    Call delfileapp("./app"&insurl)
'    Call contrl_message ("ID为："&id&" 的应用删除成功。", "parent.location.href", "", "", "parent")
  End If
  End If

else'不为模板时
sql="delete from tr_myapp where id="&id&" "'810725
conn.execute(sql)
call deltemfun170205(id)
end if
Else
  'Call contrl_message ("ID为："&id&" 的应用不存在，刷新页面后若仍然提示不存在则说明此应用已删除。", "", "", "", "")
End If
rs.Close

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
</html>