<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%' Option Explicit %>
<% On Error Resume Next %>
<!--#include file="../inc/charsetasp.asp"-->
<% Server.ScriptTimeout=99999999 %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<%
Const dbpath = "../"
Const funpath = "../"
%>
<!--#include file="admin_fun.asp"-->

<%
If InStr(","&adminlimitidrange&",", ",")<1 Then contrl_message "您没有访问权限，请尝试其他版块。", "", "", "", ""
id = CInt(request.querystring("id"))
  If id<>"" And adminctrllimit = 1 Then Call contrl_message ("您只有查看权限，没有管理权限，请升级权限后继续操作。", "", "", "", "")
filename=request.QueryString("inspath")
inspath = "app/"&filename
appid = trim(request.QueryString("appid"))

if isnumeric(appid) then
filestr = filepathHttpPage("http://www.55tr.com/outapp.asp?id="&appid&"")
If InStr(filestr, "|")>0 Then
  spfile = Split(filestr, "|")
  appname = spfile(1)
  appv = spfile(2)
  apptype = spfile(3)
End If
    sql = "select * from tr_myapp where appid="&appid&""
    rs.Open sql, conn, 1, 3
    If rs.EOF Then
      rs.addnew()
      id = GetMaxValue("tr_myapp", "id") + 1
      Rs("id") = id
      Rs("appname") = appname
      Rs("appv") = appv
      Rs("addtime") = FormatDate(Now(), 1)
      'Rs("unstr")=unstr
      Rs("apptype") = apptype
      Rs("appid") = appid
      If apptype = 2 Then
        Rs("insurl") = filename
      End If
      If apptype<>2 Then
        Rs("insurl") = filename
      End If
      Rs("insactive") = 0
      rs.update
    End If
    rs.Close
end if
%>
<title>天人系列管理系统在线部署工具</title>
</head>

<body>
<%
Dim strLocalPath
'得到当前文件夹的物理路径
strLocalPath = Server.MapPath("../")&"\"
adminpathdir = Server.MapPath("./")&"\"
Dim objXmlFile
Dim objNodeList
Dim objFSO
Dim objStream
Dim i, j

Set objXmlFile = Server.CreateObject("Microsoft.XMLDOM")
objXmlFile.load(Server.MapPath(inspath))

If objXmlFile.readyState = 4 Then
  If objXmlFile.parseError.errorCode = 0 Then

    Set objNodeList = objXmlFile.documentElement.selectNodes("//folder/path")
    Set objFSO = CreateObject("Scripting.FileSystemObject")

    j = objNodeList.Length -1
    For i = 0 To j
	objpathstr=objNodeList(i).text
	if instr(objpathstr,"|self|")>0 then
	objpathstr=replace(objpathstr,"|self|","")
	
      If objFSO.FolderExists(adminpathdir & objpathstr) = False Then
        objFSO.CreateFolder(adminpathdir & objpathstr)
      End If
	  
	  else
      If objFSO.FolderExists(strLocalPath & objNodeList(i).text) = False Then
        objFSO.CreateFolder(strLocalPath & objNodeList(i).text)
      End If
	  
	  end if
    Next
    Set objFSO = Nothing
    Set objNodeList = Nothing
    Set objNodeList = objXmlFile.documentElement.selectNodes("//file/path")

    j = objNodeList.Length -1
    For i = 0 To j
      Set objStream = CreateObject("ADODB.Stream")
      With objStream
        .Type = 1
        .Open
        .Write objNodeList(i).nextSibling.nodeTypedvalue
			objfilestr=objNodeList(i).text
	if instr(objfilestr,"|self|")>0 then
	objfilestr=replace(objfilestr,"|self|","")
        .SaveToFile adminpathdir & objfilestr, 2
else
        .SaveToFile strLocalPath & objNodeList(i).text, 2
end if
        Response.Write "释放文件" & objNodeList(i).text & "<br/>"
        .Close
      End With
      Set objStream = Nothing
    Next
    Set objNodeList = Nothing
  End If
End If

Set objXmlFile = Nothing
apath = "app/"
bpath = server.mappath(apath&"do.asp")
cpath = server.mappath(apath&"do0.asp")

Set objFSO = CreateObject("Scripting.FileSystemObject")
If objFSO.FileExists(cpath) Then
  response.redirect("./app/do0.asp?appid="&appid&"")
Else
  response.Write "应用源文件[ "&cpath&" ]不存在，请重新下载安装"
  response.End()
End If
Set objFSO = Nothing

Function filepathHttpPage(HttpUrl)
  Dim http
  Set http = server.CreateObject("MSX"&"ML2.Serv"&"erXM"&"LHTTP")
  Http.Open "GET", HttpUrl, False
  Http.send()
  If Http.readystate<>4 Then Exit Function
  filepathHttpPage = htmlBytesToBstr(Http._
  ResponseBody, "ut"&"f-8")
  Set http = Nothing
  If Err.Number<>0 Then Err.Clear
End Function

Function htmlBytesToBstr(Body, Cset)
  Dim Objstream
  Set Objstream = Server.CreateObject("ad"&"odb.St"&"ream")
  objstream.Type = 1
  objstream.Mode = 3
  objstream.Open
  objstream.Write body
  objstream.Position = 0
  objstream.Type = 2
  objstream.Charset = Cset
  If tmpcopy = "" Then
    htmlBytesToBstr = objstream.ReadText
  Else
    htmlBytesToBstr = tmpcopy
  End If
  objstream.Close
  Set objstream = Nothing
End Function
%>
</body>
</html>