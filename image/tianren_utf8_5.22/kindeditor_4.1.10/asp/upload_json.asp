<%@ CODEPAGE=65001 %>
<%' Option Explicit %>
<% Response.CodePage=65001 %>
<% Response.Charset="UTF-8" %>
<%
Const dbpath = "../../"
Const funpath = "../"
%>
<!--#include file="UpLoad_Class.asp"-->
<!--#include file="JSON_2.0.4.asp"-->
<!--#include file="../../core/Conn.asp"-->
<!--#include file="../../core/class/tr_admin.asp"-->
<!--#include file="../../core/fun/fun.asp"-->
<%
Set ctr_admin = New tr_admin
set upload = new AnUpLoad
Dim aspUrl, savePath, saveUrl, maxSize, fileName, fileExt, newFileName, filePath, fileUrl, dirName
Dim extStr, imageExtStr, flashExtStr, mediaExtStr, fileExtStr
Dim upload, file, fso, ranNum, hash, ymd, mm, dd, result
'定义允许上传的文件扩展名
imageExtStr = "gif|jpg|jpeg|png|bmp"
flashExtStr = "swf|flv"
mediaExtStr = "swf|flv|mp3|wav|wma|wmv|mid|avi|mpg|asf|rm|rmvb"
fileExtStr = "doc|docx|xls|xlsx|ppt|htm|html|txt|zip|rar|gz|bz2|tr"
'最大文件大小
maxSize = 5 * 1024 * 1024 '5M
dirName = Request.QueryString("dir")
If isEmpty(dirName) Then
	dirName = "image"
End If
If instr(lcase("image,flash,media,file"), dirName) < 1 Then
	showError("目录名不正确。")
End If

Select Case dirName
	Case "flash" extStr = flashExtStr
	Case "media" extStr = mediaExtStr
	Case "file" extStr = fileExtStr
	Case Else  extStr = imageExtStr
End Select


upload.Exe = extStr
upload.MaxSize = maxSize
upload.GetData()
if upload.ErrorID>0 then 
	showError(upload.Description)
end if


aspUrl = Request.ServerVariables("SCRIPT_NAME")
aspUrl = left(aspUrl, InStrRev(aspUrl, "/"))
isupapp=upload.forms("isupapp")
PATH_INFO=upload.forms("PATH_INFO")
isswf=upload.forms("isswf")
if not isnumeric(isswf) then isswf =0
'if isupapp="" then isupapp=0
if not isnumeric(isupapp) then isupapp=0

'文件保存目录路径
if isupapp=1 then
if len(ctr_admin.checkadminlogin())<=1 then response.End()
savePath = "../../"&PATH_INFO&"/app/"
'文件保存目录URL
saveUrl = aspUrl & "../../"&PATH_INFO&"/app/"
else
savePath = "../../upfiles/"
'文件保存目录URL
saveUrl = aspUrl & "../../upfiles/"
end if
Set fso = Server.CreateObject("Scripting.FileSystemObject")
If Not fso.FolderExists(Server.mappath(savePath)) Then
	showError("上传目录不存在。")
End If

'if isnumeric(upload.forms("isswf")) then
if isswf=1 then
if len(ctr_admin.checkadminloginswf())<=1 then response.End()
else
if len(ctr_admin.checkadminlogin())<=1 then response.End()
end if
'end if

'创建文件夹
if isupapp<>1 then

savePath = savePath & dirName & "/"
saveUrl = saveUrl & dirName & "/"

If Not fso.FolderExists(Server.mappath(savePath)) Then
	fso.CreateFolder(Server.mappath(savePath))
End If
mm = month(now)
If mm < 10 Then
	mm = "0" & mm
End If
'dd = day(now)
'If dd < 10 Then
'	dd = "0" & dd
'End If
ymd = year(now) & mm '& dd
savePath = savePath & ymd & "/"
saveUrl = saveUrl & ymd & "/"
If Not fso.FolderExists(Server.mappath(savePath)) Then
	fso.CreateFolder(Server.mappath(savePath))
End If
end if

set file = upload.files("imgFile")
if file is nothing then
	showError("请选择文件。")
end if

set result = file.saveToFile(savePath, 0, true)
if result.error then
	showError(file.Exception)
end if
filePath = Server.mappath(savePath & file.filename)
fileUrl = saveUrl & file.filename

'spfileurl=split(fileurl,"/")
'ubfileurl=ubound(spfileurl)
'fori=ubfileurl-4
'for i=0  to 3
'tmpurl=tmpurl&"/"&spfileurl(fori+i)
'next
'fileurl=tmpurl
Set upload = nothing
Set file = nothing
If Not fso.FileExists(filePath) Then
	showError("上传文件失败。")
End If
Response.AddHeader "Content-Type", "text/html; charset=UTF-8"
Set hash = jsObject()
hash("error") = 0
hash("url") = fileUrl
hash.Flush
Response.End

Function showError(message)
	Response.AddHeader "Content-Type", "text/html; charset=UTF-8"
	Dim hash
	Set hash = jsObject()
	hash("error") = 1
	hash("message") = message
	hash.Flush
	Response.End
End Function
%>