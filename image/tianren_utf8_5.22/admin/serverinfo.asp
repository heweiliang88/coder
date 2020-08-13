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
<script language=javascript>
 <!--
 var startTime,endTime;
 var d=new Date();
 startTime=d.getTime();
 //-->
 </script>
<%
If InStr(","&adminlimitidrange&",", ",")<1 Then contrl_message "您没有访问权限，请尝试其他版块。", "", "", "", ""
Dim starttime
starttime = Timer() * 1000
'声明待检数组
Dim ObjTotest(39, 4)
ObjTotest(0, 0) = "MSWC.AdRotator"
ObjTotest(1, 0) = "MSWC.BrowserType"
ObjTotest(2, 0) = "MSWC.NextLink"
ObjTotest(3, 0) = "MSWC.Tools"
ObjTotest(4, 0) = "MSWC.Status"
ObjTotest(5, 0) = "MSWC.Counters"
ObjTotest(6, 0) = "IISSample.ContentRotator"
ObjTotest(7, 0) = "IISSample.PageCounter"
ObjTotest(8, 0) = "MSWC.PermissionChecker"
ObjTotest(9, 0) = "Scripting.FileSystemObject"
ObjTotest(9, 1) = "(FSO 文本文件读写)"
ObjTotest(10, 0) = "adodb.connection"
ObjTotest(10, 1) = "(ADO 数据对象)"
ObjTotest(11, 0) = "DownloadClassAsp.DownSysObject"
ObjTotest(11, 1) = "(下载管理系统组件)"
ObjTotest(12, 0) = "SoftArtisans.FileUp"
ObjTotest(12, 1) = "(SA-FileUp 文件上传)"
ObjTotest(13, 0) = "SoftArtisans.FileManager"
ObjTotest(13, 1) = "(SoftArtisans 文件管理)"
ObjTotest(14, 0) = "LyfUpload.UploadFile"
ObjTotest(14, 1) = "(刘云峰的文件上传组件)"
ObjTotest(15, 0) = "Persits.Upload.1"
ObjTotest(15, 1) = "(ASPUpload 文件上传)"
ObjTotest(16, 0) = "w3.upload"
ObjTotest(16, 1) = "(Dimac 文件上传)"
ObjTotest(17, 0) = "JMail.SmtpMail"
ObjTotest(17, 1) = "(Dimac JMail 邮件收发) "
ObjTotest(18, 0) = "CDONTS.NewMail"
ObjTotest(18, 1) = "(虚拟 SMTP 发信)"
ObjTotest(19, 0) = "Persits.MailSender"
ObjTotest(19, 1) = "(ASPemail 发信)"
ObjTotest(20, 0) = "SMTPsvg.Mailer"
ObjTotest(20, 1) = "(ASPmail 发信)"
ObjTotest(21, 0) = "DkQmail.Qmail"
ObjTotest(21, 1) = "(dkQmail 发信)"
ObjTotest(22, 0) = "Geocel.Mailer"
ObjTotest(22, 1) = "(Geocel 发信)"
ObjTotest(23, 0) = "IISmail.Iismail.1"
ObjTotest(23, 1) = "(IISmail 发信)"
ObjTotest(24, 0) = "SmtpMail.SmtpMail.1"
ObjTotest(24, 1) = "(SmtpMail 发信)"
ObjTotest(25, 0) = "SoftArtisans.ImageGen"
ObjTotest(25, 1) = "(SA 的图像读写组件)"
ObjTotest(26, 0) = "W3Image.Image"
ObjTotest(26, 1) = "(Dimac 的图像读写组件)"
ObjTotest(27, 0) = "adodb.Stream"
ObjTotest(27, 1) = "(adodb.stream 组件)"
ObjTotest(28, 0) = "Msxml2.ServerXMLHTTP.6.0"
ObjTotest(28, 1) = "(XmlHttp 组件)"
ObjTotest(29, 0) = "Msxml2.ServerXMLHTTP.5.0"
ObjTotest(29, 1) = "(XmlHttp 组件)"
ObjTotest(30, 0) = "Msxml2.ServerXMLHTTP.4.0"
ObjTotest(30, 1) = "(XmlHttp 组件)"
ObjTotest(31, 0) = "Msxml2.ServerXMLHTTP.3.0"
ObjTotest(31, 1) = "(XmlHttp 组件)"
ObjTotest(32, 0) = "Msxml2.ServerXMLHTTP"
ObjTotest(32, 1) = "(XmlHttp 组件)"
ObjTotest(33, 0) = "Msxml2.XMLHTTP.6.0"
ObjTotest(33, 1) = "(XmlHttp 组件)"
ObjTotest(34, 0) = "Msxml2.XMLHTTP.5.0"
ObjTotest(34, 1) = "(XmlHttp 组件)"
ObjTotest(35, 0) = "Msxml2.XMLHTTP.4.0"
ObjTotest(35, 1) = "(XmlHttp 组件)"
ObjTotest(36, 0) = "Msxml2.XMLHTTP.3.0"
ObjTotest(36, 1) = "(XmlHttp 组件)"
ObjTotest(37, 0) = "Msxml2.XMLHTTP"
ObjTotest(37, 1) = "(XmlHttp 组件)"
ObjTotest(38, 0) = "Msxml2"
ObjTotest(38, 1) = "(XmlHttp 组件)"
ObjTotest(39, 0) = "Persits.Jpeg"
ObjTotest(39, 1) = "(Persits.Jpeg 水印组件)"

Public IsObj, VerObj, TestObj
'检查预查组件支持情况及版本
Dim i
For i = 0 To 38
  On Error Resume Next
  IsObj = false
  VerObj = ""
  'dim TestObj
  TestObj = ""
  Set TestObj = Server.CreateObject(ObjTotest(i, 0))
  If -2147221005 <> Err Then '感谢网友iAmFisher的宝贵建议
    IsObj = True
    VerObj = TestObj.Version
    If VerObj = "" Or IsNull(VerObj) Then VerObj = TestObj.about
  End If
  ObjTotest(i, 2) = IsObj
  ObjTotest(i, 3) = VerObj
  '	ObjTotest(i,3)=err.description
  '	ObjTotest(i,3)=err
Next
'检查组件是否被支持及组件版本的子程序

Sub ObjTest(strObj)
  On Error Resume Next
  IsObj = false
  VerObj = ""
  TestObj = ""
  Set TestObj = Server.CreateObject (strObj)
  If -2147221005 <> Err Then '感谢网友iAmFisher的宝贵建议
    IsObj = True
    VerObj = TestObj.Version
    If VerObj = "" Or IsNull(VerObj) Then VerObj = TestObj.about
  End If
End Sub
%>
<title>服务器信息 - 天人文章管理系统</title>
<link rel="stylesheet" type="text/css" href="./skin/default/style.css">
<base target="_self">
<script type="text/javascript" src="./js/js.js"></script>
<script type="text/javascript" src="./js/hover.js"></script>
<link rel="shortcut icon" type="image/x-icon" href="../favicon.ico">
<link rel="bookmark" type="image/x-icon" href="../favicon.ico">
</head>
<body style="background:#fff; " class="w1">
<table width="100%" align="center" border="0" cellpadding="3" cellspacing="1" id="mytable" class="size3" >
  <tr class="tr1" >
    <td class="td_title" colspan="4"><!--#include file="incgoback.asp"-->■ 服务器基本信息</td>
  </tr>
  <tr>
    <td colspan="1" align="right">&nbsp;服务器名：</td>
    <td colspan="1" >&nbsp;<%=Request.ServerVariables("SERVER_NAME")%></td>
    <td  colspan="1" align="right">&nbsp;服务器IP：</td>
    <td colspan="1">&nbsp;<%=Request.ServerVariables("LOCAL_ADDR")%></td>
  </tr>
  <tr>
    <td colspan="1" align="right">&nbsp;服务器端口：</td>
    <td colspan="1" >&nbsp;<%=Request.ServerVariables("SERVER_PORT")%></td>
    <td  colspan="1" align="right">服务器时间：</td>
    <td colspan="1">&nbsp;<%=now%></td>
  </tr>
  <tr>
    <td colspan="1" align="right">&nbsp;&nbsp;程序根目录路径：</td>
    <td colspan="1" >&nbsp;<% = server.MapPath("../")
%></td>
    <td  colspan="1" align="right">服务器CPU数量：</td>
    <td colspan="1">&nbsp;<%=Request.ServerVariables("NUMBER_OF_PROCESSORS")%> 个</td>
  </tr>
  <tr>
    <td colspan="1" align="right">&nbsp;&nbsp;服务器解译引擎：</td>
    <td colspan="1" >&nbsp;<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
    <td  colspan="1" align="right">服务器操作系统：</td>
    <td colspan="1">&nbsp;<%=Request.ServerVariables("OS")%></td>
  </tr>
  <tr>
    <td height="20" align="right" class="forumRowHighlight">&nbsp;IIS版本：</td>
    <td class="forumRow">&nbsp;<%=Request.ServerVariables("SERVER_SOFTWARE")%></td>
    <td  colspan="1" align="right">脚本超时时间：</td>
    <td colspan="1">&nbsp;<%=Server.ScriptTimeout%></td>
  </tr>
  <tr>
    <td colspan="1" align="right">&nbsp;</td>
    <td colspan="1" >&nbsp;</td>
    <td  colspan="1" align="right">&nbsp;</td>
    <td colspan="1">&nbsp;</td>
  </tr>
  <%If ObjTest("Scripting.FileSystemObject") Then
  Set fsoobj = server.CreateObject("Scripting.FileSystemObject")
%>
  <tr class="tr1">
    <td colspan="4" class="td_title" >■ 服务器磁盘信息</td>
  </tr>
  <tr height="20" align=center>
    <td >盘符和磁盘类型与状态</td>
    <td  >文件系统</td>
    <td  >可用空间</td>
    <td >总空间</td>
  </tr>
  <%
' 测试磁盘信息的想法来自"COCOON ASP 探针"
Set drvObj = fsoobj.Drives
For Each d in drvObj
%>
  <tr height="18" align=center>
    <td  align="right"><%=cdrivetype(d.DriveType) & " " & d.DriveLetter%>: <%=cIsReady(d.isReady)%></td>
    <%
If d.DriveLetter = "A" Then '为防止影响服务器，不检查软驱
Else
%>
    <td ><%=d.FileSystem%></td>
    <td align="right" ><%=cSize(d.FreeSpace)%></td>
    <td align="right" ><%=cSize(d.TotalSize)%></td>
    <%
End If
%>
  </tr>
<%
Next
%>
      <tr class="tr1" >
    <td class="td_title" colspan="4">■ 服务器组件支持情况</td>
  </tr>
  <tr>
    <td colspan="1" align="center">&nbsp;IIS自带的ASP组件：</td>
    <td colspan="1" align="center">&nbsp;组件名称</td>
    <td  colspan="1" align="center">&nbsp;是否支持</td>
    <td colspan="1" align="center">&nbsp;版本</td>
  </tr>
  	<%For i=0 to 10%>
  <tr>
    <td colspan="1" align="right">&nbsp;</td>
    <td colspan="1" >&nbsp;<%=ObjTotest(i,0) & "<font color=#888888>&nbsp;" & ObjTotest(i,1)%></td>
    <td  colspan="1" align="center">&nbsp;<%
If Not ObjTotest(i, 2) Then
  Response.Write "<font color=red><b>×</b></font>"
Else
  Response.Write "<font class=fonts><b>√</b></font>"
End If
%></td>
    <td colspan="1"><%
'		If Not ObjTotest(i,2) Then
'			Response.Write "<font color=red><b>×</b></font>"
'		Else
Response.Write "<a title='" & ObjTotest(i, 3) & "'>" & Left(ObjTotest(i, 3), 11) & "</a>"
'		End If
%></td>
  </tr>
  	<%next%>
  <%For i=27 to 38%>
  <tr>
    <td colspan="1" align="right">&nbsp;</td>
    <td colspan="1" >&nbsp;<%=ObjTotest(i,0) & "<font color=#888888>&nbsp;" & ObjTotest(i,1)%></td>
    <td  colspan="1" align="center">&nbsp;<%
If Not ObjTotest(i, 2) Then
  Response.Write "<font color=red><b>×</b></font>"
Else
  Response.Write "<font class=fonts><b>√</b></font>"
End If
%></td>
    <td colspan="1"><%
'		If Not ObjTotest(i,2) Then
'			Response.Write "<font color=red><b>×</b></font>"
'		Else
Response.Write "<a title='" & ObjTotest(i, 3) & "'>" & Left(ObjTotest(i, 3), 11) & "</a>"
'		End If
%></td>
  </tr>
  	<%next%>
  <tr>
    <td colspan="1" align="center">&nbsp;常见的文件上传和管理组件：</td>
    <td colspan="1" align="center">&nbsp;组件名称（如未特定开发功能不用在意）</td>
    <td  colspan="1" align="center">&nbsp;是否支持</td>
    <td colspan="1" align="center">&nbsp;版本</td>
  </tr>
  	<%For i=11 to 16%>
  <tr>
    <td colspan="1" align="right">&nbsp;</td>
    <td colspan="1" >&nbsp;<%=ObjTotest(i,0) & "<font color=#888888>&nbsp;" & ObjTotest(i,1)%></td>
    <td  colspan="1" align="center">&nbsp;<%
If Not ObjTotest(i, 2) Then
  Response.Write "<font color=red><b>×</b></font>"
Else
  Response.Write "<font class=fonts><b>√</b></font>"
End If
%></td>
    <td colspan="1"><%
'		If Not ObjTotest(i,2) Then
'			Response.Write "<font color=red><b>×</b></font>"
'		Else
Response.Write "<a title='" & ObjTotest(i, 3) & "'>" & Left(ObjTotest(i, 3), 11) & "</a>"
'		End If
%></td>
  </tr>
  	<%next%>
  <tr>
    <td colspan="1" align="center">&nbsp;常见的收发邮件组件：</td>
    <td colspan="1" align="center">&nbsp;组件名称（如未特定开发功能不用在意）</td>
    <td  colspan="1" align="center">&nbsp;是否支持</td>
    <td colspan="1" align="center">&nbsp;版本</td>
  </tr>
  	<%For i=17 to 24%>
  <tr>
    <td colspan="1" align="right">&nbsp;</td>
    <td colspan="1" >&nbsp;<%=ObjTotest(i,0) & "<font color=#888888>&nbsp;" & ObjTotest(i,1)%></td>
    <td  colspan="1" align="center">&nbsp;<%
If Not ObjTotest(i, 2) Then
  Response.Write "<font color=red><b>×</b></font>"
Else
  Response.Write "<font class=fonts><b>√</b></font>"
End If
%></td>
    <td colspan="1"><%
'		If Not ObjTotest(i,2) Then
'			Response.Write "<font color=red><b>×</b></font>"
'		Else
Response.Write "<a title='" & ObjTotest(i, 3) & "'>" & Left(ObjTotest(i, 3), 11) & "</a>"
'		End If
%></td>
  </tr>
  	<%next%>
  <tr>
    <td colspan="1" align="center">&nbsp;图像处理组件：</td>
    <td colspan="1" align="center">&nbsp;组件名称（如未特定开发功能不用在意）</td>
    <td  colspan="1" align="center">&nbsp;是否支持</td>
    <td colspan="1" align="center">&nbsp;版本</td>
  </tr>
  	<%For i=25 to 26%>
  <tr>
    <td colspan="1" align="right">&nbsp;</td>
    <td colspan="1" >&nbsp;<%=ObjTotest(i,0) & "<font color=#888888>&nbsp;" & ObjTotest(i,1)%></td>
    <td  colspan="1" align="center">&nbsp;<%
If Not ObjTotest(i, 2) Then
  Response.Write "<font color=red><b>×</b></font>"
Else
  Response.Write "<font class=fonts><b>√</b></font>"
End If
%></td>
    <td colspan="1"><%
'		If Not ObjTotest(i,2) Then
'			Response.Write "<font color=red><b>×</b></font>"
'		Else
Response.Write "<a title='" & ObjTotest(i, 3) & "'>" & Left(ObjTotest(i, 3), 11) & "</a>"
'		End If
%></td>
  </tr>
  	<%next%>
<%i=39
if i=39 then%>	
  <tr>
    <td colspan="1" align="right">&nbsp;</td>
    <td colspan="1" >&nbsp;<%=ObjTotest(i,0) & "<font color=#888888>&nbsp;" & ObjTotest(i,1)%></td>
    <td  colspan="1" align="center">&nbsp;<%
If Not ObjTotest(i, 2) Then
  Response.Write "<font color=red><b>×</b></font>"
Else
  Response.Write "<font class=fonts><b>√</b></font>"
End If
%></td>
    <td colspan="1"><%
'		If Not ObjTotest(i,2) Then
'			Response.Write "<font color=red><b>×</b></font>"
'		Else
Response.Write "<a title='" & ObjTotest(i, 3) & "'>" & Left(ObjTotest(i, 3), 11) & "</a>"
'		End If
%></td>
  </tr>
<%end if%>	
	
  <tr>
    <td colspan="1" align="right">&nbsp;</td>
    <td colspan="1" >&nbsp;</td>
    <td  colspan="1" align="right">&nbsp;</td>
    <td colspan="1">&nbsp;</td>
  </tr>
  <tr>
    <td colspan="1" align="right">&nbsp;</td>
    <td colspan="1" >&nbsp;</td>
    <td  colspan="1" align="right">&nbsp;</td>
    <td colspan="1">&nbsp;</td>
  </tr>
  <tr class="tr1">
    <td colspan="4" class="td_title" >■ 其他组件支持情况检测</td>
  </tr>
  <tr height=23 align=center>
    <td height="30"  colspan="4" align="left">
      <%
Dim strClass
strClass = Trim(Request.Form("classname"))
If "" <> strClass Then
  Response.Write "<br>您指定的组件的检查结果："
  Dim Verobj1
  ObjTest(strClass)
  If Not IsObj Then
    Response.Write "<br><font color=""red"">很遗憾，该服务器不支持 " & strclass & " 组件！</font>"
  Else
    If VerObj = "" Or IsNull(VerObj) Then
      Verobj1 = "无法取得该组件版本"
    Else
      Verobj1 = "该组件版本是：" & VerObj
    End If
    Response.Write "<br><font class=fonts>恭喜！该服务器支持 " & strclass & " 组件。" & verobj1 & "</font>"
  End If
  Response.Write "<br>"
End If
%>
        <FORM action="" method="post" id="form1" name="form1">
在右侧的文本框中输入你要检测的组件的ProgId或ClassId：
      <input class="input2 size1" type="text" value="" name="classname" size="40" />
      <input type="submit" value=" 确 定 " class="button8" id="submit1" name="submit1" />
    <form action="" method="post" id="form1" name="form1">
    </form>
    </td>
  </tr>
  <tr class="tr1">
    <td colspan=4 class="td_title">■ 当前程序根目录信息
      <%
dPath = server.MapPath("/")
Set dDir = fsoObj.GetFolder(dPath)
Set dDrive = fsoObj.GetDrive(dDir.Drive)
%>
      文件夹: <%=dPath%></td>
  </tr>
  <tr height="23" align="center">
    <td >已用空间/可用空间</td>
    <td >程序所有文件夹及文件数</td>
    <td >创建时间</td>
    <td >修改时间</td>
  </tr>
  <tr height="20" align="center">
    <td ><%=cSize(dDir.Size)%>/<%=cSize(dDrive.AvailableSpace)%></td>
    <td ><%=BianLi (dPath,0,0)%></td>
    <td ><%=dDir.DateCreated%>
</td>
    <td ><%=dDir.DateLastModified%></td>
  </tr>
  <tr class="tr1">
    <td colspan=6 class="td_title">■ 磁盘文件操作速度测试（1000毫秒内都算可以）</td>
  </tr>
  <tr height="20" align=center>
    <td colspan="4" ><%
Response.Write "正在重复创建、写入和删除文本文件50次==>"
Dim thetime3, tempfile, iserr
iserr = false
t1 = Timer
tempfile = server.MapPath("./") & "\aspchecktest.txt"
For i = 1 To 50
  Err.Clear
  Set tempfileOBJ = FsoObj.CreateTextFile(tempfile, true)
  If Err <> 0 Then
    Response.Write "创建文件错误！"
    iserr = true
    Err.Clear
    Exit For
  End If
  tempfileOBJ.WriteLine "Only for test. Ajiang ASPcheck"
  If Err <> 0 Then
    Response.Write "写入文件错误！"
    iserr = true
    Err.Clear
    Exit For
  End If
  tempfileOBJ.Close
  Set tempfileOBJ = FsoObj.GetFile(tempfile)
  tempfileOBJ.Delete
  If Err <> 0 Then
    Response.Write "删除文件错误！"
    iserr = true
    Err.Clear
    Exit For
  End If
  Set tempfileOBJ = Nothing
Next
t2 = Timer
If iserr <> true Then
  thetime3 = CStr(Int(( (t2 - t1) * 10000 ) + 0.5) / 10)
  Response.Write " 已完成！==> <font color=red>" & thetime3 & "毫秒</font>。"
%></td>
  </tr>
  <%
End If
Set fsoobj = Nothing
End If
%>
  <tr class="tr1">
    <td colspan="4" class="td_title">■ ASP脚本解释和运算速度测试（1000毫秒内都算可以）</td>
  </tr>
  <tr height="20" align=center>
    <td colspan="4" ><%
'因为只进行50万次计算，所以去掉了是否检测的选项而直接检测
Response.Write "整数运算测试，正在进行50万次加法运算==>"
Dim t1, t2, lsabc, thetime, thetime2
t1 = Timer
For i = 1 To 500000
  lsabc = 1 + 1
Next
t2 = Timer
thetime = CStr(Int(( (t2 - t1) * 10000 ) + 0.5) / 10)
Response.Write " 已完成！ ==> <font color=red>" & thetime & "毫秒</font>。<br>"
Response.Write "浮点运算测试，正在进行20万次开方运算==>"
t1 = Timer
For i = 1 To 200000
  lsabc = 2^0.5
Next
t2 = Timer
thetime2 = CStr(Int(( (t2 - t1) * 10000 ) + 0.5) / 10)
Response.Write " 已完成！ ==> <font color=red>" & thetime2 & "毫秒</font>。<br>"
%></td>
  </tr>
  <tr>
    <td align=center style="LINE-HEIGHT: 150%" colspan="4"> 页面装载（非脚本执行时间，此值与服务器带宽，客户端带宽及其电脑性能有关，与服务器无关）：
      <script language="javascript">
d=new Date();endTime=d.getTime();document.write((endTime-startTime)/1000);
  </script>
      秒 </td>
  </tr>
  <tr>
    <td align=center style="LINE-HEIGHT: 150%" colspan="4">  </td>
  </tr>
</table>
<script type="text/javascript" >
showtable("mytable")
</script>
</body>
</html>
<%
Function cdrivetype(tnum)
  Select Case tnum
    Case 0
       cdrivetype = "未知"
    Case 1
       cdrivetype = "可移动磁盘"
    Case 2
       cdrivetype = "本地硬盘"
    Case 3
       cdrivetype = "网络磁盘"
    Case 4
       cdrivetype = "CD-ROM"
    Case 5
       cdrivetype = "RAM 磁盘"
  End Select
End Function

Function cIsReady(trd)
  Select Case trd
    Case true
       cIsReady = "<font class=fonts><b>√</b></font>"
    Case false
       cIsReady = "<font color='red'><b>×</b></font>"
  End Select
End Function

Function cSize(tSize)
  If tSize>= 1073741824 Then
    cSize = Int((tSize / 1073741824) * 1000) / 1000 & " GB"
  ElseIf tSize>= 1048576 Then
    cSize = Int((tSize / 1048576) * 1000) / 1000 & " MB"
  ElseIf tSize>= 1024 Then
    cSize = Int((tSize / 1024) * 1000) / 1000 & " KB"
  Else
    cSize = tSize & "B"
  End If
End Function
Set ctr_admin = Nothing
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