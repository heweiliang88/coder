<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../inc/charsetasp.asp"-->
<!DOCTYPE html>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%Server.ScriptTimeOut = 600
'版权声明：
'天人QQ技术网站系统，以下简称“本程序、程序”可免费用于个人非商业用途，其他用途需付费购买版权，请按照如下约定执行：
'··本程序可免费用于个人非商业用途，除个人非商业外的用途均需要付费购买版权，并在官网登记域名作为正规版权记录，并可获得电子版版权证书(不是绑定域名，而是在官网数据库中登记域名旨在法律维权时避开已购买版权的域名)，侵权使用必将承担法律风险并追究高额赔偿
'··购买版权将获得稳定健全的升级服务、免费安装插件、模板、配套工具免费使用、官方优惠优先获得等特权。并可用于企业、政府、校园、单位、团体、个人、机构、部门、机关、公司及其他商业用途
'··购买版权后，除内核源代码中版权声明，后台左下角的“源程序：55tr.com”字样（这两种版权字样都不会在前台显示，只作为程序版权归属）以外所有代码都可以修改及删除，包括前台所有页面的官网版权信息都可修改。
'··购买版权后可用本程序为别人或自己提供建设网站，修改源码，二次开发的服务，但禁止发布或传播本程序及本程序的任何形态！禁止出售本程序及本程序的任何形态。说白了就是不能将程序出售或传播，否则侵权。（例如：小李与客户约定使用本程序制作某中学的官网，并收取某中学的网站制作服务费用1万元，并购买版权，此行为不侵权。小赵与客户约定使用本程序制作某小学的官网，并收取某小学的网站制作服务费用1万元，未购买版权，此行为侵权(因为未购买版权)，请终止！小孙使用本程序建立了个人漫画网站供用户浏览并获取广告收益，未购买版权，此行为不侵权。小张使用本程序制作为企业网站进行出售，并购买版权，此行为侵权（因为传播出售了程序），请终止！小王又将本程序制作为小说网站源码在网络免费发布传播提供下载，并购买版权，此行为侵权（因为传播出售了程序），请终止！）
'··不购买版权的个人非商业用途依然享受升级、安装插件、模板的权利，可修改源代码，但不得修改源代码中“官网链接”、“版权声明”、“应用中心”相关的代码，前台页面底部的“Copyright 20XX 天人系列管理系统 版权所有，授权xxx使用 Powered by 55TR.COM”字样不得删除,否则请购买版权
'··不购买版权的除个人非商业外的用途请及时购买版权或终止使用本程序，否则侵权
'··计算机软件著作权登记证书登记号：2016SR204636，侵权必究，作者QQ81962480 欢迎洽谈基于此程序开发、修改业务。
'··版权声明会存在修改的偶然性，请以官网公布的对应程序版本版权声明为准！具体的版权声明、特权、及功能介绍，请浏览：http://www.55tr.com/gwappshow.asp?id=23

	if err.number<>0 then
response.write "<script type=""text/javascript"">" & vbCrLf
response.write "alert(""（aspweb、NETBOX、小旋风、Aws和本程序自带的服务器软件）等简易的asp服务器软件功能不完善，请使用iis或上传到服务器再安装应用（模板、插件、升级包）"");" & vbCrLf
response.write "</script> "
	response.End()
	err.clear
end if
Const dbpath = "../"
Const funpath = "../"
%>
<!--#include file="admin_fun.asp"-->
<style>
.trgogw{width:500px; height:50px; line-height:20px; color:#fff; font-size:14px; font-weight:500; text-align:center; margin:10px auto;padding:10px;}
.tralert{background:#F00;color:#fff;}
.tralert a{color:#fff;}
.trnotice{background:#09F;color:#fff;}
.trnotice a{color:#fff;}

</style>



</head>
<body style="background:#2D2D2D;color:#fff;">
<%
'response.Flush()
'response.end()
If InStr(","&adminlimitidrange&",", ",")<1 Then contrl_message "您没有访问权限，请尝试其他版块。", "", "", "", ""
If adminctrllimit = 1 Then Call contrl_message ("您只有查看权限，没有管理权限，请升级权限后继续操作。", "", "", "", "")
appid = Trim(request.querystring("id"))
t=trim(getfieldvalue("tr_system", "systemtype", "",2, ""))

sql = "select * from tr_myapp where appid="&appid&""
rs.Open sql, conn, 0, 1
If rs.EOF Then
ishadapp=false
else
ishadapp=true
end if
rs.close


if ishadapp=false then
%>
<div style="width:500px; height:300px; line-height:80px; color:#04A2FF; font-size:30px; font-weight:600; text-align:center; margin:100px auto 0 auto;"> 正在加载..<br>
（耗时1-10分钟，请勿刷新页面）<img src="skin/default/img/appload.gif" width="80"/> </div>

<div style="" class="trgogw trnotice">若10分钟后页面未跳转，则请<a href="http://www.55tr.com/gwappshow.asp?id=<%=appid%>" target="_blank" style="color:#fff; background:#09F;padding:3px;border-radius:3px; text-decoration:none;"> 直接点此下载应用文件 </a>，然后到您的“网站后台--应用中心--上传安装中进行安装。应用ID为：<%=appid%></div>
<%
response.Flush()

filestr = filepathHttpPage("http://www.55tr.com/outapp.asp?id="&appid&"&t="&t&"&u="&GetholeUrl&"&updatekey="&application(siternd & "55trupdatekey")&"")
'response.write filestr
If InStr(filestr, "|")>0 Then
  spfile = Split(filestr, "|")
  filepath =spfile(0)
  appname = spfile(1)
  appv = spfile(2)
  apptype = spfile(3)
Elseif instr(filestr,"55tr")>0 and InStr(filestr, "|")<1 then
response.write filestr
response.End()
else
response.write "<div style="""" class=""trgogw tralert"">暂时无法读取该应用相关数据，请稍候再试，或<a href=""http://www.55tr.com/gwappshow.asp?id="&appid&""">直接点此下载应用文件</a>，然后到您的“网站后台--应用中心--上传安装中进行安装。应用ID为："&appid&"”</div>"
Call contrl_message ("", "", "", "", "")

End If
if len(filepath)<24 then
response.write "<div style="""" class=""trgogw tralert"">暂时无法读取该应用相关数据，请稍候再试，或<a href=""http://www.55tr.com/gwappshow.asp?id="&appid&""">直接点此下载应用文件</a>，然后到您的“网站后台--应用中心--上传安装中进行安装。应用ID为："&appid&"”</div>"
  Call contrl_message ("", "", "", "", "")
end if


filesizenum=getRemoteFileSize(filepath)
If filesizenum>0 and filesizenum<30000000 Then '大于0小于30M
call CreateMultiFolder("./app")
  If InStr(filepath, "/")>0 Then
    If apptype = 1 Then
      Call getfiles(filepath, 1)
    ElseIf apptype = 2 Then
      Call getfiles(filepath, 2)
    Else
      Call getfiles(filepath, 1)
    End If
    spurl = Split(filepath, "/")

    filename = spurl(UBound(spurl))
    sql = "select * from tr_myapp where appid="&appid&""
    rs.Open sql, conn, 1, 3
    If rs.EOF Then
'	response.Flush()

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
    If apptype = 1 Then
	response.write "<meta http-equiv=""refresh"" content=""2; url='uncompress.asp?inspath="&filename&"&appid="&appid&"'"">"
    ElseIf apptype = 2 Then
	response.write "<meta http-equiv=""refresh"" content=""2; url='uncompress.asp?inspath="&filename&"&appid="&appid&"'"">"
    Else
	response.write "<meta http-equiv=""refresh"" content=""2; url='uncompress.asp?inspath="&filename&"&appid="&appid&"'"">"
    End If
	%>
<div style="width:500px; height:50px; line-height:30px; color:#04A2FF; font-size:20px; font-weight:500; text-align:center; margin:50px auto 0 auto;"> <a href="uncompress.asp?inspath=<%=filename%>&appid=<%=appid%>" style="color:#FFF">如果10秒钟后页面没有自动跳转，请点此继续</a> </div>
<%
  End If
else
response.write "<div style="""" class=""trgogw tralert"">官网不存在该应用。应用的官网ID为："&appid&"，请联系作者QQ81962480修正该错误。</div>"
Call contrl_message ("", "", "", "", "")

End If


Else

'ComeUrl = LCase(Trim(request.ServerVariables("HTTP_REFERER")))
response.write "<div style="""" class=""trgogw tralert"">该应用已安装，如需重新安装，请先到“您的网站后台--应用中心--（我的模板）或（我的插件）中卸载后再安装”</div>"
Call contrl_message ("", "", "", "", "")

end if
%>
<%

Function getRemoteFileSize(url)
on error resume next
  Dim xmlHTTP
  Set xmlHTTP = Server.CreateObject("MSX"&"ML2.XM"&"LHT"&"TP")
  xmlHTTP.Open "get", url, false
  xmlHTTP.setRequestHeader "range", "bytes=-1"
  xmlHTTP.send()
  uber=ubound(Split(xmlHTTP.GetResponseHeader("Con"&"tent-Ran"&"ge"), "/"))
  if uber<1 then
response.write "<div style="""" class=""trgogw tralert"">官网不存在该应用。应用的官网ID为："&appid&"，请联系作者QQ81962480修正该错误。</div>"
Call contrl_message ("", "", "", "", "")

  end if
'  Call contrl_message ("官网不存在该应用。应用的官网ID为："&appid&"，请联系作者QQ81962480修正该错误。", "", "", "", "")

  getRemoteFileSize = Split(xmlHTTP.GetResponseHeader("Con"&"tent-Ran"&"ge"), "/")(1)
  Set xmlHTTP = Nothing
if err.number<>0 then
Call contrl_message (err.description&"getRemoteFileSize", "", "", "", "")
err.clear
end if
End Function


Function getfiles(url, lx)
  If url = "" Then Exit Function
  Set xmlhttp = server.CreateObject("Mic"&"ro"&"soft.X"&"MLH"&"TTP")
  xmlhttp.Open "get", url, false
  xmlhttp.send
  img = xmlhttp._
  ResponseBody
  contenttype = LCase(xmlhttp.getResponseHeader("con"&"tent-typ"&"e"))'获取响应头类型
  Set xmlhttp = Nothing
  If lx = 1 Then
    savepath = "app/" '保存路径
  ElseIf lx = 2 Then
    savepath = "app/" '保存路径
  Else
    savepath = "app/" '保存路径
  End If
  If InStr(url, "/")>0 Then
    spurl = Split(url, "/")
  Else
    Exit Function
  End If
  filename = spurl(UBound(spurl))
  Set objAdostream = server.CreateObject("ADO"&"DB.Str"&"eam")
  objAdostream.Open()
  objAdostream.Type = 1
  objAdostream.Write(img)
  objAdostream.SaveToFile server.MapPath(savepath&filename), 2
  objAdostream.SetEOS
  Set objAdostream = Nothing
End Function

Function filepathHttpPage(HttpUrl)
'response.write HttpUrl
  Dim http
  Set http = server.CreateObject("MSX"&"ML2.Serv"&"erXM"&"LHTTP")
  Http.Open "GET", HttpUrl, False
  Http.send()
  If Http.readystate<>4 Then Exit Function
  filepathHttpPage = htmlBytesToBstr(Http._
  ResponseBody,"ut"&"f-8")
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

Public Function GetholeUrl()
  GetholeUrl =  Request.ServerVariables("SERVER_NAME")
  If Request.ServerVariables("SERVER_PORT") <> 80 Then GetholeUrl = GetholeUrl &":" & Request.ServerVariables("SERVER_PORT")
  GetholeUrl = lcase(GetholeUrl & Request.ServerVariables("URL"))
'  If Trim(Request.QueryString) <>"" Then GetholeUrl = GetholeUrl &"?" & Trim(Request.QueryString)
GetholeUrl=replace(GetholeUrl,"http://","")
End Function 
Function CreateMultiFolder(ByVal CFolder) 
    BlInfo = False 
    CreateFolder = CFolder 
    Set objFSO = Server.CreateObject("Scripting.FileSystemObject") 
    CreateFolder = Replace(CreateFolder, "", "/") 
    If Left(CreateFolder, 1) = "/" Then 
        CreateFolder = Right(CreateFolder, Len(CreateFolder) -1) 
    End If 
    If Right(CreateFolder, 1) = "/" Then 
        CreateFolder = Left(CreateFolder, Len(CreateFolder) -1) 
    End If 
    CreateFolderArray = Split(CreateFolder, "/") 
    For i = 0 To UBound(CreateFolderArray) 
        CreateFolderSub = "" 
        For ii = 0 To i 
            CreateFolderSub = CreateFolderSub & CreateFolderArray(ii) & "/" 
        Next 
        PhCreateFolderSub = Server.MapPath(CreateFolderSub) 
        If Not objFSO.FolderExists(PhCreateFolderSub) Then 
            objFSO.CreateFolder(PhCreateFolderSub) 
        End If 
    Next 
End Function 
%>
</body>
</html>
<%'勿删改此行%>


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
