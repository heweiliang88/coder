<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../inc/charsetasp.asp"-->
<!DOCTYPE html>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%
Const dbpath = "../"
Const funpath = "../"
%>
<!--#include file="admin_fun.asp"-->
<!--#include file="../core/class/tr_page.asp"-->
<%
Set ctr_page = New tr_page
If InStr(","&adminlimitidrange&",", ",")<1 Then contrl_message "您没有访问权限，请尝试其他版块。", "", "", "", ""
andpropertys = " and apptype=2 "
pageno = tr_killstr(request.querystring("pageno"), 1, 8, 1, "", "", "", "", 1)
  searchaction = request.querystring("searchaction")
  If searchaction = "search" Then
    words = request.querystring("words")
    types = request.querystring("types")
    If types<>"" And words<>"" Then str2 = " and "&types&" like '%"&words&"%' "
    Str = str1&str2
  End If
  
Public Function GetholeUrl()
  GetholeUrl =  Request.ServerVariables("SERVER_NAME")
  If Request.ServerVariables("SERVER_PORT") <> 80 Then GetholeUrl = GetholeUrl &":" & Request.ServerVariables("SERVER_PORT")
  GetholeUrl = lcase(GetholeUrl & Request.ServerVariables("URL"))
GetholeUrl=replace(GetholeUrl,"http://","")
End Function 
  
%>
<title>模板列表<%=application(siternd & "55trsitename")%></title>
<link rel="stylesheet" type="text/css" href="./skin/default/style.css">
<base target="_self">
<script type="text/javascript" src="./js/js.js"></script>
<script type="text/javascript" src="./js/hover.js"></script>
<link rel="shortcut icon" type="image/x-icon" href="skin/default/img/favicon.ico">
<link rel="bookmark" type="image/x-icon" href="skin/default/img/favicon.ico">

<style>
a.bluefont1{background:#09f;color:#fff;padding:3px;}
a.redfont1{background:#f00;color:#fff;padding:3px;}
</style>
</head><body style="background:#ffffff; " class="w1">
<table width="100%" align="center" border="0" cellpadding="3" cellspacing="1" id="mytable2" class="size3" >
  <tr class="tr1" >
    <td class="td_title" colspan="10"><!--#include file="incgoback.asp"-->■
      <%if isdel=1 then%>
      回收站
      <%end if%>
      我的模板列表
      </td>
  </tr>
  <tr>
    <td colspan="10" valign="middle"><form target="_self" name="form1" id="form1" method="get" >
        <select name="types" id="types" class="size10">
          <option value="appname" <%if types="appname" then%>selected<%end if%> >模板名称
          <option value="appv" <%if types="appv" then%>selected<%end if%> >版本号
          <option value="appid" <%if types="appid" then%>selected<%end if%> >应用官网ID
        </select>
        <input type="text" name="words" id="words" value="<%=words%>" class="input2 size2">
        <input type="hidden" value="search" name="searchaction" id="searchaction">
        <input type="submit" value=" 搜索 " name="submit1" id="submit1" class="button4">
      </form></td>
  </tr>
  <tr class="font1">
    <td width="" align="center" nowrap>ID</td>
    <td width="" align="center" nowrap>序号</td>
    <td width="" align="center" nowrap>操作</td>
    <td width="" align="center" nowrap>管理</td>
    <td width="" align="center" nowrap>删除模板</td>
    <td width="200" align="center">模板名称</td>
    <td width="" align="center" nowrap>官网详情</td>
    <td width="" align="center" nowrap>版本号</td>
    <td width="" align="center" nowrap>安装状态</td>
    <td width="" align="center" nowrap>安装时间</td>
  </tr>
    <%
page_size = 40
pagei = 0
n = (pageno -1) * page_size
sql = "select * from tr_myapp where 1=1 "&Str&andpropertys&" and id is not null order by id desc "
'response.write sql
'response.End()
Call ctr_page.create_rs(sql, rs, conn, page_size, pageno)
If Not rs.EOF And Not rs.bof Then
  Do While Not rs.EOF And pagei<page_size
    pagei = pagei + 1
    n = n + 1
id = Rs("id")
appname = Rs("appname")
appv = Rs("appv")
addtime = Rs("addtime")
apptype = Rs("apptype")
insurl = Rs("insurl")
insactive = Rs("insactive")
iscanunins = Rs("iscanunins")
isopen = isint(Rs("isopen"))
if isopen="" then isopen=0
iscanclose = Rs("iscanclose")
appinfo = Rs("appinfo")
appweb = Rs("appweb")
appid = Rs("appid")
managepage = trim(Rs("managepage"))

%>
    <tr class="" id="tr<%=Rs("Id")%>">
      <td width="" align="center" nowrap>
        ID:<%=id%></td>
      <td width="" align="center" nowrap><%=n%></td>
      <td width="" align="center" class="tdbt2" nowrap><%if isopen=0 then%><a href="opentem.asp?id=<%=id%>" target="_blank" onClick="return confirms()" class="bluefont1">启用</a><%else%>正在使用<%end if%></td>
      <td width="" align="center" class="tdbt" nowrap><%
	  if managepage<>"" then
	  spmanagepage=split(managepage,",")
	  ubmanagepage=ubound(spmanagepage)
	  for nsi=0 to ubmanagepage
	  spmnsi=split(spmanagepage(nsi),"|")
	  ubspmnsi=ubound(spmnsi)
	  if ubspmnsi>0 then
	  if len(spmnsi(1))>0 then
urlstr170503="<a href="""&spmnsi(0)&""" class=""bluefont1"">"&spmnsi(1)&"</a>&nbsp;"	
else
urlstr170503="<a href="""&spmnsi(0)&""" class=""bluefont1"">功能"&nsi+1&"</a>&nbsp;"	  
end if  
	  else
urlstr170503="<a href="""&spmnsi(0)&""" class=""bluefont1"">功能"&nsi+1&"</a>&nbsp;"	  
	  end if
response.write urlstr170503
	  next
	  else
	  %>无需控制<%end if%></td>
      <td width="" align="center" class="tdbt" nowrap><a href="deltem.asp?id=<%=id%>" target="ifr1" class="redfont1">删此模板</a></td>
      <td width="400" align="left" style="word-break:break-all;"><%=appname%></td>
      <td align="center" style="word-break:break-all;" nowrap><a href="http://www.55tr.com/gwappshow.asp?id=<%=appid%>&u=<%=GetholeUrl()%>" target="_blank" class="bluefont1">点击查看(获取注册码)</a></td>
      <td width="" align="center" nowrap><%=appv%></td>
      <td width="" align="center" nowrap><%if insactive=1 then%><font color="#0000FF">已安装</font><%else%><font color="#FF0000">未安装</font><%end if%></td>
      <td width="" align="center" nowrap><%=FormatDate(addtime, 11) %></td>
    </tr>
    <%
rs.movenext
Loop
%>
    <tr>
      <td width="" align="center" ></td>
      <td align="center">&nbsp;</td>
      <td colspan="10" align="left">&nbsp;</td>
    </tr>
  <tr>
    <%if ctr_page.page_count>1 then%>
    <td colspan="10" align="left"><%= ctr_page.create_page("",ctr_page.page_count,pageno,"trpage")%></td>
    <% end if%>
  </tr>
  <%
Else
%>
  <tr>
    <td colspan="10" align="center">暂无相关项目
      <%if propertys<>"isdel" then%>
      ，请先添加。
      <%end if%></td>
  </tr>
  <%end if%>
  <tr>
    <td colspan="10" align="left">小提示：<br>
<p style="color:#F00;">安装模板后，之前的插件都要重装。更换域名后要用新域名重新获取注册码</p>      
    模板卸载不会删除数据库中内容，请放心卸载<br>
    启用模板相当于重新安装一次模板（释放模板文件），所以不要关闭启用模板时弹出的安装窗口<br>
    可通过上面列表“官网详情”下面的“点击查看（获取注册码）”链接获取注册码
如果某个页面出现文字提示“天人系列管理系统提示您：未检索到应用的注册码..”字样。请在上面的列表中“官网详情”下面点击“点击查看（获取注册码）”链接都点击一遍，到官网获取一次注册码即可<br>
</td>
  </tr>
</table>
<script type="text/javascript" >
showtable("mytable2")
function changev(value){
document.getElementById("action").value	=value
}
</script> 
<!--#include file="../inc/ifr.asp"-->
</body>
</html>
