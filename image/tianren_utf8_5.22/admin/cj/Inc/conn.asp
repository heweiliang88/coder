<%
'on error resume next
Const dbpath = "../../"
Const funpath = "../../"
Const data_type = "access"'为access或mssql
Const back_cookies_time = "3600"'单位为秒，建议填写1200秒
Const siternd = "55tr"'同服务器下放置多套程序时需将此改为不同的字符

Private Sub contrl_message(Msg, gotoUrl, buttomid, formid, whojump)
  Response.Write("<script language=""javascript"" >")& vbCrLf
  If Msg<>"" Then
    Response.Write("alert(""" & Msg & "\n小提示：按空格或回车可快速关闭此提示框"");")& vbCrLf
  End If
  If buttomid<>"" Then
    response.Write "parent.document.getElementById('"&buttomid&"').disabled=false ;"& vbCrLf
  End If
  If formid<>"" Then
    response.Write "parent.document.getElementById('"&formid&"').reset() ;"& vbCrLf
  End If
  If gotoUrl = "" Then
  Else
    If InStr(gotourl, "http://")<>1 And InStr(gotourl, "location.href")<1 And InStr(gotourl, "./")<1 Then gotourl = "http://"&gotourl
    If InStr(gotourl, "location.href")<1 Then
      Response.Write(""&whojump&".location='" & gotoUrl & "';")& vbCrLf
    Else
      Response.Write(""&whojump&".location=" & gotoUrl & ";")& vbCrLf
    End If
  End If
  Response.Write("</script>")
  call CloseConn()
  Response.End()
End Sub

  '检测管理员是否登陆并输出权限及权限范围
  Function checkadminlogin()
   
  if instr(funpath,"../")>0 then
  iub=ubound(split(funpath,"../"))
  if iub>1 then
  for dz=1 to iub-1
  tmpgo=tmpgo&"../"
  next
  gopath=tmpgo
  else
  gopath="./"
  end if
  else
  gopath="./"
  end if
'  response.write iub
'  response.End()
    id=request.Cookies("55trcms")("id")
    If id = "" Then 
    session("tr_cms") = ""
	Call contrl_message("管理员未登录，请登陆后操作。", gopath, "", "", "top")'未从cookies中获取到id，故退出重新登陆。
    end if
	id = tr_killstr(Trim(request.Cookies("55trcms")("id")), 1, 16, 2, "", "", "", "", 2)
    sql = "select md5name,adminname,passwd,cookieskey,cookiestime,isfree,ctrllimit,limitidrange from tr_admin where id="&id&""
    rs.Open sql, conn, 1, 3
    If Not rs.EOF Then'如果找到此ID有结果，则
      cookieskey = rs("cookieskey")
      cookiestime = FormatDate(rs("cookiestime"), 1)
      isfree = rs("isfree")
      nowtime = FormatDate(Now(), 1)
      If isfree = 0 Then
        session("tr_cms") = ""
        Call contrl_message("该用户已冻结，请更换其他帐户登陆。", gopath, "", "", "top")'如果此ID用户锁定，则提示退出
        response.End()
      End If
      If Int(DateDiff("s", cookiestime, nowtime))>Int(back_cookies_time) Then
        session("tr_cms") = ""
        Call contrl_message("管理员登陆后"&back_cookies_time&"秒内未操作，请重新登陆。\n若想延长无操作时间请修改core/conn.asp中的back_cookies_time参数。", gopath, "", "", "top")'长时间未操作，需重新登陆
        response.End()
      Else
        reqmd5name = request.cookies(cookieskey)("md5name")
        reqpasswd = request.cookies(cookieskey)("passwd")
        If reqmd5name = rs("md5name") And reqpasswd = rs("passwd") Then
          rs("cookiestime") = nowtime
          rs.update
          checkadminlogin = rs("adminname")&"|"&rs("ctrllimit")&"|"&rs("limitidrange")&"|"&id&"|"&rs("cookieskey")
          'theadminname=rs("adminname")
        Else
          session("tr_cms") = ""
          Call contrl_message("管理员用户名或密码错误，请重新登陆。", gopath, "", "", "top")'"管理员用户名或密码错误，请重新登陆。
        End If
      End If
    Else
      checkadminlogin = 0'未找到sql的select结果，未登录
      session("tr_cms") = ""
	Call contrl_message("管理员未登录，请登陆后操作。", gopath, "", "", "top")'未从cookies中获取到id，故退出重新登陆。
    End If
    rs.Close
  End Function

Public Function FormatDate(DateAndTime, para)
  On Error Resume Next
  Dim y, m, d, h, mi, s, strDateTime
  FormatDate = DateAndTime
  If Not IsNumeric(para) Then Exit Function
  If Not IsDate(DateAndTime) Then Exit Function
  y = CStr(Year(DateAndTime))
  m = CStr(Month(DateAndTime))
  If Len(m) = 1 Then m = "0" & m
  d = CStr(Day(DateAndTime))
  If Len(d) = 1 Then d = "0" & d
  h = CStr(Hour(DateAndTime))
  If Len(h) = 1 Then h = "0" & h
  mi = CStr(Minute(DateAndTime))
  If Len(mi) = 1 Then mi = "0" & mi
  s = CStr(Second(DateAndTime))
  If Len(s) = 1 Then s = "0" & s
  Select Case para
    Case "1"
      strDateTime = y & "-" & m & "-" & d & " " & h & ":" & mi & ":" & s
    Case Else
      strDateTime = DateAndTime
  End Select
  FormatDate = strDateTime
End Function

Function tr_killstr(Str, p55traratype, cutstr, pa55trratype, gotourl, buttomid, formid, whojump, inull)
  onestr = "=|-|;|_|\|^|(|)|+|$|'|*|%|#|&| |>|<|/|?|eva|copy|format|and|exec|insert|select|delete|update|count|chr|mid|master|truncate|char|declare|like|in|between|union|alert|escape|id|iframe|onerror|img|script|where|drop|alter|sor|javascript|jscript|vbscript|expression|this|name|onclick|on|function|var|cookie|document|or|dir|cmd|js"
  onestr2 = "-|_|\|^||+|$|'|*|%|#|&|@|?|html |eva|copy|format|and|exec|insert|select|delete|update|count|chr|mid|master|truncate|char|declare|like|in|between|union|alert|escape|id|frame|onerror|script|where|drop|alter|sor|javascript|jscript|vbscript|expression|this|name|onclick|on|function|var|cookie|document|or|dir|cmd|about|js|get|htc"
  par55tratype = Trim(Left(LCase(Str), cutstr))
  If inull = 2 And (par55tratype = "" Or IsNull(par55tratype) ) Then contrl_message "参数["&par55tratype&"]不允许为空，请输入字符。", gotourl, buttomid, formid, whojump
  If p55traratype = 1 Then
    If Not IsNumeric(par55tratype) Or par55tratype = "" Or IsNull(par55tratype) Then
      Select Case pa55trratype
        Case 1
          par55tratype = 1
        Case 2
          contrl_message "参数"&par55tratype&"应为数字型，请使用数字型。", gotourl, buttomid, formid, whojump
          response.End()
        Case 3
          par55tratype = 0
      End Select
    End If
  ElseIf p55traratype = 2 Then
    If par55tratype = "" Or IsNull(par55tratype) Then
      par55tratype = ""
    Else
      arronestr = Split(onestr, "|")
      For hgi = 0 To UBound(arronestr)
        If InStr(par55tratype, arronestr(hgi))>0 Then
          Select Case pa55trratype
            Case 1
              par55tratype = ""
            Case 2
              contrl_message "您提交的字符含有[ "&arronestr(hgi)&" ]不合法字符请返回重新编辑。", gotourl, buttomid, formid, whojump
              response.End()
          End Select
          Exit For
        End If
      Next
    End If
  ElseIf p55traratype = 3 And par55tratype<>"" Then
    trstr = par55tratype
    trstr = Replace(trstr, "<", "&lt;")
    trstr = Replace(trstr, ">", "&gt;")
    trstr = Replace(trstr, Chr(13) & Chr(10), "<br>") 
    trstr = Replace(trstr, Chr(32), "&nbsp;") 
    trstr = Replace(trstr, Chr(9), "&nbsp;")
    trstr = Replace(trstr, Chr(39), "&#39;")
    trstr = Replace(trstr, Chr(34), "&quot;")
    trstr = Replace(trstr, Chr(92), "&#92;")
    par55tratype = trstr
  ElseIf p55traratype = 4 And par55tratype<>"" Then
    arronestr = Split(onestr2, "|")
    For hgi = 0 To UBound(arronestr)
      If InStr(par55tratype, arronestr(hgi))>0 Then
        par55tratype = Replace(par55tratype, arronestr(hgi), "["&arronestr(hgi)&"]")
      End If
    Next
  End If
  tr_killstr = par55tratype
End Function

Function gHTTPPage(url) 
'On Error Resume Next
dim http 
set http=Server.createobject("Microsoft.XMLHTTP") 
Http.open "GET",url,false 
Http.send() 
if Http.readystate<>4 then exit function  
gHTTPPage=BToBstr(Http.responseBody,"utf-8")
set http=nothing
If Err.number<>0 then 
Err.Clear
End If  
End Function

Function BToBstr(body,Cset)
dim objstream
set objstream = Server.CreateObject("adodb.stream")
objstream.Type = 1
objstream.Mode =3
objstream.Open
objstream.Write body
objstream.Position = 0
objstream.Type = 2
objstream.Charset = Cset
BToBstr = objstream.ReadText 
objstream.Close
set objstream = nothing
End Function

Function ReadFromTextFile (FileUrl,CharSet) 
    dim str 
    set stm=server.CreateObject("adodb.stream") 
     stm.Type=2 '以本模式读取 
     stm.mode=3  
     stm.charset=CharSet 
     stm.open 
     stm.loadfromfile server.MapPath(FileUrl) 
     str=stm.readtext 
     stm.Close 
    set stm=nothing
     ReadFromTextFile=str 
End Function 

function ReplaceExp(patrn,str,typ)'typ是1时获取全部集合也就是Matches(0),2时获取子集合也就是Match.SubMatches(0)
	dim regEx,Matches,mypattern
	if patrn="" then
		ReplaceExp=str
		Exit function
	else
		mypattern=patrn
	end if
	set regEx=new regExp
	regEx.pattern=mypattern
	regEx.IgnoreCase=true
	regEx.Global=true
 Set Matches = regEx.Execute(str)   
    For Each Match in Matches   
        RetStr = RetStr & Match.SubMatches(0) '输出提取出来的数字   
    Next   
	Set regEx=nothing
	if typ=1 then
	if Matches.count>0 then
	ReplaceExp=Matches(0)
	else
	ReplaceExp=str
'	response.write Matches.count
'	response.End()
	end if
	elseif typ=2 then
	ReplaceExp=RetStr
	end if
end function

Private Sub CloseConn()
  Conn.Close
  Set Conn = Nothing
  Set rs = Nothing
  Set rs1 = Nothing
  Set rs2 = Nothing
  Set ctr_page = Nothing
  Set ctr_user = Nothing
  Set ctr_admin = Nothing
  Set ctr_ads = Nothing
  Set ctr_article = Nothing
  Set ctr_comment = Nothing
  Set ctr_column = Nothing
  Set ctr_link = Nothing
  Set ctr_site = Nothing
  Set ctr_ulev = Nothing
  Set ctr_link = Nothing
End Sub

Private Sub OpenConntr()
  On Error Resume Next
  Conn.Open ConnStr
  If Err.Number<>0 Then
'    Response.Write("数据库连接出现错误")
'    Response.Write(Err.Description)
    Err.Clear
    Response.End()
  Else
    On Error GoTo 0
  End If
End Sub

Function GetLocationURL() 
Dim Url 
Dim ServerPort,ServerName,ScriptName,QueryString 
ServerName = Request.ServerVariables("SERVER_NAME") 
ServerPort = Request.ServerVariables("SERVER_PORT") 
ScriptName = Request.ServerVariables("SCRIPT_NAME") 
QueryString = Request.ServerVariables("QUERY_STRING") 
Url="http://"&ServerName 
If ServerPort <> "80" Then Url = Url & ":" & ServerPort 
Url=Url&ScriptName 
If QueryString <>"" Then Url=Url&"?"& QueryString 
GetLocationURL=Url 
End Function 

conncj= ReadFromTextFile ("../../core/Conn.asp","utf-8") 
str2=ReplaceExp("\bdburl.*=.*dbpath.*&.*""(.*)""",conncj,2)
dim dburl,ser_dburl,fs,conn,connstr,rs,rs1,rs2,havefile
'*******************数据库设置********************
Select Case data_type
  Case "access"
    dburl = dbpath&str2
    ser_dburl = server.MapPath(dburl)
    Set fs = Server.CreateObject("Scripting.FileSystemObject")
    Set conn = server.CreateObject("adox.catalog")
    If fs.FileExists(ser_dburl) Then '判断文件是否存在
      havefile = 1
    End If
    Set conn = Nothing
    connstr = "provider=microsoft.jet.oledb.4.0;data source="&ser_dburl
End Select
Set rs = server.CreateObject("adodb.recordset")
Set rs1 = server.CreateObject("adodb.recordset")
Set rs2 = server.CreateObject("adodb.recordset")
Set conn = server.CreateObject("adodb.connection")
Call OpenConntr()


If Application(siternd & "55trsitetitle") = "" Then
  Application.Lock()
  Sql = "select top 1 * from tr_system order by id desc "
  Rs.Open Sql, Conn, 1, 1
  d = rs.fields.Count -1
  For ns = 0 To d
    application(siternd&"55tr"& rs.fields(ns).Name) = rs(ns).Value
  Next
  Rs.Close
  Application.UnLock()
End If

%>

<%
checkadminloginvalue = checkadminlogin
spcheckadminlogin = Split(checkadminloginvalue, "|")
adminname = spcheckadminlogin(0)
adminctrllimit = spcheckadminlogin(1)
adminlimitidrange = spcheckadminlogin(2)
adminid = spcheckadminlogin(3)
admincookieskey = spcheckadminlogin(4)

domainurl=Request.ServerVariables("Http_Host")'获取一级域名
if instr(domainurl,".")>0 then
if isnumeric(replace(replace(domainurl,".",""),":","")) then
strurl=lcase(domainurl)
else
spstrurl=split(domainurl,".")
ubstrurl=ubound(spstrurl)
if ubstrurl>1 then
if instr("com,net,org,gov",spstrurl(ubstrurl-1))<1 then
tmpurl=spstrurl(ubstrurl-1)&"."&spstrurl(ubstrurl)
else
tmpurl=spstrurl(ubstrurl-2)&"."&spstrurl(ubstrurl-1)&"."&spstrurl(ubstrurl)
end if
strurl=tmpurl
end if
end if
else
strurl=lcase(domainurl)
end if
'response.write strurl
'response.End()
isright=true
if instr(strurl,"localhost")<1 and instr(strurl,"127.0")<1 then 

isright=false

if not isright then
leftstr= ReadFromTextFile ("../left.asp","utf-8") 
if instr(1,leftstr,"<li>源程序：<a target=""_blank"" href=""http://www.55tr.com"">55TR.COM</a></li>",1)>0 and instr(1,leftstr,"www.55tr.com",1)>0 then 
isright=true
neednext=true
'response.write 1
else
isright=false
neednext=false
'response.write 2
end if
leftstr=""
end if
end if
if not isright then
%>
<div style=" width:700px; line-height:20px; font-size:12px; margin:20px 0 0 20px;">
<p style="font-size:18px; text-align:center; color:#F90; ">温馨提示</p>
<p>此功能暂时不能使用，原因可能如下：<br>
1、后台文件夹admin中的left.asp，底部代码“源程序：55tr.com”可能被修改了，此代码不影响您的网站外观及功能，但是作为作者版权声明有着重要的作用，恳请保留恢复原貌。<br>
2、请将以下源代码粘贴回原位：&lt;!--此行代码切勿改动--&gt;&lt;li&gt;源程序：&lt;a target=&quot;_blank&quot; href=&quot;http://www.55tr.com&quot;&gt;55TR.COM&lt;/a&gt;&lt;/li&gt;&lt;!--此行代码切勿改动--&gt; <br>
2、您的支持是我们前行的动力，我们坚持免费使用，坚持开发更多免费功能，即使付费的模板、功能模块也是几元钱的价格，一餐午饭的价格让你使用各种功能。告别高昂的二次开发费用，告别华而不实的商业版程序，就在<a href="http://www.55tr.com" target="_blank">www.55tr.com</a>。<br>
</p>
</div>
<%
response.end()
end if
%>
<%
'err.Clear
Const DataBaseType=1 
Const SysVer=1
Dim SqlNowString,SiteSN,DataServer,DataUser,DataBaseName,DataBasePsw,CMSDataBase,CollcetConnStr
	If data_type="access" Then
		SqlNowString = "Now()"
	Else
		SqlNowString = "getdate()"
	End If
dim db,ErrMsg,ErrMsg_lx,Site

dim connstrItem
dim dbItem
dim connItem
dbItem="Data/##8r5a7&2&f55tr.asa"
Set connItem = Server.CreateObject("ADODB.Connection")
connstrItem="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(dbItem)
connItem.Open connstrItem
If Err Then
   err.Clear
   Set ConnItem = Nothing
'   Response.Write err.description
   Response.Write "采集数据库连接出错，请检查连接字串数据库文件在admin/cj/data文件夹。"
   Response.End
End If
Sub CloseConnItem()
   On Error Resume Next
   ConnItem.close
   Set ConnItem=nothing
End sub

Sub OpenConn() 
	On Error Resume Next
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.open ConnStr
	If Err.Number<>0 Then
		Err.Clear
		Set Conn = Nothing
		Response.Write "数据库连接出错，请不要改动cj/data/##8r5a7&2&f55tr.asp 这个数据库文件名。"
'		Response.Write err.number
		Response.End
		Else
			On Error Goto 0
	End If
End Sub

Sub CloseConn()
	Conn.close
	Set Conn=nothing
End sub

%>
