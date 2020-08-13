<%
'*******************程序基础设置********************
Const data_type = "access"'为access或mssql
Const back_cookies_time = "1200"'单位为秒，建议填写1200秒
Const siternd = "55tr"'同服务器下放置多套程序时需将此改为不同的字符
ishtml = application(siternd & "55trurltype")'程序是否为静态版本
htmlpath = application(siternd & "55trhtmlurl")&"/"'静态页面存储文件夹
if not isnumeric(ishtml) then
ishtml=1
end if
mishtml = application(siternd & "55trmurltype")'手机版程序是否为静态版本
if not isnumeric(mishtml) then
mishtml=1
end if
mhtmlpath = application(siternd & "55trmhtmlurl")&"/"'手机版静态页面存储文件夹
'response.write ishtml
Const trversion = "55TR_AH_Acc_GV"'程序版本

dim dburl,ser_dburl,fs,conn,connstr,rs,rs1,rs2,havefile
'*******************数据库设置********************
Select Case data_type
  Case "access"
    dburl = dbpath&"db/%66$9&&##_55tr.asa"
    ser_dburl = server.MapPath("./"&dburl)
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
Call OpenConn()
Private Sub OpenConn()
  On Error Resume Next
  Conn.Open ConnStr
  If Err.Number<>0 Then
    Response.Write("数据库连接出现错误")
    Response.Write(Err.Description)
    Err.Clear
    Response.End()
  Else
    On Error GoTo 0
  End If
End Sub

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
