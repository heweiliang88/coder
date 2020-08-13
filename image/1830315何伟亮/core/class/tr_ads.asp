<%
Class tr_ads
  Public id
  Public orderid
  Public addtime
  Public types
  Public content
  Public clicks
  Public callid
  Public title
  Public ispass
  Public endtime

  Public Sub Add()
    Call getform()
    sql = "select * from tr_ads where 1=2 "
    rs.Open sql, conn, 1, 3
    rs.addnew()
    Call toupdate(rs)
    rs.update
    rs.Close
  End Sub

  Public Sub edit()
    Call getform()
    sql = "select * from tr_ads where id="&id&""
    rs.Open sql, conn, 1, 3
    Call toupdate(rs)
    rs.update
    rs.Close
  End Sub

  Public Sub del()
    id = Trim(Request.Form("id"))
    sql = "delete from tr_ads where id="&id&""
    conn.Execute(sql)
  End Sub

  Private Function getform()
    id = Trim(Request.Form("id"))
    If id = "" Then id = GetMaxValue("tr_ads", "id") + 1
    orderid = Trim(Request.Form("orderid"))
    If orderid = "" Then orderid = GetMaxValue("tr_ads", "orderid") + 1
    addtime = FormatDate(Now(), 1)
    types = Trim(Request.Form("types"))
    If types = "" Then types = 0
    content = Trim(Request.Form("content"))
	content=ddunok020104(content)
    clicks = Trim(Request.Form("clicks"))
    If clicks = "" Then clicks = 0
    callid = id
    title = Trim(Request.Form("title"))
    ispass = Trim(Request.Form("ispass"))
    If ispass = "" Then ispass = 0
    endtime = FormatDate(Trim(Request.Form("endtime")), 1)
  End Function

  Private Sub toupdate(ByRef Rs)
    Rs("id") = id
    Rs("orderid") = orderid
    Rs("addtime") = addtime
    Rs("types") = types
    Rs("content") = content
    Rs("clicks") = clicks
    Rs("callid") = callid
    Rs("title") = title
    Rs("ispass") = ispass
    If endtime = "" Then endtime = null
    Rs("endtime") = endtime
  End Sub

  Public Sub createads()
    sql = "select * from tr_ads where ispass=1 order by callid desc "
    rs.Open sql, conn, 0, 1
    tmpads = "function _55tr_com(callid){"& vbCrLf
    tmpads = tmpads&"switch(callid){"& vbCrLf
    Do While Not rs.EOF
      tmpads = tmpads& "//"&rs("title")&"" & vbCrLf
      tmpads = tmpads&"case ""tr"&rs("callid")&""":" & vbCrLf
      tmpads = tmpads& ""&replace(tojs(rs("content")),"trgoddns","<%=jsl"&chr(37)&">")&"" & vbCrLf
      tmpads = tmpads& "break;" & vbCrLf
      rs.movenext
    Loop
    rs.Close
    tmpads = tmpads&"}"& vbCrLf
    tmpads = tmpads&"}"& vbCrLf
    Call createfile(tmpads, "../crinc/tradsjs.asp", 2, "gb2312")
  End Sub

End Class
%>
