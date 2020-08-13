<%
Class tr_link
  Public id
  Public orderid
  Public types
  Public title
  Public homepage
  Public intro
  Public picfile
  Public addtime
  Public ispass

  Public Sub Add()
    Call getform()
    sql = "select * from tr_link where 1=2 "
    rs.Open sql, conn, 1, 3
    rs.addnew()
    Call toupdate(rs)
    rs.update
    rs.Close
  End Sub

  Public Sub edit()
    Call getform()
    sql = "select * from tr_link where id="&id&""
    rs.Open sql, conn, 1, 3
    Call toupdate(rs)
    rs.update
    rs.Close
  End Sub

  Public Sub del()
    id = Trim(Request.Form("id"))
    sql = "delete from tr_link where id="&id&""
    conn.Execute(sql)
  End Sub

  Public Function getform()
    id = Trim(Request.Form("id"))
    If id = "" Then id = GetMaxValue("tr_link", "id") + 1
    orderid = Trim(Request.Form("orderid"))
    If orderid = "" Then orderid = GetMaxValue("tr_link", "orderid") + 1
    addtime = FormatDate(Now(), 1)
    types = Trim(Request.Form("types"))
    If types = "" Then types = 0
    ispass = Trim(Request.Form("ispass"))
    If ispass = "" Then ispass = 0
    title = Trim(Request.Form("title"))
    homepage = formaturl(Trim(Request.Form("homepage")))
    intro = Trim(Request.Form("intro"))
    picfile = Trim(Request.Form("picfile"))
  End Function

  Private Sub toupdate(ByRef Rs)
    Rs("id") = id
    Rs("orderid") = orderid
    Rs("types") = types
    Rs("title") = title
    Rs("homepage") = homepage
    Rs("intro") = intro
    Rs("picfile") = picfile
    Rs("addtime") = addtime
    Rs("ispass") = ispass
  End Sub

End Class
%>
