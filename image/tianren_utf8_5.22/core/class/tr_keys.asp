<%
Class tr_keys
  Public id
  Public addtime
  Public keys
  Public notes
  Public isdel

  Public Sub Add()
    Call getform()
    addtime = FormatDate(Now(), 1)
    sql = "select * from tr_keys where 1=2 "
    rs.Open sql, conn, 1, 3
    rs.addnew()
    Call toupdate(rs)
    Rs("addtime") = addtime
    rs.update
    rs.Close
  End Sub

  Public Sub edit()
    Call getform()
    sql = "select * from tr_keys where id="&id&""
    rs.Open sql, conn, 1, 3
    Call toupdate(rs)
    rs.update
    rs.Close
  End Sub

  Public Sub dodel()
    id = Trim(Request.Form("id"))
    sql = "delete from tr_keys where id="&id&""
    conn.Execute(sql)
  End Sub

  Public Function getform()
    id = Trim(Request.Form("id"))
    If id = "" Then id = GetMaxValue("tr_keys", "id") + 1
    keys = Trim(Request.Form("keys"))
    notes = Trim(Request.Form("notes"))
  End Function

  Private Sub toupdate(ByRef Rs)
    Rs("id") = id
    Rs("keys") = keys
    Rs("notes") = notes
  End Sub

End Class
%>
