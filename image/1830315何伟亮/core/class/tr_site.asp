<%
Class tr_site
  Public Sub edit()
    sql = "select * from tr_system"
    rs.Open sql, conn, 1, 3
    d = rs.fields.Count -1
    For ns = 0 To d
      field_value = request.Form(rs.fields(ns).Name)
      If InStr(rs.fields(ns).Name, "open") = 1 And field_value = "" Then field_value = 0
      If field_value<>"" Then
        rs(ns) = CHTMLEncode(field_value)
      End If
    Next
    rs("forbidusername") = ","&trhead_foot(request.Form("forbidusername"),",")&","
    rs("statisticscode") = ddunok020104(request.Form("statisticscode"))
    rs.update
    rs.Close
    application(siternd&"55trsitetitle") = ""
  End Sub
End Class
%>
