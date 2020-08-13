<%
Class tr_ulev
  Public id
  Public addtime
  Public levname
  Public userjf

  Public Sub Add()
    Call getform()
    addtime = FormatDate(Now(), 1)
    sql = "select * from tr_ulev where 1=2 "
    rs.Open sql, conn, 1, 3
    rs.addnew()
    Call toupdate(rs)
    Rs("addtime") = addtime
    rs.update
    rs.Close
  End Sub

  Public Sub edit()
    Call getform()
    sql = "select * from tr_ulev where id="&id&""
    rs.Open sql, conn, 1, 3
    Call toupdate(rs)
    rs.update
    rs.Close
  End Sub

  Public Sub del()
    id = Trim(Request.Form("id"))
    sql = "delete from tr_ulev where id="&id&""
    conn.Execute(sql)
  End Sub

  Public Function getform()
    id = Trim(Request.Form("id"))
    If id = "" Then id = GetMaxValue("tr_ulev", "id") + 1
    userjf = Trim(Request.Form("userjf"))
    levname = Trim(Request.Form("levname"))
  End Function

  Private Sub toupdate(ByRef Rs)
    Rs("id") = id
    Rs("userjf") = userjf
    Rs("levname") = levname
  End Sub

  Function create_select(ids)
    tmpstr = ""
    cosql = " select id,userjf,levname from tr_ulev order by userjf asc , id asc "
    rs2.Open cosql, conn, 0, 1
    Do While Not rs2.EOF
      id = rs2("id")
      userjf = rs2("userjf")
      levname = rs2("levname")
      tmplist = tmplist&"<option value="""&id&""" <%if ctr_article.limittype="&id&" then "&Chr(37)&"> selected=""selected"" <% end if "&Chr(37)&">>"&levname&vbCrLf
      If tmpid = "" Then
        tmpid = id
      Else
        tmpid = tmpid&","&id
      End If
      If tmpuserjf = "" Then
        tmpuserjf = userjf
      Else
        tmpuserjf = tmpuserjf&","&userjf
      End If
      If tmplevname = "" Then
        tmplevname = levname
      Else
        tmplevname = tmplevname&","&levname
      End If
      rs2.movenext
    Loop
    rs2.Close
    create_select = tmplist
    create_ulevstr = "<%"&vbCrLf&"idstr="","&tmpid&","""&vbCrLf&"userjfstr="""&tmpuserjf&""""&vbCrLf&"levnamestr="""&tmplevname&""""&vbCrLf&Chr(37)&">"
    Call createfile(create_select, funpath&"crinc/ulev_select.asp", 2, "gb2312")
    Call createfile(create_ulevstr, funpath&"crinc/ulevstr.asp", 2, "gb2312")
  End Function

End Class
%>
