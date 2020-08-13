<%
Class tr_column
  Public id
  Public colname
  Public keywords
  Public descriptions
  Public intro
  Public types
  Public addtime
  Public orderid
  Public queueid
  Public columnid
  Public jumpurl
  Public isnav
  Public isindex
  Public islist
  Public isshow
  Public isopen
  Public tpagesize
  Public isimgtext

  Public Sub Add()
    Call getform()
    addtime = FormatDate(Now(), 1)
    queueid = queueid&id&","
    sql = "select * from tr_column where 1=2 "
    rs.Open sql, conn, 1, 3
    rs.addnew()
    Call toupdate(rs)
    Rs("addtime") = addtime
    rs.update
    rs.Close
  End Sub
	
  Public Sub edit()
    Call getform()
    sql = "select * from tr_column where id="&id&""
    rs.Open sql, conn, 1, 3
    Call toupdate(rs)
    rs.update
    rs.Close
  End Sub

  Public Sub del()
    id = Trim(Request.Form("id"))
    sql = "select * from tr_article where queueid like '%,"&id&",%' "
    rs.Open sql, conn, 0, 1
    If Not rs.EOF Then Call contrl_message ("该栏目下存在文章，请先删除文章或清空文章回收站或转移文章后操作。", "", "", "", "")
    rs.Close
    sql = "select * from tr_column where queueid like '%,"&id&",%' and id<>"&id&" "
    rs.Open sql, conn, 0, 1
    If Not rs.EOF Then Call contrl_message ("该栏目下存在下级栏目，请先删除或转移栏目后操作。", "", "", "", "")
    rs.Close
    sql = "delete from tr_column where id="&id&""
    conn.Execute(sql)
  End Sub
	
  Public Function getform()
    id = Trim(Request.Form("id"))
    If id = "" Then id = GetMaxValue("tr_column", "id") + 1
    colname = Trim(Request.Form("colname"))
    keywords = Trim(Request.Form("keywords"))
    descriptions = Trim(Request.Form("descriptions"))
    intro = Trim(Request.Form("intro"))
    types = Trim(Request.Form("types"))
    orderid = Trim(Request.Form("orderid"))
    If orderid = "" Then orderid = GetMaxValue("tr_column", "orderid") + 1
    queueid = Trim(Request.Form("queueid"))
    columnid = Trim(Request.Form("columnid"))
    If columnid = "" Then columnid = 0
    jumpurl = Trim(Request.Form("jumpurl"))
    isnav = Trim(Request.Form("isnav"))
    If isnav = "" Then isnav = 0
    isindex = Trim(Request.Form("isindex"))
    If isindex = "" Then isindex = 0
    islist = Trim(Request.Form("islist"))
    If islist = "" Then islist = 0
    isshow = Trim(Request.Form("isshow"))
    If isshow = "" Then isshow = 0
    isopen = Trim(Request.Form("isopen"))
    If isopen = "" Then isopen = 0
    tpagesize = Trim(Request.Form("tpagesize"))
    If tpagesize = "" Then tpagesize = 20
    isimgtext = Trim(Request.Form("isimgtext"))
  End Function
	
  Private Sub toupdate(ByRef Rs)
    Rs("id") = id
    Rs("colname") = colname
    Rs("keywords") = keywords
    Rs("descriptions") = descriptions
    Rs("intro") = intro
    Rs("types") = types
    Rs("orderid") = orderid
    Rs("queueid") = queueid
    Rs("columnid") = columnid
    Rs("jumpurl") = jumpurl
    Rs("isnav") = isnav
    Rs("isindex") = isindex
    Rs("islist") = islist
    Rs("isshow") = isshow
    Rs("isopen") = isopen
    Rs("tpagesize") = tpagesize
    Rs("isimgtext") = isimgtext
  End Sub

  Function create_select(ids)
    Set rscr = server.CreateObject("adodb.recordset")
    tmpstr = ""
    cosql = " select id,colname,queueid,isopen from tr_column where columnid="&ids&" order by id asc "
    rscr.Open cosql, conn, 0, 1
    Do While Not rscr.EOF
      id = rscr("id")
      colname = rscr("colname")
      isopen = rscr("isopen")
      queueid = trhead_foot(rscr("queueid"), ",")
      sp_queueid = Split(queueid, ",")
      ub_queueid = UBound(sp_queueid)
      tmpstr = ""
      For i = 0 To ub_queueid -2
        tmpstr = tmpstr&"│ "
      Next
      If isopen = 1 Then
        tmplist = tmplist&"<option value="""&id&""" <%if columnid="&id&" then "&Chr(37)&"> selected=""selected"" <% end if "&Chr(37)&">>"&tmpstr&"├"&colname&vbCrLf& create_select(id)
      Else
        tmplist = tmplist&vbCrLf& create_select(id)
      End If
      rscr.movenext
    Loop
    rscr.Close
    Set rscr = Nothing
    create_select = tmplist
  End Function
	
  Public Sub cr()
    Call createfile(create_select(0), funpath&"crinc/column_select.asp", 2, "gb2312")
  End Sub
End Class
%>
