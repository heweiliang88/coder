<%

function tr_trindexnewjj(columnid,topn,strn,showclick, showtime, showcolumn, timetype, ordertype,data_type)

If columnid<>"" Then
getqueueid=getfieldvalue("tr_column", " queueid ", " and id="&columnid&" ", 1, "")
if getqueueid="" then getqueueid=",0,"
andsql =  " queueid like '" & getqueueid & "%' and"
else
andsql =  " queueid like ',0,%' and"
end if
  Select Case ordertype
    Case "listtop"
      andsql = andsql&" islisttop=1 and "
    Case "commend"
      andsql = andsql&" iscommend=1 and "
  End Select
  sql = "select top "&topn&" id,title,jumpurl,titlecolor,addtime,clicks,columnid,columnname from tr_article where "&andsql&" ispass=1 and isdel=0 and isabout=0 "
  Select Case ordertype
    Case "new"
      sql = sql&" order by id desc "
    Case "hot"
      sql = sql&" order by clicks desc , id desc "
    Case "rnd"
	if data_type="access" then
      Randomize
      sql = sql&" order by Rnd(-(id+"&CInt(Int((30 + 1) * Rnd) )&")) , id desc "
	elseif data_type="mssql" then
      sql = sql&" order by newid() "
	end if
    Case else
      sql = sql&" order by id desc "
  End Select
  response.Write "    <ul class=""trindexnewul"">" & vbCrLf
rs.open sql,conn,0,1
  If Not rs.EOF Then
    trtail = ""
    nn = 0
    dd = 0
    Do While Not rs.EOF
      If showclick Then
        dd = dd + 1
        cc = cc + 1
        If dd>4 Then cc = 4
        spanstr = "<span class=""hotspdw trbgdw"&cc&" "">"&dd&"</span>"
      End If
      trclicks = ""
      trtime = ""
      trcolumn = ""
	  columnid=rs("columnid")
	  coljumpurl=outputurl(columnid, "","", ishtml)
	  
arcolname=getfieldvalue("tr_column", "colname", "and id="&columnid&"", 1, "")
If showcolumn Then trcolumn = "<a href="""&coljumpurl&""" >["&lefts(arcolname, 10)&"]</a> | "
      If showclick Then trclicks = rs("clicks")&"´Î | "
      If showtime Then trtime = FormatDate(rs("addtime"), timetype)
      trtail = trcolumn&trclicks&trtime
      trtail = trhead_foot(trcolumn&trclicks&trtime, " | ")
      If trtail<>"" Then
        showtrtail = "<span class=""spw1"">"&trtail&"</span>"
      End If
      jumpurl = outputurl("", rs("id"), rs("jumpurl"), ishtml)
      If rs("titlecolor")<>"" Then
        titlestr = "<font color="""&rs("titlecolor")&""">"&lefts(rs("title"), strn)&"</font>"
      Else
        titlestr = lefts(rs("title"), strn)
      End If

      nn = nn + 1
      response.Write "<li>"&spanstr&showtrtail&"<a href="""&jumpurl&""" title="""&RemoveHTML(rs("title"))&""">"&titlestr&"</a></li>" & vbCrLf
      rs.movenext
    Loop
  End If
rs.close
  response.Write "</ul>" & vbCrLf
end function


function tr_trindexhotjj(columnid,topn,strn,showclick, showtime, showcolumn, timetype, ordertype,data_type)

If columnid<>"" Then
getqueueid=getfieldvalue("tr_column", " queueid ", " and id="&columnid&" ", 1, "")
if getqueueid="" then getqueueid=",0,"
andsql =  " queueid like '" & getqueueid & "%' and"
else
andsql =  " queueid like ',0,%' and"
end if
  Select Case ordertype
    Case "listtop"
      andsql = andsql&" islisttop=1 and "
    Case "commend"
      andsql = andsql&" iscommend=1 and "
  End Select
  sql = "select top "&topn&" id,title,jumpurl,titlecolor,addtime,clicks,columnid,columnname from tr_article where "&andsql&" ispass=1 and isdel=0 and isabout=0 "
  Select Case ordertype
    Case "new"
      sql = sql&" order by id desc "
    Case "hot"
      sql = sql&" order by clicks desc , id desc "
    Case "rnd"
	if data_type="access" then
      Randomize
      sql = sql&" order by Rnd(-(id+"&CInt(Int((30 + 1) * Rnd) )&")) , id desc "
	elseif data_type="mssql" then
      sql = sql&" order by newid() "
	end if
    Case else
      sql = sql&" order by id desc "
  End Select
  response.Write "    <ul class=""trindexhotul"">" & vbCrLf
rs.open sql,conn,0,1
  If Not rs.EOF Then
    trtail = ""
    nn = 0
    dd = 0
    Do While Not rs.EOF
      If showclick Then
        dd = dd + 1
        cc = cc + 1
        If dd>4 Then cc = 4
        spanstr = "<span class=""hotspdw trbgdw"&cc&" "">"&dd&"</span>"
      End If
      trclicks = ""
      trtime = ""
      trcolumn = ""
	  columnid=rs("columnid")
	  coljumpurl=outputurl(columnid, "","", ishtml)
	  
arcolname=getfieldvalue("tr_column", "colname", "and id="&columnid&"", 1, "")
If showcolumn Then trcolumn = "<a href="""&coljumpurl&""" >["&lefts(arcolname, 10)&"]</a> | "
      If showclick Then trclicks = rs("clicks")&"´Î | "
      If showtime Then trtime = FormatDate(rs("addtime"), timetype)
      trtail = trcolumn&trclicks&trtime
      trtail = trhead_foot(trcolumn&trclicks&trtime, " | ")
      If trtail<>"" Then
        showtrtail = "<span class=""spw1"">"&trtail&"</span>"
      End If
      jumpurl = outputurl("", rs("id"), rs("jumpurl"), ishtml)
      If rs("titlecolor")<>"" Then
        titlestr = "<font color="""&rs("titlecolor")&""">"&lefts(rs("title"), strn)&"</font>"
      Else
        titlestr = lefts(rs("title"), strn)
      End If

      nn = nn + 1
      response.Write "<li>"&spanstr&showtrtail&"<a href="""&jumpurl&""" title="""&RemoveHTML(rs("title"))&""">"&titlestr&"</a></li>" & vbCrLf
      rs.movenext
    Loop
  End If
rs.close
  response.Write "</ul>" & vbCrLf
end function


%>