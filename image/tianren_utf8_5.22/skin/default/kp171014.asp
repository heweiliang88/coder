<%

Private Sub tr_headnavnd(topn)
response.write "<div class=""container"">" & vbCrLf
response.write "<div class=""collapse navbar-collapse trcollapse"" id=""navbar-collapse"">" & vbCrLf
response.write "<div class=""trrow1200"">" & vbCrLf
response.write "<ul class=""nav navbar-nav trnavul"">" & vbCrLf
response.Write "<li class=""trnobg""><a href="""&funpath&""" target=""_self"" >首页</a></li>" & vbCrLf
  sql = "select top "&topn&" id,colname,jumpurl,queueid from tr_column where isnav=1 and isopen=1 and columnid=0 order by orderid asc , id desc "
  rs.Open sql, conn, 0, 1
  If Not rs.EOF Then
  spque=split(replace(queueid,",0,",","),",")
  ubque=ubound(spque)
  	stopac=false
    Do While Not rs.EOF
	if not stopac then
	for gh=ubque-1 to 1 step -1
	if instr(rs("queueid"),","&spque(gh)&",")>0 then
	trnavac="class=""trnavac"""
	stopac=true
	exit for
	else
	trnavac=""
	end if
	next
	end if
      tmpurl = outputurl(rs("id"), "", rs("jumpurl"), ishtml)
response.write "<li><a href="""&tmpurl&""" "&trnavac&">"&rs("colname")&"</a></li>" & vbCrLf
	trnavac=""
      rs.movenext
    Loop
  End If
  rs.Close
response.write "</ul>" & vbCrLf
response.write "</div>" & vbCrLf
response.write "</div>" & vbCrLf
response.write "</div>"
End Sub



Private Sub tr_rolltextnd(columnid, topn, strn, updown, showtype)

If columnid<>"" Then
getqueueid=getfieldvalue("tr_column", " queueid ", " and id="&columnid&" ", 1, "")
if getqueueid="" then getqueueid=",0,"
andsql =  " queueid like '" & getqueueid & "%' and "
else
andsql =  " queueid like ',0,%' and "
end if
  sql = "select top "&topn&" id,title,jumpurl,titlecolor from tr_article where "&andsql&" isroll=1 and ispass=1 and isdel=0 and isabout=0 "
  sql = sql & " order by id desc"
  rs.Open sql, conn, 0, 1
  response.Write "    <div id=""notice"" class=""trtopnotice trfl"">" & vbCrLf
  response.Write "      <ul id=""noticecontent"">" & vbCrLf
  If Not rs.EOF Then
    Do While Not rs.EOF
      jumpurl = outputurl("", rs("id"), rs("jumpurl"), ishtml)
      If rs("titlecolor")<>"" Then
        titlestr = "<font color="""&rs("titlecolor")&""">"&lefts(rs("title"), strn)&"</font>"
      Else
        titlestr = lefts(rs("title"), strn)
      End If
      response.Write "<li><a href="""&jumpurl&""" title="""&rs("title")&""">"&titlestr&"</a></li>" & vbCrLf
      rs.movenext
    Loop
  End If
  rs.Close
  response.Write "      </ul>" & vbCrLf
  response.Write "    </div>" & vbCrLf
  response.Write "    <script type=""text/javascript"">" & vbCrLf
  response.Write "new Marquee([""notice"",""noticecontent""],"&updown&",2,400,35,20,4000,3000,35,"&showtype&")" & vbCrLf
  response.Write "</script>" & vbCrLf
End Sub


Private Sub tr_focusnd(columnid, Width, height, topn, textn,trfocuspc,trfocusm)
If columnid<>"" Then
getqueueid=getfieldvalue("tr_column", " queueid ", " and id="&columnid&" ", 1, "")
if getqueueid="" then getqueueid=",0,"
andsql =  " queueid like '" & getqueueid & "%' and"
else
andsql =  " queueid like ',0,%' and"
end if
Sql = "select top "&topn&" Id,title,picfiles,jumpurl,titlecolor,mpicfiles from tr_article where "&andsql&" ispass=1 and isfocus=1 and isdel=0 and isabout=0 "
  Sql = Sql & " order by Id desc"
'  response.write sql
  Rs.Open Sql, Conn, 0, 1
  If rs.EOF Then
    rs.Close
    Exit Sub
  End If
  rscount = rs.getrows()
  rs.Close
response.write "<div class=""trslider"" id=""trslider""> <a href=""javascript:void(0)"" class=""unslider-arrow prev glyphicon glyphicon-menu-left"" target=""_self"" id=""trleftimg""></a>" & vbCrLf
response.write "<a href=""javascript:void(0)"" class=""unslider-arrow next glyphicon glyphicon-menu-right"" target=""_self"" id=""trrightimg""></a>" & vbCrLf
response.write "<ul class=""tabcon"">" & vbCrLf
  i = 0
  ubrscount = UBound(rscount, 2)
  For i = 0 To ubrscount
    id = rscount(0, i)
    titles = Lefts(rscount(1, i), textn)
    picfiles = rscount(2, i)
    jumpurl = outputurl("", id, rscount(3, i), ishtml)
	mpicfiles = rscount(5, i)
    titlecolor = rscount(4, i)
	    If titlecolor<>"" Then
      titlestr = "<font color="""&rscount(4, i)&""">"&titles&"</font>"
    Else
      titlestr = titles
    End If
response.write "<li><a href="""&jumpurl&""" target=""_blank""><img src="""&picfiles&""" alt="""&rscount(1, i)&""" class=""trfocus1200 "&trfocuspc&""" width="""& Width &""" height="""& height &""" ><img src="""&mpicfiles&""" alt="""&rscount(1, i)&""" class=""trfocus1199 "&trfocusm&""" width=""1000"" height=""300"">" & vbCrLf
response.write "<p>"&titlestr&"</p>" & vbCrLf
response.write "</a></li>" & vbCrLf
Next

response.write "</ul>" & vbCrLf
response.write "</div>" & vbCrLf
response.write "<script type=""text/javascript"">" & vbCrLf
response.write "var otrslider=document.getElementById('trslider');" & vbCrLf
response.write "var otrleftimg=document.getElementById('trleftimg');" & vbCrLf
response.write "var otrrightimg=document.getElementById('trrightimg');" & vbCrLf
response.write "otrslider.onmouseover=function(){" & vbCrLf
response.write "otrleftimg.style.display=""block"";" & vbCrLf
response.write "otrrightimg.style.display=""block"";" & vbCrLf
response.write "}" & vbCrLf
response.write "otrslider.onmouseout=function(){" & vbCrLf
response.write "otrleftimg.style.display=""none"";" & vbCrLf
response.write "otrrightimg.style.display=""none"";" & vbCrLf
response.write "}" & vbCrLf
response.write "$(function() {" & vbCrLf
response.write "	if(window.chrome){$('.trslider li').css('background-size', '100% 100%');}" & vbCrLf
response.write "	$('.trslider').unslider({" & vbCrLf
response.write "		speed: 500," & vbCrLf
response.write "		delay: 3000," & vbCrLf
response.write "		complete: function() {}," & vbCrLf
response.write "		keys: false," & vbCrLf
response.write "		dots: true," & vbCrLf
response.write "		fluid: true" & vbCrLf
response.write "	});" & vbCrLf
response.write "	var unslider = $('.trslider').unslider();" & vbCrLf
response.write "	$('.unslider-arrow').click(function(){" & vbCrLf
response.write "		var fn = this.className.split(' ')[1];" & vbCrLf
response.write "		unslider.data('unslider')[fn]();" & vbCrLf
response.write "	});" & vbCrLf
response.write "var slides = jQuery('.trslider')," & vbCrLf
response.write "i = 0;" & vbCrLf
response.write "slides.on('swipeleft', function(e) {" & vbCrLf
response.write "unslider.data('unslider')['next']();" & vbCrLf
response.write "}).on('swiperight', function(e) {" & vbCrLf
response.write "unslider.data('unslider')['prev']();" & vbCrLf
response.write "});" & vbCrLf
response.write "});" & vbCrLf
response.write "" & vbCrLf
response.write "</script>" & vbCrLf


End Sub

function tr_smalllistkp2nd(columnid,topn,strn,showclick, showtime, showcolumn, timetype, ordertype,data_type)

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
  response.Write "    <ul class=""tr_smalllistkpul2"">" & vbCrLf
rs.open sql,conn,0,1
  If Not rs.EOF Then
    trtail = ""
    nn = 0
    dd = 0
    Do While Not rs.EOF
'      If istophot = 1 Then
'        dd = dd + 1
'        cc = cc + 1
'        If dd>4 Then cc = 4
'        spanstr = "<span class=""hotspdw trbgdw"&cc&" "">"&dd&"</span>"
'      End If
      trclicks = ""
      trtime = ""
      trcolumn = ""
	  columnid=rs("columnid")
	  coljumpurl=outputurl(columnid, "","", ishtml)
	  
arcolname=getfieldvalue("tr_column", "colname", "and id="&columnid&"", 1, "")
If showcolumn Then trcolumn = "<a href="""&coljumpurl&""" >["&lefts(arcolname, 10)&"]</a> | "
      If showclick Then trclicks = rs("clicks")&"次 | "
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
  response.Write "      </ul>" & vbCrLf
end function


Private Sub tr_indextopnd(columnid, topn, tstrn, nstrn, isdes)
If columnid<>"" Then
getqueueid=getfieldvalue("tr_column", " queueid ", " and id="&columnid&" ", 1, "")
if getqueueid="" then getqueueid=",0,"
andsql =  " queueid like '" & getqueueid & "%' and"
else
andsql =  " queueid like ',0,%' and"
end if
  sql = "select top "&topn&" id,title,jumpurl,titlecolor,descriptions,content from tr_article where "&andsql&" ispass=1 and isdel=0 and isindextop=1 and isabout=0  "
  sql = sql & " order by id desc"
'  response.write sql
  rs.Open sql, conn, 0, 1
  If Not rs.EOF Then
    response.Write "    <div class=""trnewstop "">" & vbCrLf
    Do While Not rs.EOF
        desstr = lefts(RemoveHTML(rs("descriptions")), nstrn)
		if desstr="" then
        desstr = lefts(RemoveHTML(rs("content")), nstrn)
      End If
      If rs("titlecolor")<>"" Then
        titlestr = "<font color="""&rs("titlecolor")&""">"&lefts(rs("title"), tstrn)&"</font>"
      Else
        titlestr = lefts(rs("title"), tstrn)
      End If

      jumpurl = outputurl("", rs("id"), rs("jumpurl"), ishtml)
      response.Write "      <h3><a href="""&jumpurl&""" title="""&rs("title")&""">"&titlestr&"</a></h3>" & vbCrLf
      response.Write "      <p>"&desstr&"</p>" & vbCrLf
      rs.movenext
    Loop
    response.Write "    </div>" & vbCrLf
    response.Write "    <div class=""trdivline""></div>" & vbCrLf
  End If
  rs.Close
End Sub


function tr_smalllistkpnd(columnid,topn,strn,showclick, showtime, showcolumn, timetype, ordertype,data_type)
If columnid<>"" Then
getqueueid=getfieldvalue("tr_column", " queueid ", " and id="&columnid&" ", 1, "")
if getqueueid="" then getqueueid=",0,"
andsql =  " queueid like '" & getqueueid & "%' and"
else
andsql =  " queueid like ',0,%' and"
end if
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
  End Select
  response.Write "    <ul class=""tr_smalllistkpul1"">" & vbCrLf
'  response.write sql
rs.open sql,conn,0,1
  If Not rs.EOF Then
    trtail = ""
    nn = 0
    dd = 0
    Do While Not rs.EOF
'      If istophot = 1 Then
'        dd = dd + 1
'        cc = cc + 1
'        If dd>4 Then cc = 4
'        spanstr = "<span class=""hotspdw trbgdw"&cc&" "">"&dd&"</span>"
'      End If
      trclicks = ""
      trtime = ""
      trcolumn = ""
	  columnida=rs("columnid")
	  coljumpurl=outputurl(columnida, "","", ishtml)
	  
arcolname=getfieldvalue("tr_column", "colname", "and id="&columnida&"", 1, "")
If showcolumn Then trcolumn = "<a href="""&coljumpurl&""" >["&lefts(arcolname, 10)&"]</a> | "
      If showclick Then trclicks = rs("clicks")&"次 | "
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
  response.Write "      </ul>" & vbCrLf
end function



function tr_hotsmalllistnd(columnid,topn,strn,showclick, showtime, showcolumn, timetype, ulclass)
andsql=""
If columnid<>"" Then
getqueueid=getfieldvalue("tr_column", " queueid ", " and id="&columnid&" ", 1, "")
if getqueueid="" then getqueueid=",0,"
andsql =  " queueid like '" & getqueueid & "%' and"
else
andsql =  " queueid like ',0,%' and"
end if
sql="select top "&topn&" id,title,jumpurl,titlecolor,addtime,clicks,columnid,columnname from tr_article where "&andsql&" ispass=1 and isdel=0 and isabout=0 order by clicks desc "
  response.Write "    <ul class="""&ulclass&""">" & vbCrLf
'  response.write sql
rs.open sql,conn,0,1
  If Not rs.EOF Then
    trtail = ""
    nn = 0
    dd = 0
    Do While Not rs.EOF
'      If istophot = 1 Then
        dd = dd + 1
        cc = cc + 1
        If dd>4 Then cc = 4
        spanstr = "<span class=""hotsp trbg"&cc&" "">"&dd&"</span>"
'      End If
      trclicks = ""
      trtime = ""
      trcolumn = ""
If showcolumn Then
	  columnida=rs("columnid")
	  coljumpurl=outputurl(columnida, "","", ishtml)
	  
arcolname=getfieldvalue("tr_column", "colname", "and id="&columnida&"", 1, "")
trcolumn = "<a href="""&coljumpurl&""" >["&lefts(arcolname, 10)&"]</a> | "
end if
      If showclick Then trclicks = rs("clicks")&"次 | "
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
  response.Write "      </ul>" & vbCrLf
end function


Private Sub tr_rollimgkpnd(columnid, topn, tstrn, imgw, imgh, boxw, boxh)

If columnid<>"" Then
getqueueid=getfieldvalue("tr_column", " queueid ", " and id="&columnid&" ", 1, "")
if getqueueid="" then getqueueid=",0,"
andsql =  " queueid like '" & getqueueid & "%' and"
else
andsql =  " queueid like ',0,%' and"
end if


  sql = "select top "&topn&" id,title,jumpurl,titlecolor,picfiles from tr_article where "&andsql&" ispass=1 and isdel=0 and isrollimg=1 and isabout=0 "
  sql = sql & " order by id desc"
  rs.Open sql, conn, 0, 1
  response.Write "  <div id=""trrollimgnr"" class=""trrollimgnr"">" & vbCrLf
  response.Write "    <ul id=""trrollimgul"">" & vbCrLf
  If Not rs.EOF Then
    Do While Not rs.EOF
      If Len(rs("picfiles"))>8 Then
        jumpurl = outputurl("", rs("id"), rs("jumpurl"), ishtml)
        If rs("titlecolor")<>"" Then
          titlestr = "<font color="""&rs("titlecolor")&""">"&lefts(rs("title"), tstrn)&"</font>"
        Else
          titlestr = lefts(rs("title"), tstrn)
        End If
        response.Write "      <li><a href="""&jumpurl&""" title="""&rs("title")&"""><img src="""&rs("picfiles")&""" width="""&imgw&""" height="""&imgh&""" alt="""&rs("title")&"""/></a><span><a href="""&jumpurl&""">"&titlestr&"</a></span></li>" & vbCrLf
      End If
      rs.movenext
    Loop
  End If
  rs.Close
  response.Write "    </ul>" & vbCrLf
  response.Write "  </div>" & vbCrLf
  response.Write "  <script type=""text/javascript"">" & vbCrLf
  response.Write "new Marquee(" & vbCrLf
  response.Write "{" & vbCrLf
  response.Write "	MSClass	  : [""trrollimgnr"",""trrollimgul""]," & vbCrLf
  response.Write "	Direction : 2," & vbCrLf
  response.Write "	Step	  : 0.3," & vbCrLf
  response.Write "	Width	  : "&boxw&"," & vbCrLf
  response.Write "	Height	  : "&boxh&"," & vbCrLf
  response.Write "	Timer	  : 20," & vbCrLf
  response.Write "	DelayTime : 3000," & vbCrLf
  response.Write "	WaitTime  : 0," & vbCrLf
  response.Write "	ScrollStep: 164," & vbCrLf
  response.Write "	SwitchType: 0," & vbCrLf
  response.Write "	AutoStart : true" & vbCrLf
  response.Write "});" & vbCrLf
  response.Write "</script> " & vbCrLf
End Sub


Private Sub tr_imgtextnd(columnid, topn, imgw, imgh, tstrn, nstrn, isdes)

If columnid<>"" Then
getqueueid=getfieldvalue("tr_column", " queueid ", " and id="&columnid&" ", 1, "")
if getqueueid="" then getqueueid=",0,"
andsql =  " queueid like '" & getqueueid & "%' and"
else
andsql =  " queueid like ',0,%' and"
end if

  sql = "select top "&topn&" id,title,jumpurl,titlecolor,picfiles,descriptions,content from tr_article where "&andsql&" ispass=1 and isdel=0 and isimgtext=1 and isabout=0 "
  sql = sql & " order by id desc"
  rs.Open sql, conn, 0, 1
  If Not rs.EOF Then
    jumpurl = outputurl("", rs("id"), rs("jumpurl"), ishtml)
desstr = lefts(RemoveHTML(rs("descriptions")), nstrn)
if desstr="" then
desstr = lefts(RemoveHTML(left(rs("content"),500)), nstrn)
if len(desstr)<nstrn/2 then desstr = lefts(RemoveHTML(left(rs("content"),2500)), nstrn)
if len(desstr)<nstrn/2 then desstr = lefts(RemoveHTML(rs("content")), nstrn)
end if
    Do While Not rs.EOF
      If rs("titlecolor")<>"" Then
        titlestr = "<font color="""&rs("titlecolor")&""">"&lefts(rs("title"), tstrn)&"</font>"
      Else
        titlestr = lefts(rs("title"), tstrn)
      End If
response.write "<div class=""trimgtext clearfix""><a href="""&jumpurl&""" ><img src="""&rs("picfiles")&""" width="""&imgw&""" height="""&imgh&"""/></a>" & vbCrLf
response.write "<div class=""trtext1 trovh "">" & vbCrLf
response.write "<h3><a href="""&jumpurl&""" title="""&rs("title")&""">"&titlestr&"</a></h3>" & vbCrLf
response.write "<p>"&desstr&"</p>" & vbCrLf
response.write "</div>" & vbCrLf
response.write "</div>"
      rs.movenext
    Loop
  End If
  rs.Close
End Sub


function tr_bksmalllistnd(columnid,topn,strn,showclick, showtime, showcolumn, timetype, ordertype,istophot)
andsql=""
  Select Case ordertype
    Case "listtop"
      andsql = " and islisttop=1 "
    Case "commend"
      andsql = " and iscommend=1 "
  End Select
If columnid<>"" Then
getqueueid=getfieldvalue("tr_column", " queueid ", " and id="&columnid&" ", 1, "")
if getqueueid="" then getqueueid=",0,"
andsql =  ""&andsql&" and queueid like '" & getqueueid & "%' "
end if
  sql = "select top "&topn&" id,title,jumpurl,titlecolor,addtime,clicks,columnid,columnname from tr_article where ispass=1 and isdel=0 and isabout=0 "&andsql&" "
  Select Case ordertype
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
  response.Write "    <ul class=""tridxul"">" & vbCrLf
'  response.write sql
rs.open sql,conn,0,1
  If Not rs.EOF Then
    trtail = ""
    nn = 0
    dd = 0
    Do While Not rs.EOF
'      If istophot = 1 Then
'        dd = dd + 1
'        cc = cc + 1
'        If dd>4 Then cc = 4
'        spanstr = "<span class=""hotspdw trbgdw"&cc&" "">"&dd&"</span>"
'      End If
      trclicks = ""
      trtime = ""
      trcolumn = ""
If showcolumn Then
	  columnida=rs("columnid")
	  coljumpurl=outputurl(columnida, "","", ishtml)
arcolname=getfieldvalue("tr_column", "colname", "and id="&columnida&"", 1, "")
trcolumn = "<a href="""&coljumpurl&""" >["&lefts(arcolname, 10)&"]</a> | "
end if
      If showclick Then trclicks = rs("clicks")&"次 | "
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
  response.Write "      </ul>" & vbCrLf
end function



Private Sub tr_listtopnd(columnid, topn, strn, showclick, showtime, showcolumn, timetype, ordertype)
If columnid<>"" Then
getqueueid=getfieldvalue("tr_column", " queueid ", " and id="&columnid&" ", 1, "")
if getqueueid="" then getqueueid=",0,"
andsql =  " queueid like '" & getqueueid & "%' and"
else
andsql =  " queueid like ',0,%' and"
end if
  sql = "select "&topstr&" id,title,jumpurl,titlecolor,addtime,clicks,columnid,columnname from tr_article where "&andsql&" ispass=1 and isdel=0 and isabout=0 and islisttop=1  "
      sql = sql&" order by id desc "
  rs.Open sql, conn, 0, 1
  If Not rs.EOF Then
    response.Write "      <ul class=""trlistul"">"
    Do While Not rs.EOF
      jumpurl = outputurl("", rs("id"), rs("jumpurl"), ishtml)
      If rs("titlecolor")<>"" Then
        titlestr = "<font color="""&rs("titlecolor")&""">"&lefts(rs("title"), strn)&"</font>"
      Else
        titlestr = lefts(rs("title"), strn)
      End If
      response.Write "      <li><span>[置顶]</span><a href="""&jumpurl&""" title="""&rs("title")&""">"&titlestr&"</a></li>" & vbCrLf
      rs.movenext
    Loop
    response.Write "</ul>"
  End If
  rs.Close
End Sub


Private Sub tr_articlelstnd(columnidkeyword, topnpagesize, strn, showclick, showtime, showcolumn, timetype, ordertype, ulclass, pagetype, showorsearch,data_type )
  If pagetype = 1 Then
    page_size = topnpagesize
    pagei = 0
    n = (pageno -1) * page_size
    topstr = ""
  Else
    topstr = " top "&topnpagesize
  End If
  sql = "select "&topstr&" id,title,jumpurl,titlecolor,addtime,clicks,columnid,columnname from tr_article where ispass=1 and isdel=0 and isabout=0  "
  If showorsearch = 1 Then

If columnidkeyword<>"" Then
getqueueid=getfieldvalue("tr_column", " queueid ", " and id="&columnidkeyword&" ", 1, "")
if getqueueid="" then getqueueid=",0,"
sql = sql &  " and queueid like '" & getqueueid & "%' "
else
sql = sql & " and queueid like ',0,%' "
end if
  Else
    If columnidkeyword = "" Then
      sql = sql&" and 1=2 "
    Else
      sql = sql&" and title like '%"&columnidkeyword&"%' "
    End If
  End If
  Select Case ordertype
    Case "new"
      sql = sql&" order by id desc "
    Case "hot"
      sql = sql&" order by clicks desc , id desc "
'    Case "listtop"
'      sql = sql&" order by islisttop desc , addtime desc ,id desc "
'    Case "commend"
'      sql = sql&" order by iscommend desc , addtime desc ,id desc "
    Case "rnd"
	if data_type="access" then
      Randomize
      sql = sql&" order by Rnd(-(id+"&CInt(Int((30 + 1) * Rnd) )&")) , id desc "
	 elseif data_type="mssql" then
      sql = sql&" order by newid() "
	 end if
'      Randomize
'      sql = sql&" order by Rnd(-(id+"&CInt(Int((30 + 1) * Rnd) )&")) , id desc "
  End Select
'response.write sql
  response.Write "    <ul class="""&ulclass&""">" & vbCrLf
  If pagetype = 1 Then
    Call ctr_page.create_rs(sql, rs, conn, page_size, pageno)
  Else
    rs.Open sql, conn, 0, 1
  End If
  If Not rs.EOF Then
    trtail = ""
    nn = 0
    dd = 0
    Do While Not rs.EOF
'      If istophot = 1 Then
'        dd = dd + 1
'        cc = cc + 1
'        If dd>4 Then cc = 4
'        spanstr = "<span class=""hotsp trbg"&cc&" "">"&dd&"</span>"
'      End If
      If pagetype = 1 Then
        If pagei>= page_size Then Exit Do
        pagei = pagei + 1
        n = n + 1
      End If
      trclicks = ""
      trtime = ""
      trcolumn = ""
      If showcolumn Then trcolumn = lefts(rs("columnname"), 10)&" | "
      If showclick Then trclicks = rs("clicks")&"次 | "
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
      response.Write "      <li>"&spanstr&showtrtail&"<a href="""&jumpurl&""" title="""&RemoveHTML(rs("title"))&""">"&titlestr&"</a></li>" & vbCrLf
      rs.movenext
    Loop
  Else
    If pagetype = 1 Then
      response.Write "<div class=""trzwnr"">抱歉暂无内容</div>"
    End If
  End If
  response.Write "      </ul>" & vbCrLf
  rs.Close
  If pagetype = 1 Then
  if ishtml=1 then
    response.Write ctr_page.create_page("", ctr_page.page_count, pageno, "trpage")
elseif showorsearch=2 then
    response.Write ctr_page.create_page("", ctr_page.page_count, pageno, "trpage")
else
    response.Write ctr_page.html_page("", ctr_page.page_count, pageno, "trpage")
end if
  End If
End Sub


Private Sub tr_piclistnd(columnid, page_size, Width, height, strn, showtext, ulclass)
  j = 0
  pagei = 0
  'n=(pageno-1)*page_size
If columnid<>"" Then
getqueueid=getfieldvalue("tr_column", " queueid ", " and id="&columnid&" ", 1, "")
if getqueueid="" then getqueueid=",0,"
'if getqueueid=0 then getqueueid=",0,"
andsql =  " queueid like '" & getqueueid & "%' and"
else
andsql =  " queueid like ',0,%' and"
end if
  sql = "select id,title,jumpurl,picfiles,titlecolor from tr_article where "&andsql&" IsPass=1 and IsDel=0 "
  sql = sql& " order by id desc "
  Call ctr_page.create_rs(sql, rs, conn, page_size, pageno)
  response.Write "    <ul class="""&ulclass&""">" & vbCrLf
  If Not rs.EOF Then
    Do While Not Rs.EOF
      If pagei>= page_size Then Exit Do
      pagei = pagei + 1
      j = j + 1
      jumpurl = outputurl("", rs("id"), rs("jumpurl"), ishtml)
      If j >1 Then
        mar = " class=""trmar6 "" "
      Else
        mar = ""
      End If
      response.Write "<li"&mar&">" & vbCrLf
      response.Write "<a target=""_blank"" href=""" & jumpurl & """>" & vbCrLf
      response.Write "<img width=""" & Width & """ height=""" & height & """ border=""0"" alt=""" & Rs("Title") & """ src=""" & formaturl(Rs("picfiles")) & """/>" & vbCrLf
      If showtext = 1 Then
        If rs("titlecolor")<>"" Then
          titlestr = "<font color="""&rs("titlecolor")&""">"&lefts(rs("title"), strn)&"</font>"
        Else
          titlestr = lefts(rs("title"), strn)
        End If
        response.Write "<p>" & vbCrLf
        response.Write titlestr&vbCrLf
        response.Write "</p>" & vbCrLf
      End If
      response.Write "</a>" & vbCrLf
      response.Write "</li>" & vbCrLf
      If j>4 Then j = 0
      rs.movenext
    Loop
  Else
    response.Write "<div class=""trzwnr"">抱歉暂无内容</div>"
  End If
  response.Write "      </ul>" & vbCrLf
  rs.Close
  if ishtml=1 then
    response.Write ctr_page.create_page("", ctr_page.page_count, pageno, "trpage")
	else
    response.Write ctr_page.html_page("", ctr_page.page_count, pageno, "trpage")
end if
End Sub



function tr_smalllistrightnd(columnid,topn,strn,showclick, showtime, showcolumn, timetype, ordertype,data_type)

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
  sql = "select top "&topn&" id,title,jumpurl,titlecolor,addtime,clicks,columnid,columnname from tr_article where "&andsql&" ispass=1 and isdel=0 and isabout=0  "
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
'      Randomize
'      sql = sql&" order by Rnd(-(id+"&CInt(Int((30 + 1) * Rnd) )&")) , id desc "
    Case else
      sql = sql&" order by id desc "
  End Select
  response.Write "    <ul class=""trsmallblock1ul"">" & vbCrLf
'  response.write sql
rs.open sql,conn,0,1
  If Not rs.EOF Then
    trtail = ""
    nn = 0
    dd = 0
    Do While Not rs.EOF
'      If istophot = 1 Then
'        dd = dd + 1
'        cc = cc + 1
'        If dd>4 Then cc = 4
'        spanstr = "<span class=""hotspdw trbgdw"&cc&" "">"&dd&"</span>"
'      End If
      trclicks = ""
      trtime = ""
      trcolumn = ""
	  columnid=rs("columnid")
	  coljumpurl=outputurl(columnid, "","", ishtml)
	  
arcolname=getfieldvalue("tr_column", "colname", "and id="&columnid&"", 1, "")
If showcolumn Then trcolumn = "<a href="""&coljumpurl&""" >["&lefts(arcolname, 10)&"]</a> | "
      If showclick Then trclicks = rs("clicks")&"次 | "
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
  response.Write "      </ul>" & vbCrLf
end function


Private Sub tr_getlinknd(topn, imgw, imgh, strn)
  sql = "select top "&topn&" id,title,homepage,types,picfile from tr_link where ispass=1 order by orderid asc "
  rs.Open sql, conn, 0, 1
  response.Write "  <div class="" linktext clearfix"">" & vbCrLf
  response.write "<ul>" & vbCrLf
  haveimglink = 0
  If Not rs.EOF Then
    Do While Not rs.EOF
      If rs("types") = 1 Then
        response.Write "      <li><a href="""&rs("homepage")&""" title="""&rs("title")&""">"&lefts(rs("title"), strn)&"</a></li>" & vbCrLf
      ElseIf rs("types") = 2 Then
        imglinktmp = imglinktmp&" <li><a href="""&rs("homepage")&""" title="""&rs("title")&"""><img src="""&rs("picfile")&""" width="""&imgw&""" height="""&imgh&""" /></a></li>" & vbCrLf
        haveimglink = 1
      End If
      rs.movenext
    Loop
  End If
  rs.Close
  response.Write "    </ul>" & vbCrLf
  response.Write "  </div>" & vbCrLf
  If haveimglink = 1 Then
    response.Write "  <div class="" linkimg clearfix trovh"">" & vbCrLf
    response.Write "    <ul class="""">" & vbCrLf
    response.Write imglinktmp
    response.Write "    </ul>" & vbCrLf
    response.Write "  </div>" & vbCrLf
  End If
End Sub





%>