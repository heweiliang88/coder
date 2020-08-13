<%
Class tr_page

  Public page_count

  Public Sub create_rs(sql, rs, conn, page_size, pageno)
    pageno = clng(pageno)
    rs.Open sql, conn, 1, 1
    If Not rs.EOF And Not rs.bof Then
      rs.pagesize = page_size
      page_count = clng(rs.pagecount)
      If pageno>page_count Then pageno = page_count
      rs.absolutepage = pageno
    End If
  End Sub

  Function create_page(urlstr, page_count, pageno, pageclass)
    page_count = clng(page_count)
    pageno = clng(pageno)
    QUERY_STRING = Request.ServerVariables("QUERY_STRING")
    Url = Request.ServerVariables("Url")
    If QUERY_STRING<>"" Then
      QUERY_STRING = Replace(Replace(QUERY_STRING, "trcms=1&", ""), "&&", "&")
    End If
    If InStr(QUERY_STRING, "pageno=")>0 Then QUERY_STRING = Split(QUERY_STRING, "pageno=")(0)
    urlstr = Url&"?trcms=1&"&QUERY_STRING&"&pageno="
    urlstr = Replace(urlstr, "&&", "&")
    If IsNumeric(pageno) Then
      If pageno<1 Then pageno = 1
      If pageno>page_count Then pageno = page_count
    Else
      pageno = 1
    End If
    '列出11个页码，当前页左边5个，右边5
    If pageno>= 6 Then
      leftnum = pageno -6
      If pageno + 5>page_count Then leftnum = page_count -11
    ElseIf pageno<6 Then
      leftnum = 0
    End If
    If leftnum<0 Then leftnum = 0
    previous_page = pageno -1
    If previous_page<1 Then previous_page = 1
    next_page = pageno + 1
    If next_page>page_count Then next_page = page_count
    If page_count>1 Then
      tmppage = tmppage& "<ul class="""&pageclass&" "">"
      tmppage = tmppage& "<a href="""&urlstr&"1"" target=""_self"" >首页</a>"
      tmppage = tmppage& "<a href="""&urlstr&previous_page&""" target=""_self"" >前一页</a>"
      For i = 0 To 10
        leftnum = leftnum + 1
        If leftnum>page_count Then Exit For
        If leftnum = pageno Then
          theclass = "class=""now"""
          nowurl = "javascript:void(0)"
        Else
          theclass = ""
          nowurl = urlstr&leftnum
        End If
        tmppage = tmppage& "<a href="""&nowurl&""" "&theclass&" target=""_self"" >"&leftnum&"</a>"
      Next
      tmppage = tmppage& "<a href="""&urlstr&next_page&""" target=""_self"" >后一页</a>"
      tmppage = tmppage& "<a href="""&urlstr&page_count&""" target=""_self"" >尾页</a>"
      tmppage = tmppage& "</ul>" & vbCrLf
      tmppage = tmppage& "<!--勿删此注释调用页码用"&page_count&"-->"
    End If
    create_page = tmppage
  End Function

  Function create_page2(urlstr, page_count, pageno, pageclass, parentid)
    page_count = clng(page_count)
    pageno = clng(pageno)
    QUERY_STRING = Request.ServerVariables("QUERY_STRING")
    Url = Request.ServerVariables("Url")
    If QUERY_STRING<>"" Then
      QUERY_STRING = Replace(Replace(QUERY_STRING, "trcms=1&", ""), "&&", "&")
    End If
    If InStr(QUERY_STRING, "pageno=")>0 Then QUERY_STRING = Split(QUERY_STRING, "pageno=")(0)
    urlstr = Url&"?trcms=1&"&QUERY_STRING&"&pageno="
    urlstr = Replace(urlstr, "&&", "&")
    If IsNumeric(pageno) Then
      If pageno<1 Then pageno = 1
      If pageno>page_count Then pageno = page_count
    Else
      pageno = 1
    End If
    '列出11个页码，当前页左边5个，右边5
    If pageno>= 6 Then
      leftnum = pageno -6
      If pageno + 5>page_count Then leftnum = page_count -11
    ElseIf pageno<6 Then
      leftnum = 0
    End If
    If leftnum<0 Then leftnum = 0
    previous_page = pageno -1
    If previous_page<1 Then previous_page = 1
    next_page = pageno + 1
    If next_page>page_count Then next_page = page_count
    If page_count>1 Then
      tmppage = tmppage& "<ul class="""&pageclass&""">" & vbCrLf
      tmppage = tmppage& "<a href=""javascript:void(0);"" onclick=""reloadAbleJSFn('"&parentid&"','"&urlstr&"1')"">首页</a>"
      tmppage = tmppage& "<a href=""javascript:void(0);"" onclick=""reloadAbleJSFn('"&parentid&"','"&urlstr&previous_page&"')"">前一页</a>"
      For i = 0 To 10
        leftnum = leftnum + 1
        If leftnum>page_count Then Exit For
        If leftnum = pageno Then
          theclass = "class=""now"""
        Else
          theclass = ""
        End If
        tmppage = tmppage& "<a href=""javascript:void(0);"" "&theclass&" onclick=""reloadAbleJSFn('"&parentid&"','"&urlstr&leftnum&"')"">"&leftnum&"</a>"
      Next
      tmppage = tmppage& "<a href=""javascript:void(0);"" onclick=""reloadAbleJSFn('"&parentid&"','"&urlstr&next_page&"')"">后一页</a>"
      tmppage = tmppage& "<a href=""javascript:void(0);"" onclick=""reloadAbleJSFn('"&parentid&"','"&urlstr&page_count&"')"">尾页</a>"
      tmppage = tmppage& "</ul>" & vbCrLf
    End If
    create_page2 = tmppage
  End Function

function html_page(urlstr,page_count,pageno,pageclass)
page_count=clng(page_count)
pageno=clng(pageno)
QUERY_STRING = Request.ServerVariables("QUERY_STRING")
Url = Request.ServerVariables("Url")
if QUERY_STRING<>"" then
QUERY_STRING=replace(replace(QUERY_STRING,"trcms=1&",""),"&&","&")
end if
tid=request.querystring("id")
if isnumeric(pageno) then
if pageno<1 then pageno=1
if pageno>page_count then pageno=page_count
else
pageno=1
end if
'列出11个页码，当前页左边5个，右边5
if pageno>=6 then
leftnum=pageno-6
if pageno+5>page_count then leftnum=page_count-11
elseif pageno<6 then
leftnum=0
'result=pageno-4
end if
if leftnum<0 then leftnum=0
previous_page=pageno-1
if previous_page<1 then previous_page=1
next_page=pageno+1
if next_page>page_count then next_page=page_count
if page_count>1 then
tmppage=tmppage& "<ul class="""&pageclass&""">"
'tmppage=tmppage&"<a href="""&htmlurl(url,tid,1)&""" >首页</a>"
tmppage=tmppage& "<a href="""&htmlurl(url,tid,1)&""" target=""_self"" >首页</a>"
tmppage=tmppage& "<a href="""&htmlurl(url,tid,previous_page)&""" target=""_self"" >前一页</a>"
for i=0 to 10
leftnum=leftnum+1
if leftnum>page_count then exit for
if leftnum=pageno then
theclass="class=""now"""
nowurl="javascript:void(0)"
else
theclass=""
nowurl=htmlurl(url,tid,leftnum)
end if
tmppage=tmppage& "<a href="""&nowurl&""" "&theclass&" target=""_self"" >"&leftnum&"</a>"
next
tmppage=tmppage& "<a href="""&htmlurl(url,tid,next_page)&""" target=""_self"" >后一页</a>"
tmppage=tmppage& "<a href="""&htmlurl(url,tid,page_count)&""" target=""_self"" >尾页</a>"
tmppage=tmppage& "</ul>" & vbCrLf
tmppage=tmppage& "<!--勿删此注释调用页码用"&page_count&"-->"
end if
html_page=tmppage
end function


function htmlurl(url,tid,pn)
if instr(Url,"/show.asp")>0 then
if pn=1 then
urlstr=funpath&htmlpath&tid&".html"
else
urlstr=funpath&htmlpath&tid&"_"&pn&".html"
end if
elseif instr(Url,"/list.asp")>0 then
if pn=1 then
urlstr=funpath&htmlpath&"list_"&tid&".html"
else
urlstr=funpath&htmlpath&"list_"&tid&"_"&pn&".html"
end if
end if
htmlurl=urlstr
end function


End Class
%>
