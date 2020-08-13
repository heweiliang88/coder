<%
Class tr_comment
  Public id
  Public title
  Public content
  Public ownarticle
  Public owncolumn
  Public owncolumnqueue
  Public addtime
  Public types
  Public username
  Public mail
  Public homepage
  Public qq
  Public tel
  Public ip
  Public ispass
  Public isdel
  Public answer
  Public astime
  Public asuser
  Public istop
  Public iscommand
  Public Files
  Public picfiles
  Public headpic

    Public Sub admanswer()
      id = Trim(Request.Form("id"))
      answer = CHTMLEncode(Trim(Request.Form("answer"&id)))
      sql = "select * from tr_comment where id="&id&""
      rs.Open sql, conn, 1, 3
      Rs("answer") = answer
      Rs("astime") = Now()
      Rs("asuser") = adminname
      rs.update
      rs.Close
    End Sub

    Public Sub addcommentguest(how)'1为评论，2为留言
      vcode = tr_killstr(Trim(request.Form("vcode")), 1, 4, 2, "", "submit1", "", "", 2)
      If vcode<>session("getcode") Then '验证码正确
        Call contrl_message("验证码错误请重新输入", "", "submit1", "", "")'验证码错误
      End If
      id = GetMaxValue("tr_comment", "id") + 1
      content = tr_killstr(Trim(request.Form("content")), 3, 5000, 2, "", "submit1", "", "", 2)
      If how = 1 Then
        ownarticle = tr_killstr(Trim(request.Form("ownarticle")), 1, 16, 2, "", "submit1", "", "", 2)
        sql = "select * from tr_article where id="&ownarticle&" "
        rs.Open sql, conn, 0, 1
        If rs.EOF Or rs.bof Then
        Else
          owncolumn = Rs("columnid")
          owncolumnqueue = Rs("queueid")
        End If
        rs.Close
      End If
      types = how
      username = tr_killstr(Trim(request.Form("username")), 3, 30, 2, "", "submit1", "", "", 2)
      mail = tr_killstr(Trim(request.Form("mail")), 3, 30, 2, "", "submit1", "", "", 1)
      homepage = tr_killstr(Trim(request.Form("homepage")), 3, 60, 2, "", "submit1", "", "", 1)
      ip = getIP()
      If application(siternd & "55tropencheckcomment") = 1 Then
        ispass = 0
      Else
        ispass = 1
      End If
      isdel = 0
      istop = 0
      iscommand = 0
      addtime = FormatDate(Now(), 1)
      sql = "select * from tr_comment where 1=2 "
      rs.Open sql, conn, 1, 3
      rs.addnew()
      Rs("id") = id
      Rs("content") = content
      Rs("ownarticle") = ownarticle
      Rs("owncolumn") = owncolumn
      Rs("owncolumnqueue") = owncolumnqueue
      Rs("addtime") = Now()
      Rs("types") = types
      Rs("username") = username
      Rs("mail") = mail
      Rs("homepage") = homepage
      Rs("ip") = ip
      Rs("ispass") = ispass
      Rs("isdel") = isdel
      Rs("istop") = istop
      Rs("iscommand") = iscommand
      rs.update
      rs.Close
      If application(siternd & "55tropencheckcomment") = 1 Then
        Call contrl_message("留言/评论已提交，待管理员审核后显示。", "", "submit1", "form1", "")
      Else
        Call contrl_message("留言/评论发表成功。", "top.location.href", "submit1", "form1", "top")
      End If
    End Sub

    Public Sub del()
      id = Trim(Request.Form("id"))
      sql = "delete from tr_comment where id="&id&""
      conn.Execute(sql)
    End Sub

  End Class
%>
