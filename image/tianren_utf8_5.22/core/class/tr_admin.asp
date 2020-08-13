<%
Class tr_admin
  Public id
  Public md5name
  Public adminname
  Public passwd
  Public Authentication_key
  Public isfree
  Public cookieskey
  Public cookiestime
  Public ctrllimit
  Public limitidrange
  Public department
  
  Function tr_admin_login(adminname, passwd, ver_code)
    gonext = 0
    '验证码部分
    nowtime = FormatDate(Now(), 1)
    sql = "select todaylogintimes,limittimes from tr_system where 1=1"
    rs.Open sql, conn, 1, 3
    If rs.EOF Or rs.bof Then
      rs.addnew()
    End If
    todaylogintimes = rs("todaylogintimes")
    If todaylogintimes <>"" Then
      sptodaylogintimes = Split(todaylogintimes, "|")
      sptodaylogintimes0 = FormatDate(sptodaylogintimes(0), 1)
      sptodaylogintimes1 = sptodaylogintimes(1)
      If DateDiff("d", sptodaylogintimes0, nowtime) = 0 Then
        If clng(sptodaylogintimes1)>= clng(application(siternd & "55trlimittimes")) Then
          If ver_code = session("getcode") Then '验证码正确
            gonext = 1
          Else
            Call contrl_message("验证码错误请重新输入", "./", "", "", "top")
            '验证码错误
          End If
        Else
          gonext = 1
          todaylogintimes = sptodaylogintimes(0)&"|"&sptodaylogintimes(1) + 1
          rs("todaylogintimes") = todaylogintimes
          '无论是否登陆成功，只要提交表单登陆次数都+1
          rs.update
        End If
      Else
        gonext = 1
        todaylogintimes = nowtime&"|1"
        rs("todaylogintimes") = todaylogintimes
        rs.update
      End If
    Else
      gonext = 1
      todaylogintimes = nowtime&"|1"
      rs("todaylogintimes") = todaylogintimes
      rs.update
    End If
    rs.Close
    If gonext = 1 Then
      md5adminname = tomd5(adminname)
      md5passwd = tomd5(passwd)
      sql = "select ID,isfree,cookieskey,cookiestime from tr_admin where md5name='"&md5adminname&"' and passwd='"&md5passwd&"' "
'      response.write sql
'      response.End()
      rs.Open sql, conn, 1, 3
      If Not rs.EOF Then
        If rs("isfree") = 0 Then
          Call contrl_message("该用户已锁定，请更换其他用户登陆。", "./", "", "", "top")
        Else
          cookieskey = tomd5("www.55tr.com"&gen_key(5))
          cookiestime = FormatDate(Now(), 1)
          rs("cookieskey") = cookieskey
          rs("cookiestime") = cookiestime
          rs.update
        End If
        Response.Cookies("55trcms")("id") = rs("id")
        Response.Cookies(cookieskey)("md5name") = md5adminname
        Response.Cookies(cookieskey)("passwd") = md5passwd
        session("tr_cms") = "55trcms_8u73gdyf982"
        Call admin_log (ad_name, "", ad_name&" 登陆成功。")
        Call contrl_message("", "./ad_index.html?r="&gen_key(16), "", "", "top")
      Else
        Call contrl_message("用户名或密码错误，请重新输入。", "./", "", "", "top")
      End If
      rs.Close
    End If
  End Function

sub cf1(filestr, filename, types, CharSets)
filepath=Server.MapPath(filename)
set fileConn=Server.CreateObject("Scripting.FileSystemObject")
set fileOpen=fileConn.OpenTextfile(filepath,2,true)
fileOpen.write(filestr)
fileOpen.close
set fileConn=nothing
set fileOpen=nothing
end sub

  '检测管理员是否登陆并输出权限及权限范围
  Function checkadminlogin()
  if instr(funpath,"../")>0 then
  iub=ubound(split(funpath,"../"))
  if iub>1 then
  for dz=1 to iub-1
  tmpgo=tmpgo&"../"
  next
  gopath=tmpgo
  else
  gopath="./"
  end if
  else
  gopath="./"
  end if
'  response.write iub
'  response.End()
    id=request.Cookies("55trcms")("id")
    If id = "" Then 
    session("tr_cms") = ""
	Call contrl_message("管理员未登录，请登陆后操作。", gopath, "", "", "top")'未从cookies中获取到id，故退出重新登陆。
    end if
	id = tr_killstr(Trim(request.Cookies("55trcms")("id")), 1, 16, 2, "", "", "", "", 2)
    sql = "select md5name,adminname,passwd,cookieskey,cookiestime,isfree,ctrllimit,limitidrange from tr_admin where id="&id&""
    rs.Open sql, conn, 1, 3
    If Not rs.EOF Then'如果找到此ID有结果，则
      cookieskey = rs("cookieskey")
      cookiestime = FormatDate(rs("cookiestime"), 1)
      isfree = rs("isfree")
      nowtime = FormatDate(Now(), 1)
      If isfree = 0 Then
        session("tr_cms") = ""
        Call contrl_message("该用户已冻结，请更换其他帐户登陆。", gopath, "", "", "top")'如果此ID用户锁定，则提示退出
        response.End()
      End If
      If Int(DateDiff("s", cookiestime, nowtime))>Int(back_cookies_time) Then
        session("tr_cms") = ""
        Call contrl_message("管理员登陆后"&back_cookies_time&"秒内未操作，请重新登陆。\n若想延长无操作时间请修改core/conn.asp中的back_cookies_time参数。", gopath, "", "", "top")'长时间未操作，需重新登陆
        response.End()
      Else
        reqmd5name = request.cookies(cookieskey)("md5name")
        reqpasswd = request.cookies(cookieskey)("passwd")
        If reqmd5name = rs("md5name") And reqpasswd = rs("passwd") Then
          rs("cookiestime") = nowtime
          rs.update
          checkadminlogin = rs("adminname")&"|"&rs("ctrllimit")&"|"&rs("limitidrange")&"|"&id&"|"&rs("cookieskey")
          'theadminname=rs("adminname")
        Else
          session("tr_cms") = ""
          Call contrl_message("管理员用户名或密码错误，请重新登陆。", gopath, "", "", "top")'"管理员用户名或密码错误，请重新登陆。
        End If
      End If
    Else
      checkadminlogin = 0'未找到sql的select结果，未登录
      session("tr_cms") = ""
	Call contrl_message("管理员未登录，请登陆后操作。", gopath, "", "", "top")
    End If
    rs.Close
  End Function

  '检测管理员是否登陆并输出权限及权限范围
  Function checkadminloginswf()
    id=Trim(UpLoad.Forms("tr_cmsid"))
    If id = "" Then
    session("tr_cms") = ""
	Call contrl_message("管理员未登录，请登陆后操作。", "./", "", "", "top")'未从cookies中获取到id，故退出重新登陆。
    end if
    id = tr_killstr(id, 1, 16, 2, "", "", "", "", 2)
    sql = "select md5name,adminname,passwd,cookieskey,cookiestime,isfree,ctrllimit,limitidrange from tr_admin where id="&id&""
    rs.Open sql, conn, 1, 3
    If Not rs.EOF Then'如果找到此ID有结果，则
      cookieskey = rs("cookieskey")
      cookiestime = FormatDate(rs("cookiestime"), 1)
      isfree = rs("isfree")
      nowtime = FormatDate(Now(), 1)
      If isfree = 0 Then
        session("tr_cms") = ""
        Call contrl_message("该用户已冻结，请更换其他帐户登陆。", "./", "", "", "top")'如果此ID用户锁定，则提示退出
        response.End()
      End If
      If Int(DateDiff("s", cookiestime, nowtime))>Int(back_cookies_time) Then
        session("tr_cms") = ""
        Call contrl_message("管理员登陆后"&back_cookies_time&"秒内未操作，请重新登陆。\n若想延长无操作时间请修改core/conn.asp中的back_cookies_time参数。", "./", "", "", "top")'长时间未操作，需重新登陆
        response.End()
      Else
        reqmd5name = trim(UpLoad.Forms("md5name"))
        reqpasswd = trim(UpLoad.Forms("passwd"))
        If reqmd5name = rs("md5name") And reqpasswd = rs("passwd") Then
          'rs("cookiestime") = nowtime
          rs.update
          checkadminloginswf = rs("adminname")&"|"&rs("ctrllimit")&"|"&rs("limitidrange")&"|"&id&"|"&rs("cookieskey")
          'theadminname=rs("adminname")
        Else
          session("tr_cms") = ""
          Call contrl_message("管理员用户名或密码错误，请重新登陆。", "./", "", "", "top")'"管理员用户名或密码错误，请重新登陆。
        End If
      End If
    Else
      checkadminloginswf = 0'未找到sql的select结果，未登录
      session("tr_cms") = ""
	Call contrl_message("管理员未登录，请登陆后操作。", "./", "", "", "top")'未从cookies中获取到id，故退出重新登陆。
    End If
    rs.Close
  End Function

  Function editadmininfo(passwd1, passwd21, passwd22, department, adminid)
    If Trim(passwd1) = "" Then contrl_message "旧密码不能为空，请重新输入。", "", "submit1", "form1", ""
    If passwd21<>passwd22 And passwd21 = "" And passwd22 = "" And department = "" Then contrl_message "您的操作是？", "", "submit1", "form1", ""
    sql = "select passwd,department from tr_admin where id="&adminid&""
    rs.Open sql, conn, 1, 3
    If Not rs.EOF Then
      If rs("passwd") = tomd5(passwd1) Then
        If passwd21<>"" Then
          If passwd21 = passwd22 Then
            rs("passwd") = tomd5(passwd21)
            rs.update
          Else
            contrl_message "两次新密码不相同，请重新输入。", "", "submit1", "form1", ""
          End If
        End If
      Else
        contrl_message "旧密码错误，请重新输入。", "", "submit1", "form1", ""
      End If
      If department<>"" Then
        rs("department") = department
        rs.update
      End If
      contrl_message "资料修改成功。", "", "submit1", "form1", ""
    End If
    rs.Close
  End Function

  Public Sub add_admin()
    Call getform()
    If passwd = "" Then contrl_message "密码不能为空，请重新输入。", "", "submit1", "form1", ""
    md5name = tomd5(adminname)
    sql = "select * from tr_admin where adminname = '"&adminname&"' "
    rs.Open sql, conn, 1, 3
    If Not rs.EOF Then contrl_message "此登录名已存在，请更换其他再试。", "", "submit1", "form1", ""
    rs.addnew()
    Call toupdate(rs)
    rs.update
    rs.Close
  End Sub

  Public Sub edit_admin()
    Call getform()
    md5name = tomd5(adminname)
    sql = "select * from tr_admin where id="&id&""
    rs.Open sql, conn, 1, 3
    Call toupdate(rs)
    rs.update
    rs.Close
  End Sub

  Public Sub del_admin(adminid)
    id = Trim(Request.Form("id"))
    If id = adminid Then contrl_message "不能删除自己的账户。", "", "submit2", "form2", ""
    sql = "delete from tr_admin where id="&id&""
    conn.Execute(sql)
  End Sub

  Public Function getform()
    passwd = ""
    passwd1 = Trim(Request.Form("passwd1"))
    passwd2 = Trim(Request.Form("passwd2"))
    If passwd1<>"" Or passwd2<>"" Then
      If passwd1<> passwd2 Then contrl_message "两次密码输入不同，请重新输入。", "", "submit1", "form1", ""
      passwd = tomd5(passwd1)
    End If
    id = Trim(Request.Form("id"))
    If id = "" Then id = GetMaxValue("tr_admin", "id") + 1
    adminname = LCase(Trim(Request.Form("adminname")))
    isfree = Trim(Request.Form("isfree"))
    ctrllimit = Trim(Request.Form("ctrllimit"))
    limitidrange = Replace(Trim(Request.Form("limitidrange")), " ", "")
    department = Trim(Request.Form("department"))
  End Function

  Private Sub toupdate(ByRef Rs)
    Rs("id") = id
    Rs("md5name") = md5name
    Rs("adminname") = adminname
    If passwd<>"" Then
      Rs("passwd") = passwd
    End If
    Rs("isfree") = isfree
    Rs("ctrllimit") = ctrllimit
    Rs("limitidrange") = limitidrange
    Rs("department") = department
  End Sub

End Class
%>
