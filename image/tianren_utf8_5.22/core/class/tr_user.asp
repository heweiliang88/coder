<%
'版权声明：
'天人文章管理系统，以下简称“本程序、程序”可免费用于个人非商业用途，其他用途需付费购买版权，请按照如下约定执行：
'**本程序可免费用于个人非商业用途，除个人非商业外的用途均需要付费购买版权，并在官网登记域名作为正规版权记录，并可获得电子版版权证书(不是绑定域名，而是在官网数据库中登记域名旨在法律维权时避开已购买版权的域名)，侵权使用必将承担法律风险并追究高额赔偿
'**购买版权将获得稳定健全的升级服务、免费安装插件、模板、配套工具免费使用、官方优惠优先获得等特权。并可用于企业、政府、校园、单位、团体、个人、机构、部门、机关、公司及其他商业用途
'**购买版权后，除源代码中版权声明，后台左下角的“源程序：55tr.com”字样（这两种版权字样都不会在前台显示，只作为程序版权归属）以外所有代码都可以修改及删除，包括前台页面底部的官网链接及版权等代码都可修改。
'**购买版权后可用本程序为别人或自己提供建设网站，修改源码，二次开发的服务，但禁止发布或传播本程序及本程序的任何形态！禁止出售本程序及本程序的任何形态。说白了就是不能将程序出售或传播，否则侵权。（例如：小李与客户约定使用本程序制作某中学的官网，并收取某中学的网站制作服务费用1万元，并购买版权，此行为不侵权。小赵与客户约定使用本程序制作某小学的官网，并收取某小学的网站制作服务费用1万元，未购买版权，此行为侵权(因为未购买版权)，请终止！小孙使用本程序建立了个人漫画网站供用户浏览并获取广告收益，未购买版权，此行为不侵权。小张使用本程序制作为企业网站进行出售，并购买版权，此行为侵权（因为传播出售了程序），请终止！小王又将本程序制作为小说网站源码在网络免费发布传播提供下载，并购买版权，此行为侵权（因为传播出售了程序），请终止！）
'**不购买版权的个人非商业用途依然享受升级、安装插件、模板的权利，可修改源代码，但不得修改源代码中“官网链接”、“版权声明”、“应用中心”相关的代码，前台页面底部的“Copyright 2016 天人文章管理系统 版权所有，授权xxx使用 Powered by 55TR.COM”字样不得删除,否则请购买版权
'**不购买版权的除个人非商业外的用途请及时购买版权或终止使用本程序，否则侵权
'**计算机软件著作权登记证书登记号：2016SR204636，侵权必究，作者QQ81962480 欢迎洽谈基于此程序开发、修改业务。
'**版权声明会存在修改的偶然性，请以官网公布的对应程序版本版权声明为准！具体的版权声明、特权、及功能介绍，请浏览：http://www.55tr.com/gwappshow.asp?id=11
'**天人文章管理系统3.48之前的程序版本版权声明以此为准，之前的版权声明作废



Class tr_user
  Public id
  Public username
  Public md5username
  Public passwd0
  Public passwd
  Public passwd1
  Public passwd2
  Public nickname
  Public qq
  Public homepage
  Public tel
  Public regtime
  Public lastlogintime
  Public lastsigntime
  Public signs
  Public gavesignpoints
  Public regip
  Public lastloginip
  Public age
  Public mail
  Public sex
  Public address
  Public zip
  Public birthday
  Public points
  Public signature
  Public intro
  Public photo
  Public headpic
  Public isfree
  Public gavecommentpoints
  Public gavearticlepoints
  Public ismailactivated
  Public isdel
  Public articlenum
  Public commentnum
  Public mailcode
  Public opencomment
  Public openarticle
  Public openuserreg
  Public openuserlogin
  Public edittime
  Public logcount
  Public safemd5

  Public Sub getfieldv(id)
    sql = "select * from tr_user where id="&id&""
    rs.Open sql, conn, 0, 1
    If Not rs.EOF Then
      Call getrs(2)
    End If
    rs.Close
  End Sub
  
  Public Sub usersign()
    safemd5 = trmd5(session(siternd & "_55tr.uname"))
    fnow = FormatDate(Now(), 1)
    sql = "select * from tr_user where safemd5='"&safemd5&"' and (datediff('d','"&fnow&"',lastsigntime)<>0 or lastsigntime is null )"
    rs.Open sql, conn, 1, 3
    If Not rs.EOF Then
      If DateDiff("d", rs("lastsigntime"), fnow) = 1 Then
        rs("signs") = rs("signs") + 1
      Else
        rs("signs") = 1
      End If
      rs("gavesignpoints") = rs("gavesignpoints") + application(siternd & "55trsignpoints")
      rs("lastsigntime") = fnow
      rs("points") = rs("points") + application(siternd & "55trsignpoints")
      rs.update
      rs.Close
'      Call contrl_message("已签到成功。", "", "", "", "")
    Else
      rs.Close
'      Call contrl_message("您今天已经签过到了。", "", "", "", "")
    End If
  End Sub

  Public Sub udelnew()
    id = tr_killstr(request.Form("id"), 1, 8, 1, "", "", "", "", 2)
    safemd5 = trmd5(session(siternd & "_55tr.uname"))
    sql = "select isdel from tr_article where contrimd5 ='"&safemd5&"' and id="&id&" and ispass=0 "
    rs.Open sql, conn, 1, 3
    If Not rs.EOF Then
      rs("isdel") = 1
      rs.update
      Call contrl_message("删除成功。", "", "", "", "top")
    Else
      Call contrl_message("出现错误，无法删除。", "", "", "", "top")
    End If
    rs.Close
  End Sub

  Public Sub ueditnews(how)'add为添加文章，edit为编辑文章
    safemd5 = trmd5(session(siternd & "_55tr.uname"))
    sql = "select * from tr_user where safemd5='"&safemd5&"'"
    rs.Open sql, conn, 0, 1
    Call getrs(1)
    rs.Close
    title = tr_killstr(request.Form("title"), 3, 100, 1, "", "submit1", "", "", 2)
    titlecolor = tr_killstr(request.Form("titlecolor"), 3, 16, 1, "", "submit1", "", "", 1)
    content = tr_killstr(request.Form("content"), 4, 99999, 1, "", "submit1", "", "", 2)
    keywords = tr_killstr(request.Form("keywords"), 3, 200, 1, "", "submit1", "", "", 1)
    descriptions = tr_killstr(request.Form("descriptions"), 3, 200, 1, "", "submit1", "", "", 1)
    columnid = tr_killstr(request.Form("columnid"), 2, 100, 1, "", "submit1", "", "", 2)
    id = tr_killstr(request.Form("id"), 1, 100, 1, "", "submit1", "", "", 2)
    columnname = getfieldvalue("tr_column", "colname", " and id="&columnid&" ", 1, "order by id desc")
    queueid = getfieldvalue("tr_column", "queueid", " and id="&columnid&" ", 1, "order by id desc")
    contributor = username
    contrimd5 = safemd5
    iscontribute = 1
    If how = "add" Then
      sql = "select * from tr_article where 1=2 "
      rs.Open sql, conn, 1, 3
      rs.addnew
      id = GetMaxValue("tr_article", "id") + 1
    ElseIf how = "edit" Then
      sql = "select * from tr_article where id="&id&" and contrimd5='"&safemd5&"' and isdel=0 and ispass=0 "
      rs.Open sql, conn, 1, 3
      If rs.EOF Then Call contrl_message("该文章无法编辑，请编辑其他文章。", "", "", "", "")
    End If
    Rs("id") = id
    Rs("channelid") = 1
    Rs("title") = title
    Rs("titlecolor") = titlecolor
    Rs("content") = content
    Rs("keywords") = keywords
    Rs("descriptions") = descriptions
    Rs("addtime") = Now()
    Rs("author") = "网络投稿"
    Rs("sources") = "网络投稿"
    Rs("inputer") = "网络投稿"
    Rs("columnid") = columnid
    Rs("queueid") = queueid
    Rs("columnname") = columnname
    Rs("isfocus") = 0
    Rs("isroll") = 0
    Rs("isindextop") = 0
    Rs("islisttop") = 0
    Rs("iscommend") = 0
    Rs("isrollimg") = 0
    Rs("picfiles") = ""
    Rs("pgcount") = 1
    Rs("isdel") = 0
    Rs("iscontribute") = iscontribute
    Rs("contributor") = contributor
    Rs("contrimd5") = contrimd5
    Rs("isabout") = 0
    Rs("isgavepoints") = 0
    If application(siternd & "55tropencheckarticle") = 0 Then
      Rs("ispass") = 1
      Rs("isgavepoints") = 1
      sql = "select * from tr_user where safemd5='"&safemd5&"'"
      rs2.Open sql, conn, 1, 3
      rs2("gavearticlepoints") = rs2("gavearticlepoints") + application(siternd & "55trpublisharticlepoints")
      rs2("points") = rs2("points") + application(siternd & "55trpublisharticlepoints")
      rs2.update
      rs2.Close
	  givejf=true
    Else
      Rs("ispass") = 0
	  givejf=false
    End If
    rs.update
    rs.Close
	if givejf=true and how="add" then
response.write "<script type=""text/javascript"">"
response.write "parent.editor.html("""");"
response.write "</script> "
	Call contrl_message("发布文章成功，并赠送您"&application(siternd & "55trpublisharticlepoints")&"积分。", "", "", "form10", "")
	elseif givejf=false and how="add" then
response.write "<script type=""text/javascript"">"
response.write "parent.editor.html("""");"
response.write "</script> "
	Call contrl_message("发布文章成功，待审核通过后，会赠送您"&application(siternd & "55trpublisharticlepoints")&"积分。", "", "", "form10", "")
	end if
	if how="edit" then Call contrl_message("发布文章成功。", funpath&"../user/mynews.asp", "", "form1"&id&"", "top")

  End Sub

  Function truserislog(alert)'alert,1警告，2不警告
    sessionuname = session(siternd & "_55tr.uname")
    If Len(sessionuname) >1 Then
      sql = "select * from tr_user where safemd5='"&trmd5(sessionuname)&"'"
      rs.Open sql, conn, 0, 1
      If Not rs.EOF Then
        Call getrs(2)
        truserislog = session(siternd & "_55tr.uname")
      Else
        If alert = 1 Then
          Call contrl_message("您尚未登录，请登陆后进行操作。", funpath&"user/login.asp", "", "", "top")
        Else
          response.End()
        End If
      End If
      rs.Close
    Else
      truserislog = ""
      If alert = 1 Then
        Call contrl_message("您尚未登录，请登陆后进行操作。", funpath&"user/login.asp", "", "", "top")
      End If
    End If
  End Function

  Function truserlogin()
    gonext = 0
    nowtime = FormatDate(Now(), 1)
    username = LCase(tr_killstr(request.Form("username"), 5, 50, 2, "", "submit1", "", "", 2))
    passwd = left(request.Form("passwd"), 50)
    vcode = tr_killstr(request.Form("vcode"), 1, 4, 2, "", "submit1", "", "", 2)
    If vcode = tr_killstr(session("getcode"), 1, 4, 2, "", "submit1", "", "", 2) Then
      gonext = 1
    Else
      Call contrl_message("验证码错误请重新输入", "./login.asp", "", "", "top")
    End If
    If gonext = 1 Then
      md5username = tomd5(username)
      md5passwd = tomd5(passwd)
      sql = "select ID,isfree,isdel,lastloginip,username,lastlogintime,ismailactivated,logcount,points,safemd5 from tr_user where md5username='"&md5username&"' and passwd='"&md5passwd&"' "
      rs.Open sql, conn, 1, 3
      If Not rs.EOF Then
        If rs("isfree") = 0 Then
          Call contrl_message("该用户已锁定或需要审核后方能使用，请与管理员联络解锁。", "", "submit1", "form1", "")
        ElseIf rs("isdel") = 1 Then
          Call contrl_message("该用户已删除，请与管理员联络恢复此用户或重新注册。", "", "submit1", "form1", "")
        ElseIf rs("ismailactivated") = 0 Then
          Call contrl_message("该用户邮箱未激活，请激活后进行登陆。", "", "submit1", "form1", "")
        Else
          session(siternd & "_55tr.uname") = Trim(rs("username"))
          lstip = rs("lastloginip")
          If Len(lstip)>7 Then
            ublstip = UBound(Split(lstip, ","))
          ElseIf ublstip<10 Then
            lstip = lstip&","&getip()
          ElseIf ublstip = 10 Then
            lstip = lstip&",|"&getip()
          ElseIf ublstip>10 And ublistip<20 Then
            lstip = lstip&","&getip()
          ElseIf ublstip>= 20 Then
            splstip2 = Split(lstip, "|")
            lstip = splstip2(0)&Mid(splstip2(1), InStr(splstip2(1), ","), Len(splstip2(1)))&","&getip()
          End If
          If DateDiff("d", rs("lastlogintime"), nowtime)<>0 Or IsNull(rs("lastlogintime")) Then
            rs("points") = rs("points") + application(siternd & "55trloginpoints")
            If IsNull(rs("logcount")) Or rs("logcount") = "" Then
              rs("logcount") = 1
            Else
              rs("logcount") = rs("logcount") + 1
            End If
          End If
          rs("lastlogintime") = nowtime
          rs("lastloginip") = lstip
          rs.update
        End If
        Response.Redirect("./sign.asp")
      Else
        Call contrl_message("登录不成功，用户名或密码错误，请重新输入。", "", "submit1", "form1", "")
      End If
      rs.Close
    End If
  End Function

  Public Sub userreg()
    Call getform(1)
    If Len(passwd1)<5 Or passwd1<>passwd2 Then
      Call contrl_message("两次密码不同或密码长度小于5个字符，请重新输入", "", "submit1", "form1", "")
    End If
    vcode = tr_killstr(request.Form("vcode"), 1, 4, 2, "", "submit1", "", "", 2)
    If vcode<>tr_killstr(session("getcode"), 1, 4, 2, "", "submit1", "", "", 2) Then
      Call contrl_message("验证码错误请重新输入", "", "submit1", "form1", "")
    End If
    md5username = tomd5(username)
    passwd = tomd5(passwd1)
    regtime = Now()
    lastlogintime = ""
    regip = getIP()
    lastloginip = ""
    points = application(siternd & "55trinitialpoints")
    photo = ""
    headpic = ""
    If application(siternd & "55tropencheckuser") = 1 Then
      isfree = 0
    Else
      isfree = 1
    End If
    gavecommentpoints = 0
    gavearticlepoints = 0
    ismailactivated = 1
    isdel = 0
    articlenum = 0
    commentnum = 0
    mailcode = 0
    opencomment = 1
    openarticle = 1
    openuserlogin = 1
    edittime = ""
    logcount = 0
    sql = "select * from tr_user where 1=2 "
    rs.Open sql, conn, 1, 3
    rs.addnew()
    Call toupdate(rs, 1)
    rs("passwd") = passwd
    Rs("regtime") = regtime
    rs.update
    rs.Close
    Call contrl_message("注册成功，现在就可以登陆了。", "", "", "form1", "")
  End Sub
  
  Public Sub useredit()'requesttype接收类型，1会员注册，2会员编辑，3管理员添加，4管理员编辑
    safemd5 = trmd5(session(siternd & "_55tr.uname"))
    Call getform(2)
    md5passwd0 = tomd5(passwd0)
    If passwd2<>"" Then
      If Len(passwd1)<5 Or passwd1<>passwd2 Then
        Call contrl_message("两次密码不同或密码长度小于5个字符，请重新输入", "", "submit1", "form1", "")
      End If
      passwd = tomd5(passwd1)
    Else
      passwd = md5passwd0
    End If
    edittime = Now()
    sql = "select * from tr_user where safemd5='"&safemd5&"' and passwd='"&md5passwd0&"' "
    rs.Open sql, conn, 1, 3
    If Not rs.EOF Then
      Call toupdate(rs, 2)
      Rs("passwd") = passwd
      rs.update
      Call contrl_message("资料修改成功", "", "submit1", "form1", "")
    Else
      Call contrl_message("原始密码不正确，请重新输入", "", "submit1", "form1", "")
    End If
    rs.Close
  End Sub

  Public Sub adminadduser()
    Call getform(3)
    If Len(passwd1)<5 Or passwd1<>passwd2 Then
      Call contrl_message("两次密码不同或密码长度小于5个字符，请重新输入", "", "sutmit1", "form1", "")
    End If
    md5username = tomd5(username)
    passwd = tomd5(passwd1)
    regtime = Now()
    lastlogintime = ""
    regip = getIP() &"管理员ID:"&adminid&"添加"
    lastloginip = ""
    points = application(siternd & "55trinitialpoints")
    photo = ""
    headpic = ""
    isfree = 1
    gavecommentpoints = 0
    gavearticlepoints = 0
    ismailactivated = 1
    isdel = 0
    articlenum = 0
    commentnum = 0
    mailcode = 0
    opencomment = 1
    openarticle = 1
    openuserlogin = 1
    edittime = ""
    logcount = 0
    sql = "select * from tr_user where 1=2 "
    rs.Open sql, conn, 1, 3
    rs.addnew()
    Call toupdate(rs, 3)
    Rs("passwd") = passwd
    rs.update
    rs.Close
  End Sub

  Public Sub adminedit()
    Call getform(4)
    If passwd2<>"" Then
      If Len(passwd1)<5 Or passwd1<>passwd2 Then
        Call contrl_message("两次密码不同或密码长度小于5个字符，请重新输入", "", "submit1", "form1", "")
      End If
      passwd = tomd5(passwd1)
    End If

    edittime = Now()
    sql = "select * from tr_user where id="&id&""
    rs.Open sql, conn, 1, 3
    Call toupdate(rs, 4)
    If passwd2<>"" Then
      Rs("passwd") = passwd
    End If
    rs.update
    rs.Close
  End Sub

  Public Sub del()
    id = Trim(Request.Form("id"))
    sql = "delete from tr_user where id="&id&""
    conn.Execute(sql)
  End Sub

  Public Function getform(requesttype)'requesttype接收类型，1会员注册，2会员编辑，3管理员添加，4管理员编辑
    If requesttype = 2 Then
      passwd0 = Trim(Request.Form("passwd0"))
    End If
    If requesttype<>2 And requesttype<>4 Then
      username = LCase(tr_killstr(Trim(request.Form("username")), 5, 30, 2, "", "submit1", "", "", 2))
      If InStr(application(siternd & "55trforbidusername"), ","&username&",")>0 Then Call contrl_message ("登录名存在敏感词，请更换。", "", "submit1", "form1", "")
      If getfieldvalue("tr_user", "id", " and safemd5='"&trmd5(username)&"'", 1, " order by id desc ")<>"" Then Call contrl_message ("登录名已存在，请更换。", "", "submit1", "form1", "")
      safemd5 = trmd5(username)
    End If
    If requesttype<>2 Then
      id = tr_killstr(Trim(request.Form("id")), 1, 8, 3, "", "submit1", "", "", 1)
      If id = 0 Then id = GetMaxValue("tr_user", "id") + 1
      mail = LCase(tr_killstr(Trim(request.Form("mail")), 3, 30, 2, "", "submit1", "", "", 2))
      If getfieldvalue("tr_user", "id", " and mail='"&mail&"' and safemd5<>'"&trmd5(username)&"' " , 1, " order by id desc ")<>"" Then Call contrl_message ("邮箱已存在，请更换。", "", "submit1", "form1", "")
      points = Trim(Request.Form("points"))
      isfree = Trim(Request.Form("isfree"))
      If isfree = "" Then isfree = 0
      gavecommentpoints = Trim(Request.Form("gavecommentpoints"))
      If gavecommentpoints = "" Then gavecommentpoints = 0
      gavearticlepoints = Trim(Request.Form("gavearticlepoints"))
      If gavearticlepoints = "" Then gavearticlepoints = 0
      isdel = Trim(Request.Form("isdel"))
      If isdel = "" Then isdel = 0
      opencomment = Trim(Request.Form("opencomment"))
      If opencomment = "" Then opencomment = 0
      openarticle = Trim(Request.Form("openarticle"))
      If openarticle = "" Then openarticle = 0
      openuserreg = Trim(Request.Form("openuserreg"))
      If openuserreg = "" Then openuserreg = 0
      openuserlogin = Trim(Request.Form("openuserlogin"))
      If openuserlogin = "" Then openuserlogin = 0
      articlenum = Trim(Request.Form("articlenum"))
      If articlenum = "" Then articlenum = 0
      commentnum = Trim(Request.Form("commentnum"))
      If commentnum = "" Then commentnum = 0
    End If
    passwd1 = Trim(Request.Form("passwd1"))
    passwd2 = Trim(Request.Form("passwd2"))
    nickname = tr_killstr(Trim(request.Form("nickname")), 3, 30, 2, "", "submit1", "", "", 1)
    qq = tr_killstr(Trim(request.Form("qq")), 3, 30, 2, "", "submit1", "", "", 1)
    homepage = tr_killstr(Trim(request.Form("homepage")), 3, 30, 2, "", "submit1", "", "", 1)
    tel = tr_killstr(Trim(request.Form("tel")), 1, 30, 3, "", "submit1", "", "", 1)
    If tel = 0 Then tel = null
    age = tr_killstr(Trim(request.Form("age")), 1, 30, 3, "", "submit1", "", "", 1)
    If age = 0 Then age = null
    sex = tr_killstr(Trim(request.Form("sex")), 1, 30, 3, "", "submit1", "", "", 1)
    If sex = 0 Then sex = null
    zip = tr_killstr(Trim(request.Form("zip")), 1, 30, 3, "", "submit1", "", "", 1)
    If zip = 0 Then zip = null
    address = tr_killstr(Trim(request.Form("address")), 3, 50, 2, "", "submit1", "", "", 1)
    birthday = Trim(request.Form("birthday"))
    If Not IsDate(birthday) Then birthday = null
    signature = tr_killstr(Trim(request.Form("signature")), 3, 300, 2, "", "submit1", "", "", 1)
    intro = tr_killstr(Trim(request.Form("intro")), 3, 300, 2, "", "submit1", "", "", 1)
    photo = tr_killstr(Trim(request.Form("photo")), 3, 50, 2, "", "submit1", "", "", 1)
    headpic = tr_killstr(Trim(request.Form("headpic")), 3, 50, 2, "", "submit1", "", "", 1)
  End Function

  Private Sub toupdate(ByRef Rs, requesttype)'接收类型，1会员注册，2会员编辑，3管理员添加，4管理员编辑
    If requesttype = 1 Or requesttype = 3 Then
      Rs("id") = id
      Rs("username") = username
      Rs("md5username") = md5username
      rs("safemd5") = safemd5
      Rs("regtime") = regtime
      Rs("regip") = regip
      Rs("signs") = 0
      Rs("gavesignpoints") = 0
    End If
    If requesttype<>2 Then
      Rs("points") = points
      Rs("isfree") = isfree
      Rs("mail") = mail
      Rs("gavecommentpoints") = gavecommentpoints
      Rs("gavearticlepoints") = gavearticlepoints
      Rs("ismailactivated") = ismailactivated
      Rs("articlenum") = articlenum
      Rs("commentnum") = commentnum
      Rs("opencomment") = opencomment
      Rs("openarticle") = openarticle
      Rs("openuserreg") = openuserreg
      Rs("openuserlogin") = openuserlogin
    End If
    If requesttype = 4 Then
      Rs("isdel") = isdel
    End If
    If requesttype = 2 Or requesttype = 4 Then
      Rs("edittime") = edittime
    End If
    Rs("nickname") = nickname
    Rs("qq") = qq
    Rs("homepage") = homepage
    Rs("tel") = tel
    Rs("age") = age
    Rs("sex") = sex
    Rs("address") = address
    Rs("zip") = zip
    Rs("birthday") = birthday
    Rs("signature") = signature
    Rs("intro") = intro
    'Rs("photo")=photo
    'Rs("headpic")=headpic
  End Sub

  Private Sub getrs(gettype)'1为获取用户名和md5的用户名，2是获取所有
    username = Rs("username")
    If gettype<>1 Then
      id = Rs("id")
      md5username = Rs("md5username")
      passwd = Rs("passwd")
      nickname = Rs("nickname")
      qq = Rs("qq")
      homepage = Rs("homepage")
      tel = Rs("tel")
      regtime = Rs("regtime")
      lastlogintime = Rs("lastlogintime")
      regip = Rs("regip")
      lastloginip = Rs("lastloginip")
      age = Rs("age")
      mail = Rs("mail")
      sex = Rs("sex")
      address = Rs("address")
      zip = Rs("zip")
      birthday = Rs("birthday")
      points = Rs("points")
      signature = Rs("signature")
      intro = Rs("intro")
      photo = Rs("photo")
      headpic = Rs("headpic")
      isfree = Rs("isfree")
      gavecommentpoints = Rs("gavecommentpoints")
      gavearticlepoints = Rs("gavearticlepoints")
      ismailactivated = Rs("ismailactivated")
      isdel = Rs("isdel")
      articlenum = Rs("articlenum")
      commentnum = Rs("commentnum")
      mailcode = Rs("mailcode")
      opencomment = Rs("opencomment")
      openarticle = Rs("openarticle")
      openuserreg = Rs("openuserreg")
      openuserlogin = Rs("openuserlogin")
      edittime = Rs("edittime")
      signs = Rs("signs")
      lastsigntime = Rs("lastsigntime")
      gavesignpoints = Rs("gavesignpoints")
      logcount = Rs("logcount")
    End If
  End Sub

End Class
%>
