<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../core/Conn.asp"-->
<!--#include file="../core/fun/fun.asp"-->
<!--#include file="../core/fun/core.asp"-->
<!--#include file="../core/class/tr_page.asp"-->
<!--#include file="../crinc/ulevstr.asp"-->
<!--#include file="../core/class/tr_user.asp"-->
<!--#include file="../core/md5.asp"-->
<!--#include file="charsetasp.asp"-->
<%
Const dbpath = "../"
Set ctr_page = New tr_page
Set ctr_user = New tr_user
id = tr_killstr(request.querystring("id"), 1, 8, 1, "", "", "", "", 2)
pageno = tr_killstr(request.querystring("pageno"), 1, 8, 1, "", "", "", "", 1)
funpath = request.querystring("funpath")
sql = "select * from tr_article where ispass=1 and isdel=0 and id="&id&" "
rs.Open sql, conn, 0, 1
If rs.EOF Or rs.bof Then
  Call contrl_message ("���²����ڡ�", "", "", "", "")
Else
  ispass = Rs("ispass")
  isdel = Rs("isdel")
  If ispass = 1 And isdel = 0 Then
    id = Rs("id")
    title = Rs("title")
    content = Rs("content")
    keywords = Rs("keywords")
    keywords2 = Rs("keywords2")
    descriptions = Rs("descriptions")
    addtime = Rs("addtime")
    pgcount = Rs("pgcount")
    clicks = Rs("clicks")
    author = Rs("author")
    sources = Rs("sources")
    inputer = Rs("inputer")
    limittype = Rs("limittype")
    If Not IsNumeric(limittype) Then limittype = 0
    limitpoints = Rs("limitpoints")
    If Not IsNumeric(limitpoints) Then limitpoints = 0
    limitmore = Rs("limitmore")
    columnid = Rs("columnid")
    iscontribute = Rs("iscontribute")
    contributor = Rs("contributor")
  End If
End If
rs.Close
uname = ctr_user.truserislog(2)
%>
var tmpstr=""
<%
If uname = "" Then'���û���Ϊ�տ�ʼ
%>
tmpstr=tmpstr+("<div class=\"trsnotice\">");
tmpstr=tmpstr+("��������Ҫ��");
<%if limittype>0 then%>
<%
'inidstr=instr(idstr,","&limittype&",")
If InStr(idstr, ","&limittype&",")>0 Then
  spidstr = Split(idstr, ",")
  For idi = 0 To UBound(spidstr)
    If spidstr(idi)<>"" Then
      If clng(limittype) = clng(spidstr(idi)) Then
        arri = idi
        Exit For
      End If
    End If
  Next
Else
  response.End()
End If
%>
tmpstr=tmpstr+("�� <%=Split(levnamestr, ",")(arri -1)%> ���ȼ�<%if limitpoints>0 then%>������<%end if%>");<%end if%>
tmpstr=tmpstr+("<%if limitpoints>0 then%>�� ���ִ��� <%=limitpoints%> ��<%end if%>�ſ��������<br>����δ��¼��<a href=\"<%=funpath%>user/login.asp\" target=\"_blank\">���ȵ���˴����е�½�鿴�Ķ�Ȩ�ޡ�</a>");
tmpstr=tmpstr+("</div>")
<%
Else
  gonext = true
  If limittype>0 Then
    If InStr(idstr, ","&limittype&",")>0 Then
      spidstr = Split(idstr, ",")
      For idi = 0 To UBound(spidstr)
        If spidstr(idi)<>"" Then
          If clng(limittype) = clng(spidstr(idi)) Then
            arri = idi
            Exit For
          End If
        End If
      Next
    Else
      response.End()
    End If
    spuserjfstr = Split(userjfstr, ",")
    splevnamestr = Split(levnamestr, ",")
    needjf = clng(spuserjfstr(arri -1))
    points = clng(ctr_user.points)
    If points<needjf Then
%>
tmpstr=tmpstr+("<div class=\"trsnotice\">");
tmpstr=tmpstr+("��������Ҫ��Ա�ȼ�Ϊ��<%=splevnamestr(arri-1)%>�����ĵȼ�Ϊ��");
<%
For i = 0 To UBound(spuserjfstr)
  If points>= clng(spuserjfstr(i)) And points<clng(spuserjfstr(i + 1)) Then
    response.Write "tmpstr=tmpstr+("""&splevnamestr(i)&""");"
    Exit For
  ElseIf points>= clng(spuserjfstr(UBound(spuserjfstr))) Then
    response.Write "tmpstr=tmpstr+("""&splevnamestr(UBound(spuserjfstr))&""");"
    Exit For
  End If
Next
%>
tmpstr=tmpstr+("��");
<%
gonext = false
End If
End If
If limitpoints>1 Then
  points = clng(ctr_user.points)
  If points<limitpoints Then
    If gonext = false Then
%>
tmpstr=tmpstr+("��ͬʱ");<%else%>
tmpstr=tmpstr+("<div class=\"trsnotice\">");

tmpstr=tmpstr+("������");<%end if%>
tmpstr=tmpstr+("��Ҫ����Ϊ��<%=limitpoints%>�����Ļ���Ϊ����<%=points%>��");
tmpstr=tmpstr+("</div>")
<%
gonext = false
End If
End If
If gonext Then
  response.Write "var tmpstr='';"&vbCrLf
  response.Write tojs2(tr_arcontent(content, 700, 1, "tdload", pageno), "tmpstr")
End If
End If'���û���Ϊ�ս���
%>
function changeText(tmp){
top.document.getElementById('trcontenttd').innerHTML = tmp;
//top.document.getElementById('showcomment').className  = "trcomment trbor2";
}
changeText(tmpstr);
