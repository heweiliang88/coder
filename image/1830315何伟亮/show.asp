<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%'templatev20���д��빩��װ�������ɾ��%>
<!--#include file="inc/charsetasp.asp"-->
<!DOCTYPE html>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<!--#include file="core/Conn.asp"-->
<!--#include file="core/fun/fun.asp"-->
<!--#include file="core/fun/core.asp"-->
<!--#include file="core/class/tr_page.asp"-->
<!--#include file="crinc/ulevstr.asp"-->
<!--#include file="core/class/tr_user.asp"-->
<!--#include file="core/md5.asp"-->
<!--#include file="skin/default/kp171014.asp"-->
<%'include���д��빩��װ�������ɾ��%>
<%
if ishtml=2 then
funpath = "../"
else
funpath = "./"
end if
Const dbpath = "./"
Set ctr_page = New tr_page
Set ctr_user = New tr_user
id = tr_killstr(request.querystring("id"), 1, 8, 1, "", "", "", "", 2)
pageno = tr_killstr(request.querystring("pageno"), 1, 8, 1, "", "", "", "", 1)
dohtml = tr_killstr(request.querystring("dohtml"), 1, 8, 3, "", "", "", "", 1)
if dohtml<>1 then call jumphtml(id,pageno)
sql = "select * from tr_article where ispass=1 and isdel=0 and id="&id&" "
rs.Open sql, conn, 0, 1
If rs.EOF Or rs.bof Then
response.write "<meta http-equiv=""Refresh"" content=""0; url="&funpath&""" />"&vbcrlf&"</head>"&vbcrlf&"</body>"&vbcrlf&"</html>"
response.End()
'  Call contrl_message ("���²����ڡ�", "", "", "", "")
Else
  ispass = Rs("ispass")
  isdel = Rs("isdel")
  If ispass = 1 And isdel = 0 Then
    id = Rs("id")
    title = Rs("title")
    content = Rs("content")
    keywords = Rs("keywords")
'    keywords2 = Rs("keywords2")
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
    isabout = Rs("isabout")
	queueid = Rs("queueid")
	jumpurl = Rs("jumpurl")
  End If
End If
rs.Close
if jumpurl<>"" then
jumpurl = outputurl("", rs("id"), jumpurl, ishtml)
response.write "<meta http-equiv=""Refresh"" content=""0; url="&jumpurl&""" />"&vbcrlf&"</head>"&vbcrlf&"</body>"&vbcrlf&"</html>"
response.end
end if
rightinccolid=columnid
owncolname=getfieldvalue("tr_column", "colname", "and id="&columnid&"", 1, "")
%>
<%'head���д��빩��װ�������ɾ��%>
<!--#include file="inc/headtaginc.asp"-->
<title><%=title%><%if cint(pageno)>1 then%>[<%=pageno%>]<%end if%>_<%=owncolname%>_<%=application(siternd & "55trsitetitle")%></title>
<meta name="Keywords" content="<%=keywords%>">
<meta name="Description" content="<%=descriptions%>">
<!--#include file="inc/phoneewmjs.asp"-->
<%'headend���д��빩��װ�������ɾ��%>
</head>
<body>
<!--#include file="inc/head.asp"-->
<div class="container trblock " id="trblock178888">
<div class="col-lg-9 <%=colbs9%> trrow1199 trmarb1199 trshow trovh " id="trleft178888">
<div class="trshowtitle1 trovh"><span class="trcrumbs">��ǰλ�ã� <a href="<%=funpath%>">��վ��ҳ</a> &gt;
<%call tr_crumbs(columnid,20)%>
&gt; ���� </span><span class="trmcrumbs">��ǰλ�ã� <a href="<%=outputurl(columnid, "", "", ishtml)%>"><%=left(owncolname,10)%></a> &gt; ����</span></div>
<div class="publicnr trmar2">
<%'body1���д��빩��װ�������ɾ��%>
<div class="trcontentbox " id="trcontentbox178888">
<h1><%=title%></h1>
<p class="trinfo trmar1 clearfix"> <span class="trfl">ʱ�䣺<%=FormatDate(addtime, 2)%> &nbsp;&nbsp; �����</span><span id="clicks" class="trfl"></span> 
<script type="text/javascript" src="<%=funpath%>inc/clicks.asp?id=<%=id%>" ></script> 
<span class="trfl">�� &nbsp;&nbsp; </span><span class="trfl trcomfrom">��Դ��<%=sources%> &nbsp;&nbsp; ���ߣ�<%=author%></span> <span class="trfr"><a href="javascript:void(0)" onClick="setFontsize(0,'trcontenttd')">- С</a> <a href="javascript:void(0)" onClick="setFontsize(1,'trcontenttd')" >+ ��</a></span></p>
<%'body2���д��빩��װ�������ɾ��%>
<div class="trcontent" id="trcontent">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
<td class="trcontenttd" id="trcontenttd"><%If limittype<1 And limitpoints<1 Then
  response.Write tr_arcontent(content, 700, 0, "", pageno)
  limitshow = false
Else
  limitshow = true
End If
%>
<%'body7���д��빩��װ�������ɾ��%>
</td>
</tr>
</table>
<%'body8���д��빩��װ�������ɾ��%>
<%if limitshow then%>
<script id="tdload"  type="text/javascript" src="<%=funpath%>inc/newsshow.asp?id=<%=id%>&funpath=<%=funpath%>"></script>
<%end if%>
<%'body9���д��빩��װ�������ɾ��%>
</div>
<%'body3���д��빩��װ�������ɾ��%>
</div>
<div class="visible-lg-block" style=" width:90%; height:40px; margin:10px auto; text-align:center;" id="trshowshare178888"><script type="text/javascript" src="<%=funpath%>crinc/showshare.asp?l=<%=funpath%>"></script></div>
<%'body10���д��빩��װ�������ɾ��%>
<%if isabout<>1 then%>
<%call tr_nearnews()%>
<%end if%>
</div>
<%'body11���д��빩��װ�������ɾ��%>
<div class="" id="showcomment"> </div>
<script id="dataload"  type="text/javascript" src="<%=funpath%>inc/comment.asp?id=<%=id%>"></script>
<%'body4���д��빩��װ�������ɾ��%>
</div>
<div class="col-lg-3 <%=colbs3%> trrow1199 trlistright" id="trlistright178888">
<%'body5���д��빩��װ�������ɾ��%>
<!--#include file="inc/rightinc.asp"-->
<%'body6���д��빩��װ�������ɾ��%>
</div>
</div>
<%'body12���д��빩��װ�������ɾ��%>
<!--#include file="inc/bottomads.asp"-->
<div class="trpublicline clearfix "> </div>
<!--#include file="inc/foot.asp"-->
<div class="trdisnone" id="trdisnone"></div>
</body>
</html>
<%call CloseConn()%>
