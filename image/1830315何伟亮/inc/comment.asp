<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../core/Conn.asp"-->
<!--#include file="../core/fun/fun.asp"-->
<!--#include file="../core/class/tr_page.asp"-->
<!--#include file="charsetasp.asp"-->
<%
If application(siternd & "55tropencomment") = 1 Then
  Const dbpath = "../"
  Set ctr_page = New tr_page
  id = tr_killstr(request.querystring("id"), 1, 8, 1, "", "", "", "", 2)
%>
var tmpstr=("<div class=\"trunotice\"> <%=application(siternd & "55trcommentnotice")%> </div>");
tmpstr=tmpstr+("<form id=\"form1\" name=\"form1\" method=\"post\" action=\"<%if ishtml=2 then%>../<%end if%>user/usersave.asp\" target=\"ifr1\" onSubmit=\"return checkguest(this)\">");
tmpstr=tmpstr+("<table class=\"trsendtb\" >");
tmpstr=tmpstr+("<tr>");
tmpstr=tmpstr+("<td align=\"center\" colspan=\"11\"><textarea id=\"content\" name=\"content\" class=\"trgcontent trinput6\"></textarea></td>");
tmpstr=tmpstr+("</tr>");
tmpstr=tmpstr+("<tr>");
tmpstr=tmpstr+("<td align=\"left\" nowrap> 昵称<span class=\"trfont2\"> * </span></td>");
tmpstr=tmpstr+("<td align=\"left\" ><input type=\"text\" value=\"\" name=\"username\" id=\"username\" class=\" trinput3\" /></td>");
tmpstr=tmpstr+("<td align=\"left\" nowrap>email</td>");
tmpstr=tmpstr+("<td align=\"left\" ><input type=\"text\" value=\"\" name=\"mail\" id=\"mail\" class=\" trinput3\" /></td>");
tmpstr=tmpstr+("<td align=\"left\" nowrap> 主页</td>");
tmpstr=tmpstr+("<td align=\"left\" ><input type=\"text\" value=\"\" name=\"homepage\" id=\"homepage\" class=\" trinput3\" /></td>");
tmpstr=tmpstr+("<td align=\"left\" nowrap> 验证码</td>");
tmpstr=tmpstr+("<td align=\"left\" > <input name=\"vcode\" type=\"text\" class=\" trinput7\" id=\"vcode\" value=\"\" maxlength=\"4\" /></td>");
<%if ishtml=2 then%>
tmpstr=tmpstr+("<td align=\"left\"><img id=\"codeimg\" src=\"../core/Code.asp?rd=<%=gen_key(5)%>\" style=\"cursor:hand; \" title=\"点击更换\" onClick=\"javascript:this.src='../core/Code.asp?rd='+randomString(10)+''\" height=\"24\" border=\"0\" width=\"70\"></td>");
<%else%>
tmpstr=tmpstr+("<td align=\"left\"><img id=\"codeimg\" src=\"./core/Code.asp?rd=<%=gen_key(5)%>\" style=\"cursor:hand; \" title=\"点击更换\" onClick=\"javascript:this.src='./core/Code.asp?rd='+randomString(10)+''\" height=\"24\" border=\"0\" width=\"70\"></td>");
<%end if%>
tmpstr=tmpstr+("<td align=\"right\"><input type=\"submit\" value=\"立即提交\" name=\"submit1\" id=\"submit1\" class=\"trbt3 trmar1\" />");
tmpstr=tmpstr+("<input type=\"hidden\" value=\"comment\" name=\"action\" />");
tmpstr=tmpstr+("<input type=\"hidden\" value=\"<%=id%>\" name=\"ownarticle\" />");
tmpstr=tmpstr+("<input type=\"hidden\" value=\"\" name=\"owncolumn\" />");
tmpstr=tmpstr+("<input type=\"hidden\" value=\"\" name=\"owncolumnqueue\" /></td>");
tmpstr=tmpstr+("</tr>");
tmpstr=tmpstr+("</table>");
tmpstr=tmpstr+("</form>");
tmpstr=tmpstr+("<table align=\"center\" border=\"0\" cellpadding=\"3\" cellspacing=\"1\" id=\"mytable\" class=\"trtable2\" >");
<%
pageno = tr_killstr(request.querystring("pageno"), 1, 8, 1, "", "", "", "", 1)
sql = "select id,username,content,addtime,homepage,asuser,answer,astime from tr_comment where types=1 and ownarticle="&id&" and ispass=1 and isdel=0 order by addtime desc , id desc  "
page_size = 20
pagei = 0
n = (pageno -1) * page_size
Call ctr_page.create_rs(sql, rs, conn, page_size, pageno)
If Not rs.EOF Then
  Do While Not rs.EOF
    If pagei>= page_size Then Exit Do
    pagei = pagei + 1
    n = n + 1
%>
tmpstr=tmpstr+("<tr class=\"\">");
tmpstr=tmpstr+("<td width=\"\">");
tmpstr=tmpstr+("<div class=\"trcontents trfl trbookbt\">");
tmpstr=tmpstr+("<p class=\"trp1\"> <span class=\"uname \">");
  <%if isweburl(rs("homepage")) then%>
tmpstr=tmpstr+("<a href=\"#\"><%=tojs2(rs("username"),"")%></a>");
  <%else%>
  <%=tojs2(rs("username"),"tmpstr")%>
  <%end if%>
tmpstr=tmpstr+("</span> 说： <%=tojs2(rs("content"),"")%> </p>")
tmpstr=tmpstr+("<p class=\"trp2\"><%=tojs2(FormatDate(rs("addtime"), 1),"") %> </p>")
<%if len(rs("answer"))>1 then%>
tmpstr=tmpstr+("<p class=\"trp3\">管理员 回复：<%=tojs2(rs("answer"),"")%></p>");
tmpstr=tmpstr+("<p class=\"trp2\"><%=tojs2(FormatDate(rs("astime"), 1),"") %> </p>");
<%end if%>
tmpstr=tmpstr+("</div>");
tmpstr=tmpstr+("</td>");
tmpstr=tmpstr+("</tr>");
<%
rs.movenext
Loop
Else
%>
tmpstr=tmpstr+("<tr>");
tmpstr=tmpstr+("<td>还没有人评论，赶快抢个沙发吧。</td>");
tmpstr=tmpstr+("</tr>");
<%
End If
rs.Close
%>
tmpstr=tmpstr+("</table>");
<%if ctr_page.page_count>1 then%>
<%= tojs2(ctr_page.create_page2("",ctr_page.page_count,pageno,"trpage","dataload"),"tmpstr")%>
<% end if%>
function changeText(tmp){
top.document.getElementById('showcomment').innerHTML = tmp;
top.document.getElementById('showcomment').className  = "trcomment trbor2";
}
changeText(tmpstr)
<%End If
%>
