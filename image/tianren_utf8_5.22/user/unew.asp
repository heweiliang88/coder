<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../inc/charsetasp.asp"-->
<!DOCTYPE html>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<!--#include file="../core/Conn.asp"-->
<!--#include file="../core/fun/fun.asp"-->
<!--#include file="../core/fun/core.asp"-->
<!--#include file="../core/class/tr_page.asp"-->
<!--#include file="../core/class/tr_user.asp"-->
<!--#include file="../core/class/tr_article.asp"-->
<!--#include file="../core/md5.asp"-->
<!--#include file="../skin/default/kp171014.asp"-->
<%'include此行代码供安装插件用勿删改%>
<%
Const funpath = "../"
Const dbpath = "../"
Set ctr_page = New tr_page
Set ctr_user = New tr_user
Set ctr_article = New tr_article
safemd5 = trmd5(ctr_user.truserislog(1))
id = tr_killstr(request.querystring("id"), 1, 8, 3, "", "", "", "", 1)
If IsNumeric(id) And id<>0 Then
  Call ctr_article.getarticle(id)
  If ctr_article.isdel = 1 Then Call contrl_message("文章不存在，请编辑其他文章。", funpath&"user/mynews.asp", "", "", "top")
  If safemd5<>ctr_article.contrimd5 Then Call contrl_message("此文章不属于该用户，请编辑其他文章。", funpath&"user/mynews.asp", "", "", "top")
  actionval = "editnew"
Else
  actionval = "addnew"
End If
columnid = ctr_article.columnid
titlecolor = ctr_article.titlecolor
%>
<%'head此行代码供安装插件用勿删改%>
<!--#include file="../inc/headtaginc.asp"-->
<title>会员中心编辑文章_<%=application(siternd & "55trsitetitle")%></title>
<script charset="ut<%%>f-8" src="<%=funpath%>kindeditor_4.1.10/kindeditor-min.js"></script>
<script charset="ut<%%>f-8" src="<%=funpath%>kindeditor_4.1.10/lang/zh_CN.js"></script>
<script>
var editor;
KindEditor.ready(function(K) {
editor = K.create('textarea[name="content"]', {
resizeType : 1,
allowPreviewEmoticons : false,
allowImageUpload : false,
items : [
'fontname', 'fontsize', '|', 'forecolor', 'hilitecolor', 'bold', 'italic', 'underline',
'removeformat', '|', 'justifyleft', 'justifycenter', 'justifyright', 'insertorderedlist',
'insertunorderedlist',  'link']
});
});
</script>
<!--#include file="../inc/phoneewmjs.asp"-->
<%'headend此行代码供安装插件用勿删改%>
</head>
<body>
<!--#include file="../inc/head.asp"-->
<div class="container trblock" id="trblock178888">
<div class=" col-lg-9 <%=colbs9%> trrow1199 trmarb1199 trlist trovh" id="trleft178888">
<div class="trlisttitle1 "><span class="trcrumbs">当前位置： <a href="<%=funpath%>">网站首页</a> &gt;
      会员个人中心 &gt; 编辑文章</span><span class="trmcrumbs">当前位置： <a href="<%=funpath%>">网站首页</a> &gt; 编辑文章</span> </div>
<div class="publicnr trpad1">
<%'body1此行代码供安装插件用勿删改%>
<%if application(siternd & "55tropenarticle")=1 and ctr_user.openarticle=1 then%>
<div class="alert alert-warning alert-dismissible trnotice1" role="alert">
  <button type="button" class="close" data-dismiss="alert" aria-label="Close"><span aria-hidden="true">&times;</span></button>
<%=application(siternd & "55trarticlenotice")%>
</div>

<form name="form1<%=id%>" id="form1<%=id%>" target="ifr1" method="post" onSubmit="return check_uaddnews(this)" action="usersave.asp" class="form-horizontal">
<%'body5此行代码供安装插件用勿删改%>
<div class="form-group">
<label for="title" class="col-sm-2 <%=colbs2%> control-label">标题</label>
<div class="col-sm-10 <%=colbs10%> ">
<input type="text" class="form-control" id="title" name="title" placeholder="标题" value="<%=ctr_article.title%>" >
</div>
</div>
<div class="form-group">
<label for="columnid" class="col-sm-2 <%=colbs2%> control-label">所属栏目</label>
<div class="col-sm-10 <%=colbs10%> ">
<select name="columnid" id="columnid" class="form-control">
<option value="">选择栏目 
<!--#include file="../crinc/column_select.asp"-->
</select>
</div>
</div>
<div class="form-group">
<label for="titlecolor" class="col-sm-2 <%=colbs2%> control-label">标题颜色</label>
<div class="col-sm-10 <%=colbs10%> ">
<select name="titlecolor" id="titlecolor" class="form-control">
<option value="">选择 </option>
<option value="#FF0000" <%if titlecolor="#FF0000" then%>selected<%end if%>>红色 </option>
<option value="#008800" <%if titlecolor="#008800" then%>selected<%end if%>>绿色 </option>
<option value="#0000FF" <%if titlecolor="#0000FF" then%>selected<%end if%>>蓝色 </option>
<option value="#ffff00" <%if titlecolor="#ffff00" then%>selected<%end if%>>黄色 </option>
<option value="#990000" <%if titlecolor="#990000" then%>selected<%end if%>>褐色 </option>
<option value="#660000" <%if titlecolor="#660000" then%>selected<%end if%>>紫色 </option>
<option value="#ffffff" <%if titlecolor="#ffffff" then%>selected<%end if%>>白色 </option>
<option value="#ff00ff" <%if titlecolor="#ff00ff" then%>selected<%end if%>>粉色 </option>
</select>
</div>
</div>
<div class="form-group">
<label for="trmailstr" class="col-sm-2 <%=colbs2%> control-label">投稿人</label>
<div class="col-sm-10 <%=colbs10%> form-control-static" > <%=ctr_user.username%></div>
</div>
<div class="form-group">
<label for="trmailstr" class="col-sm-2 <%=colbs2%> control-label">内容</label>
<div class="col-sm-10 <%=colbs10%> form-control-static" >
<textarea id="content" name="content" style="width:90%;height:300px;">
<%=ctr_article.content%>
</textarea>
</div>
</div>
<div class="form-group">
<label for="keywords" class="col-sm-2 <%=colbs2%> control-label">关键词</label>
<div class="col-sm-10 <%=colbs10%> " >
<textarea name="keywords" class="form-control" id="keywords" rows="3"> <%=ctr_article.keywords%></textarea>
</div>
</div>
<div class="form-group">
<label for="descriptions" class="col-sm-2 <%=colbs2%> control-label">描述</label>
<div class="col-sm-10 <%=colbs10%> " >
<textarea name="descriptions" class="form-control" id="descriptions" rows="3"> <%=ctr_article.descriptions%></textarea>
</div>
</div>
<%'body6此行代码供安装插件用勿删改%>
<div class="trdivbtn">
<button name="submit1" type="submit" id="submit1" class="btn btn-primary trbtn1" onClick="editor.sync()">立即提交</button>
</div>
<input type="hidden" value="<%=actionval%>" name="action" id="action" />
<input type="hidden" value="<%=id%>" name="id" id="id" />
<%'body7此行代码供安装插件用勿删改%>
</form>
<%else%>
<div class="trunotice trfontu">发布文章功能目前已关闭。 </div>
<%end if%>
<%'body2此行代码供安装插件用勿删改%>
</div>
</div>
<div class="col-lg-3 <%=colbs3%> trrow1199 trlistright trovh" id="trlistright178888">
<%'body3此行代码供安装插件用勿删改%>
<!--#include file="../inc/umenu.asp"--> 
<!--#include file="../inc/rightinc.asp"-->
<%'body4此行代码供安装插件用勿删改%>
</div>
</div>
<!--#include file="../inc/bottomads.asp"-->
<div class="trpublicline clearfix "> </div>
<!--#include file="../inc/foot.asp"--> 
<!--#include file="../inc/ifr.asp"-->
</body>
</html>
