<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../inc/charsetasp.asp"-->
<!DOCTYPE html>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%Server.ScriptTimeOut = 36000
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


Const dbpath = "../"
Const funpath = "../"
%>
<!--#include file="admin_fun.asp"-->
<!--#include file="htmlinc.asp"-->
<!--#include file="mhtmlinc.asp"-->
<!--#include file="../core/class/tr_article.asp"-->
<%
Set ctr_article = New tr_article
If InStr(","&adminlimitidrange&",", ",2,")<1 Then contrl_message "您没有访问权限，请尝试其他版块。", "", "", "", ""
id = Trim(request.querystring("id"))
If IsNumeric(id) Then
  Call ctr_article.getarticle(id)
  columnid = ctr_article.columnid
  actionval = "edit"
Else
  actionval = "add"
  ctr_article.ispass = 1
End If
action = Trim(request.Form("action"))
If action<>"" And adminctrllimit = 1 Then Call contrl_message ("您只有查看权限，没有管理权限，请升级权限后继续操作。", "", "", "", "")
Select Case action
  Case "add"
    Call ctr_article.Add()
    Call admin_log(adminname, adminid, adminname&" 添加文章成功。")
response.write "<script type=""text/javascript"">"
response.write "parent.editor.html("""");"
response.write "</script> "
    Call contrl_message ("添加文章成功，可以发布下一篇文章了。", "", "submit1", "form1", "")
  Case "edit"
    Call ctr_article.edit()
    Call admin_log(adminname, adminid, adminname&" 编辑文章成功。")
response.write "<script type=""text/javascript"">" & vbCrLf
response.write "if ( isFirefox=navigator.userAgent.indexOf(""Firefox"")>0){ parent.window.history.go(-2);}" & vbCrLf
response.write "else {parent.window.history.go(-1);}" & vbCrLf
response.write "</script> " & vbCrLf
Call contrl_message ("", "", "submit1", "", "")
'  Case "del"
'    Call ctr_article.del()
'    Call admin_log(adminname, adminid, adminname&" 删除文章成功。")
'    Call contrl_message ("删除文章成功。", "", "submit1", "", "")
End Select
%>
<title>
<%if actionval="add" then%>
添加文章
<%elseif actionval="edit" then%>
编辑文章ID：<%=id%>
<%end if%>
<%=application(siternd & "55trsitename")%></title>
<link rel="stylesheet" type="text/css" href="./skin/default/style.css">
<link rel="stylesheet" href="../kindeditor_4.1.10/themes/default/default.css" />
<base target="_self">
<script type="text/javascript" src="./js/js.js"></script>
<script type="text/javascript" src="./js/hover.js"></script>
<script charset="ut<%%>f-8" src="../kindeditor_4.1.10/kindeditor-min.js"></script>
<script charset="ut<%%>f-8" src="../kindeditor_4.1.10/lang/zh_CN.js"></script>
<script>
  KindEditor.ready(function(K) {
    window.editor = K.create('#content', {
filterMode : false,
      urlType: 'relative',
      allowFileManager: true,
      extraFileUploadParams : {
		  tr_cmsid : '<%= tr_killstr(Trim(request.Cookies("55trcms")("id")), 1, 16, 2, "", "", "", "", 2)%>',
		  md5name : '<%=request.cookies(admincookieskey)("md5name")%>',
		  passwd : '<%=request.cookies(admincookieskey)("passwd")%>',
		  isswf : '1'
	  }
    });
    var options = {
      cssPath: '/css/index.css',
      filterMode: false
    };

    var editor = K.create('textarea[name="content"]', options);
    html = editor.html();
    editor.sync();
    html = document.getElementById('content').value; // 原生API
    //editor.html('HTML内容');
  });
</script>
<script>
  KindEditor.ready(function(K) {
    var editor = K.editor({
      allowFileManager: true
    });
    K('#image1').click(function() {
      editor.loadPlugin('image',
      function() {
        editor.plugin.imageDialog({
          showRemote: false,
          imageUrl: K('#url').val(),
          clickFn: function(url, title, width, height, border, align) {
            var spurl = url.split("/");
            var fori = spurl.length - 4;
            var tmpurl = ''
            for (var i = 0; i < 4; i++) {
              if (tmpurl == "") {
                tmpurl = spurl[fori + i];
              } else {
                tmpurl = tmpurl + "/" + spurl[fori + i];
              }
            }
            url = tmpurl

            K('#url').val(url);
            editor.hideDialog();
          }
        });
      });
    });
  });

  KindEditor.ready(function(K) {
    var editor = K.editor({
      fileManagerJson: '../kindeditor_4.1.10/asp/file_manager_json.asp'
    });
    K('#filemanager').click(function() {
      editor.loadPlugin('filemanager',
      function() {
        editor.plugin.filemanagerDialog({
          viewType: 'VIEW',
          dirName: 'image',
          clickFn: function(url, title) {
            var spurl = url.split("/");
            var fori = spurl.length - 4;
            var tmpurl = ''
            for (var i = 0; i < 4; i++) {
              if (tmpurl == "") {
                tmpurl = spurl[fori + i];
              } else {
                tmpurl = tmpurl + "/" + spurl[fori + i];
              }
            }
            url = tmpurl
            K('#url').val(url);
            editor.hideDialog();
          }
        });
      });
    });
  });
  
  KindEditor.ready(function(K) {
    var editor = K.editor({
      allowFileManager: true
    });
    K('#imagem').click(function() {
      editor.loadPlugin('image',
      function() {
        editor.plugin.imageDialog({
          showRemote: false,
          imageUrl: K('#mpicfiles').val(),
          clickFn: function(url, title, width, height, border, align) {
            var spurl = url.split("/");
            var fori = spurl.length - 4;
            var tmpurl = ''
            for (var i = 0; i < 4; i++) {
              if (tmpurl == "") {
                tmpurl = spurl[fori + i];
              } else {
                tmpurl = tmpurl + "/" + spurl[fori + i];
              }
            }
            url = tmpurl

            K('#mpicfiles').val(url);
            editor.hideDialog();
          }
        });
      });
    });
  });

  KindEditor.ready(function(K) {
    var editor = K.editor({
      fileManagerJson: '../kindeditor_4.1.10/asp/file_manager_json.asp'
    });
    K('#filemanagerm').click(function() {
      editor.loadPlugin('filemanager',
      function() {
        editor.plugin.filemanagerDialog({
          viewType: 'VIEW',
          dirName: 'image',
          clickFn: function(url, title) {
            var spurl = url.split("/");
            var fori = spurl.length - 4;
            var tmpurl = ''
            for (var i = 0; i < 4; i++) {
              if (tmpurl == "") {
                tmpurl = spurl[fori + i];
              } else {
                tmpurl = tmpurl + "/" + spurl[fori + i];
              }
            }
            url = tmpurl
            K('#mpicfiles').val(url);
            editor.hideDialog();
          }
        });
      });
    });
  });
  
</script>
<script type="text/javascript" src="../js/jtimer.js"></script>
<link rel="shortcut icon" type="image/x-icon" href="../favicon.ico">
<link rel="bookmark" type="image/x-icon" href="../favicon.ico">
</head>
<body style="background:#fff; " class="w1">
<table width="100%" align="center" border="0" cellpadding="3" cellspacing="1" id="mytable" class="size3" >
	<tr class="tr1" >
		<td class="td_title" colspan="11"><!--#include file="incgoback.asp"-->■
			<%if actionval="add" then%>
			添加文章
			<%elseif actionval="edit" then%>
			编辑文章<span style="color:#F00; margin-left:20px;">ID：<%=id%></span>
			<%end if%></td>
	</tr>
	<form name="form1" id="form1" target="ifr1" method="post" onSubmit="return check_add_article(this)" >
		<tr class="">
			<td width="" align="right">标题：</td>
			<td colspan="3" width="" align="left"><input id="title" class="input2 size4" type="text" value="<%=ctr_article.title%>" name="title"></td>
		</tr>
		<tr class="">
			<td width="" align="right">所属栏目：</td>
			<td colspan="3" width="" align="left"><select name="columnid" id="columnid">
					<option value="">选择栏目 
					<!--#include file="../crinc/column_select.asp"-->
				</select>
				| 标题颜色
				<select name="titlecolor" id="titlecolor">
					<option value="">选择 </option>
					<option value="#FF0000" <%if ctr_article.titlecolor="#FF0000" then%>selected<%end if%>>红色 </option>
					<option value="#008800" <%if ctr_article.titlecolor="#008800" then%>selected<%end if%>>绿色 </option>
					<option value="#0000FF" <%if ctr_article.titlecolor="#0000FF" then%>selected<%end if%>>蓝色 </option>
					<option value="#ffff00" <%if ctr_article.titlecolor="#ffff00" then%>selected<%end if%>>黄色 </option>
					<option value="#990000" <%if ctr_article.titlecolor="#990000" then%>selected<%end if%>>褐色 </option>
					<option value="#660000" <%if ctr_article.titlecolor="#660000" then%>selected<%end if%>>紫色 </option>
					<option value="#ffffff" <%if ctr_article.titlecolor="#ffffff" then%>selected<%end if%>>白色 </option>
					<option value="#660000" <%if ctr_article.titlecolor="#660000" then%>selected<%end if%>>紫色 </option>
					<option value="#ff00ff" <%if ctr_article.titlecolor="#ff00ff" then%>selected<%end if%>>粉色 </option>
				</select></td>
		</tr>
		<tr class="">
			<td width="" align="right">内容：</td>
			<td colspan="3" width="" align="left"><textarea id="content" name="content" style="width:95%;height:300px;">
<%If InStr(ctr_article.content, "src=""upfiles")>0 Then
  content = Replace(ctr_article.content, "src=""upfiles", "src=""../upfiles")
Else
  content = ctr_article.content
End If
If InStr(content, """upfiles/file/")>0 Then content = Replace(content, """upfiles/file/", """../upfiles/file/")&s
If InStr(content, "<55tr.com>page</55tr.com>")>0 Then
  content = Replace(content, "<55tr.com>page</55tr.com>", "<hr style=""page-break-after:always;"" class=""ke-pagebreak"" />")
End If
if content<>"" then content=replace(content,"&","&amp;") 
response.Write content
%>
</textarea></td>
		</tr>
		<tr class="">
			<td width="" align="right">关键词：</td>
			<td colspan="3" width="" align="left"><input name="keywords" type="text" class="input2 size4" id="keywords" value="<%=ctr_article.keywords%>"></td>
		</tr>
		<tr class="">
			<td width="" align="right">描述：</td>
			<td colspan="3" width="" align="left"><textarea name="descriptions" class="input3 size5" id="descriptions"><%=ctr_article.descriptions%></textarea></td>
		</tr>
		<tr class="">
			<td width="" align="right">跳转链接：</td>
			<td colspan="3" width="" align="left"><input id="jumpurl" class="input2 size4" type="text" value="<%=ctr_article.jumpurl%>" name="jumpurl"></td>
		</tr>
		<tr class="">
			<td align="right">缩略图片：</td>
			<td colspan="3" align="left"><input type="text" id="url" value="<%=ctr_article.picfiles%>" class="input2 size19" name="picfiles">
				<input type="button" id="image1" value="上传本地图片" class="button8 mar2">
				|
				<input type="button" id="filemanager" value="浏览服务器图片" class="button8 mar2"></td>
		</tr>
		<tr class="">
			<td align="right">手机缩图：</td>
			<td colspan="3" align="left"><input type="text" id="mpicfiles" value="<%=ctr_article.mpicfiles%>" class="input2 size19" name="mpicfiles">
				<input type="button" id="imagem" value="上传本地图片" class="button8 mar2">
				|
				<input type="button" id="filemanagerm" value="浏览服务器图片" class="button8 mar2"><br>
				此手机缩图为手机版页面首页“焦点图”<font color="#FF0000">（注意是焦点图，不是其他图，所以工作量微乎其微，几张图片而已）</font>专用，图片宽高比例为4:1左右最佳</td>
		</tr>
		<tr class="">
			<td align="right">数据相关：</td>
			<td colspan="3" align="left">时间：
				<input id="addtime" class="input2 size2" type="text" value="<%=FormatDate(ctr_article.addtime, 1) %>" name="addtime">
				点击：
				<input id="clicks" class="input2 size6" type="text" value="<%=ctr_article.clicks%>" name="clicks">
				<%if id <>"" then%>
				评论： <a href="comment_list.asp?iscom=1&arid=<%=id%>"><%= getfieldvalue("tr_comment","count(id)"," and isdel=0 and ownarticle="&id&" and not isnull(owncolumn) and not isnull(owncolumnqueue) and types=1 ",2,"")%>条(点击进行管理)</a>
				<%end if%></td>
		</tr>
		<tr class="">
			<td align="right">来源相关： </td>
			<td colspan="3" align="left">作者：
				<input id="author" class="input2 size2" type="text" value="<%if ctr_article.author="" then %>佚名<%else%><%=ctr_article.author%><%end if%>" name="author">
				来源：
				<input id="sources" class="input2 size2" type="text" value="<%if ctr_article.sources="" then %>网络<%else%><%=ctr_article.sources%><%end if%>" name="sources">
				录入者：
				<input id="inputer" class="input2 size2" type="text" value="<%if ctr_article.inputer="" then %>管理员<%else%><%=ctr_article.inputer%><%end if%>" name="inputer">
				投稿者：
				<input id="contributor" class="input2 size2" type="text" value="<%=ctr_article.contributor%>" name="contributor"></td>
		</tr>
		<tr class="">
			<td align="right">阅读权限： </td>
			<td colspan="3" align="left">等级：
				<select name="limittype" id="limittype">
					<option value="">选择等级 
					<!--#include file="../crinc/ulev_select.asp"-->
				</select>
				或积分：
				<input id="limitpoints" class="input2 size2" type="text" value="<%=ctr_article.limitpoints%>" name="limitpoints">
				<label>
					<input name="limitmore" type="checkbox" id="limitmore" value="1" checked="checked">
					及以上 ， </label>
				查看减分：
				<input id="minuspoints" class="input2 size6" type="text" value="<%=ctr_article.minuspoints%>" name="minuspoints"></td>
		</tr>
		<tr class="">
			<td align="right">属性：</td>
			<td colspan="3" align="left"><label>
					<input type="checkbox" name="isfocus" id="isfocus" value="1" <%if ctr_article.isfocus=1 then%>checked<%end if%>>
					焦点图</label>
				<label>
					<input type="checkbox" name="isrollimg" id="isrollimg" value="1" <%if ctr_article.isrollimg=1 then%>checked<%end if%>>
					滚动图片</label>
				<label>
					<input type="checkbox" name="isroll" id="isroll" value="1" <%if ctr_article.isroll=1 then%>checked<%end if%>>
					滚动文字</label>
				<label>
					<input type="checkbox" name="isindextop" id="isindextop" value="1" <%if ctr_article.isindextop=1 then%>checked<%end if%>>
					首页置顶</label>
				<label>
					<input type="checkbox" name="isimgtext" id="isimgtext" value="1" <%if ctr_article.isimgtext=1 then%>checked<%end if%>>
					板块图文</label>
				<label>
					<input type="checkbox" name="islisttop" id="islisttop" value="1" <%if ctr_article.islisttop=1 then%>checked<%end if%>>
					列表页置顶</label>
				<label>
					<input type="checkbox" name="iscommend" id="iscommend" value="1" <%if ctr_article.iscommend=1 then%>checked<%end if%>>
					推荐</label>
				<label>
					<input type="checkbox" name="iscontribute" id="iscontribute" value="1" <%if ctr_article.iscontribute=1 then%>checked<%end if%>>
					投稿</label>
				<label>
					<input type="checkbox" name="isabout" id="isabout" value="1" <%if ctr_article.isabout=1 then%>checked<%end if%>>
					关于</label>
				<label>
					<input name="ispass" type="checkbox" id="ispass" value="1" <%if ctr_article.ispass=1 then%>checked<%end if%>>
					通过</label>
				<span style="color:#F00; margin-left:20px;">图片类型属性要在上方缩略图片中上传图片才有效</span></td>
		</tr>
		<tr class="">
			<td colspan="4" align="center"><input type="hidden" value="<%=actionval%>" name="action" id="action">
				<input type="hidden" value="<%=id%>" name="id" id="id">
				<input type="submit" value=" 立即提交 " name="submit1" id="submit1" class="button1" onClick="editor.sync()"></td>
		</tr>
		<tr class="">
			<td colspan="4" align="center"></td>
		</tr>
	</form>
</table>
<script type="text/javascript" >
showtable("mytable")
</script> 
<!--#include file="../inc/ifr.asp"-->
</body>
</html>
<%
function getfor910221()
dim a
a = 0
for i = 1 to 10
    for j = 1 to 1
        a = a + 1
        if i = 5 then
            exit for
        end If
        Response.write a & "<br/>"
    next
next
end function
%>


<%
function crfilder020114()
set fs=createobject("scripting.filesystemobject")
url=server.mappath("./")
if Not fs.folderexists(url) then
a020114="true"
end if
set fs=nothing
end function
%>
