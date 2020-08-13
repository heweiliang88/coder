<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../inc/charsetasp.asp"-->
<!DOCTYPE html>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<script type="text/javascript" src="./js/js.js"></script>
<%
Const dbpath = "../"
Const funpath = "../"
%>
<!--#include file="admin_fun.asp"-->
<!--#include file="../core/class/tr_site.asp"-->
<%
Set ctr_site = New tr_site
If InStr(","&adminlimitidrange&",", ",1,")<1 Then contrl_message "您没有访问权限，请尝试其他版块。", "", "", "", ""
action = Trim(request.Form("action"))
If action<>"" And adminctrllimit = 1 Then Call contrl_message ("您只有查看权限，没有管理权限，请升级权限后继续操作。", "", "", "", "")
If action = "edit" Then
  Call ctr_site.edit()
  Call admin_log(adminname, adminid, adminname&" 编辑站点设置成功。")
  Call contrl_message ("编辑站点设置成功。", "", "submit1", "", "")
End If
%>
<title>站点设置 <%=application(siternd & "55trsitename")%></title>
<link rel="stylesheet" type="text/css" href="./skin/default/style.css">
<link rel="stylesheet" href="../kindeditor_4.1.10/themes/default/default.css">
<base target="_self">
<script type="text/javascript" src="./js/hover.js"></script>
<script charset="ut<%%>f-8" src="../kindeditor_4.1.10/kindeditor-min.js"></script>
<script charset="ut<%%>f-8" src="../kindeditor_4.1.10/lang/zh_CN.js"></script>
<script>
KindEditor.ready(function(K) {
  var editor = K.editor({
    allowFileManager: true
  });
  //				var options = {
  //        cssPath : '/css/index.css',
  //        filterMode : true
  //};
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
          var tmpurl = '';
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
  //				var options = {
  //        cssPath : '/css/index.css',
  //        filterMode : true
  //};
  K('#imagem').click(function() {
    editor.loadPlugin('image',
    function() {
      editor.plugin.imageDialog({
        showRemote: false,
        imageUrl: K('#mlogopic').val(),
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

          K('#mlogopic').val(url);
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
          var tmpurl = '';
          for (var i = 0; i < 4; i++) {
            if (tmpurl == "") {
              tmpurl = spurl[fori + i];
            } else {
              tmpurl = tmpurl + "/" + spurl[fori + i];
            }

          }
          url = tmpurl
          K('#mlogopic').val(url);
          editor.hideDialog();
        }
      });
    });
  });
});

</script>
<link rel="shortcut icon" type="image/x-icon" href="../favicon.ico">
<link rel="bookmark" type="image/x-icon" href="../favicon.ico">
</head>
<body style="background:#fff; " class="w1">
<form name="form1" id="form1" target="ifr1" method="post" onSubmit="return check_site(this);" >
    <table width="100%" align="center" border="0" cellpadding="3" cellspacing="1" id="mytable" class="size3" >
      <tr class="tr1" >
        <td class="td_title" colspan="5"><!--#include file="incgoback.asp"-->■ 站点基本信息</td>
      </tr>
      <tr>
        <td width="30%" align="right">升级秘钥（请勿外流）：</td>
        <td colspan="3" width="99%"><input name="updatekey" id="updatekey" type="text" class="input2 size4" value="<%=application(siternd & "55trupdatekey")%>"></td>
      </tr>
      <tr>
        <td width="30%" align="right">网站名称（仅用于站长识别网站）：</td>
        <td colspan="3" width="99%"><input name="sitename" id="sitename" type="text" class="input2 size4" value="<%=application(siternd & "55trsitename")%>"></td>
      </tr>
      <tr>
        <td align="right">&nbsp;网站标题：</td>
        <td colspan="3"><input name="sitetitle" id="sitetitle" type="text" class="input2 size4" value="<%=application(siternd & "55trsitetitle")%>"></td>
      </tr>
      <tr>
        <td align="right">&nbsp;网站关键词：</td>
        <td colspan="3"><textarea name="sitekeywords" rows="2" class="input3 size5" id="sitekeywords"><%=CHTMLEncode_fan(application(siternd & "55trsitekeywords"))%></textarea></td>
      </tr>
      <tr>
        <td align="right">&nbsp;网站描述：</td>
        <td colspan="3"><textarea name="sitedescription" rows="2" class="input3 size5" id="sitedescription"><%=CHTMLEncode_fan(application(siternd & "55trsitedescription"))%></textarea></td>
      </tr>
      <tr>
        <td align="right">全局关键词：</td>
        <td colspan="3">全局关键词可以将文章中指定的文字链接至指定的地址，关键词与地址间用竖线"|"间隔，<br>
        	例如：天人文章管理系统|http://www.55tr.com 。多条请换行写<br><textarea name="globalkeywords" rows="2" class="input3 size5" id="globalkeywords"><%=CHTMLEncode_fan(application(siternd & "55trglobalkeywords"))%></textarea></td>
      </tr>
      <tr>
        <td align="right">&nbsp;网站域名：</td>
        <td colspan="3"><input name="homepage" id="homepage" type="text" class="input2 size4" value="<%=application(siternd & "55trhomepage")%>"></td>
      </tr>
      <tr>
        <td align="right">&nbsp;网站LOGO上传：</td>
        <td colspan="3"><input type="text" id="url" name="logopic" value="<%=application(siternd & "55trlogopic")%>" class="input2 size19"> <input type="button" id="image1" value="上传本地图片" class="button8 mar2"> | <input type="button" id="filemanager" value="浏览服务器图片" class="button8 mar2"></td>
      </tr>
      <tr>
        <td align="right">备案号码：</td>
        <td colspan="3"><input name="recordnumber" id="recordnumber" type="text" class="input2 size4" value="<%=application(siternd & "55trrecordnumber")%>"></td>
      </tr>
      <tr>
        <td align="right">网站后台每日免验证码登陆次数：</td>
        <td colspan="3"><input name="limittimes" id="limittimes" type="text" class="input2 size6" value="<%=application(siternd & "55trlimittimes")%>">
          次 （建议不要大于3次）</td>
      </tr>
      <tr>
        <td align="right">开启网站：</td>
        <td colspan="3"><label><input name="opensite" id="opensite" type="checkbox" value="1" <%if application(siternd & "55tropensite")=1 then%>checked="checked"<%end if%>>
        勾选即为开启此功能</label></td>
      </tr>
      <tr>
        <td align="right">&nbsp;网站关闭文字说明：</td>
        <td colspan="3"><input name="closeintro" id="closeintro" type="text" class="input2 size4" value="<%=application(siternd & "55trcloseintro")%>"></td>
      </tr>
      <tr>
        <td align="right">&nbsp;网站统计代码：</td>
        <td colspan="3"><textarea name="statisticscode" rows="2" class="input3 size5" id="statisticscode"><%=CHTMLEncode_fan(application(siternd & "55trstatisticscode"))%></textarea></td>
      </tr>
      <tr>
        <td align="right">&nbsp;<span style="color:#ff0000;">动态、静态、伪静态：</span></td>
        <td colspan="3"><label><input name="urltype" id="urltype" type="radio" value="1" <%if application(siternd & "55trurltype")=1 then%>checked="checked"<%end if%>> 动态版</label> <label><input name="urltype" id="urltype" type="radio" value="2" <%if application(siternd & "55trurltype")=2 then%>checked="checked"<%end if%>>
        静态版(<span style="color:#ff0000;">开启后需生成静态</span>)
</label> <label><input name="urltype" id="urltype" type="radio" value="3" <%if application(siternd & "55trurltype")=3 then%>checked="checked"<%end if%>>
        伪静态版(需服务器伪静态组件支持)</label></td>
      </tr>
      <tr>
        <td align="right">自定义静态页面存储主路径：</td>
        <td colspan="3"><input name="htmlurl" id="htmlurl" type="text" class="input2 size19" value="<%=application(siternd & "55trhtmlurl")%>"> 
        就是页面存储的主文件夹，不要含有斜杠/</td>
      </tr>
      <tr class="tr1" style="display:none;">
        <td class="td_title" colspan="5">■ 手机版设置</td>
      </tr>
      <tr style="display:none;">
        <td align="right">手机版导航栏样式</td>
        <td colspan="3"><label><input name="mnavtype" id="mnavtype" type="radio" value="1" <%if application(siternd & "55trmnavtype")=1 then%>checked="checked"<%end if%>> 
        单行左右滑动</label> <label><input name="mnavtype" id="mnavtype" type="radio" value="2" <%if application(siternd & "55trmnavtype")=2 then%>checked="checked"<%end if%>>
        多行触摸隐/显 （更换完导航栏样式后，若手机版为静态版需要整站生成静态方能生效）</label></td>
      </tr>
       <tr style="display:none;">
        <td align="right">&nbsp;<span style="color:#ff0000;">手机版动态、静态、伪静态：</span></td>
        <td colspan="3"><label><input name="murltype" id="murltype" type="radio" value="1" <%if application(siternd & "55trmurltype")=1 then%>checked="checked"<%end if%>> 动态版</label> <label><input name="murltype" id="murltype" type="radio" value="2" <%if application(siternd & "55trmurltype")=2 then%>checked="checked"<%end if%>>
        静态版(<span style="color:#ff0000;">开启后需生成静态</span>)
</label> <label><input name="murltype" id="murltype" type="radio" value="3" <%if application(siternd & "55trmurltype")=3 then%>checked="checked"<%end if%>>
        伪静态版(需服务器伪静态组件支持)</label></td>
      </tr>
      <tr style="display:none;">
        <td align="right">手机版自定义静态页面存储主路径：</td>
        <td colspan="3"><input name="mhtmlurl" id="mhtmlurl" type="text" class="input2 size19" value="<%=application(siternd & "55trmhtmlurl")%>"> 
        就是页面存储的主文件夹，不要含有斜杠/</td>
      </tr>
      <tr style="display:none;">
        <td align="right">&nbsp;手机版页面网站LOGO上传：</td>
        <td colspan="3"><input type="text" id="mlogopic" name="mlogopic" value="<%=application(siternd & "55trmlogopic")%>" class="input2 size19"> <input type="button" id="imagem" value="上传本地图片" class="button8 mar2"> | <input type="button" id="filemanagerm" value="浏览服务器图片" class="button8 mar2"><br>          
          手机版页面的logo需要宽度、高度比例在5:1以上才协调，14:1最佳</td>
      </tr>
      
      
      <tr class="tr1" >
        <td class="td_title" colspan="5">■ 联系方式设置</td>
      </tr>
      <tr>
        <td align="right">公司/单位/团体名称：</td>
        <td colspan="3"><input name="companyname" id="companyname" type="text" class="input2 size4" value="<%=application(siternd & "55trcompanyname")%>"></td>
      </tr>
      <tr>
        <td align="right">联系地址：</td>
        <td colspan="3"><input name="address" id="address" type="text" class="input2 size4" value="<%=application(siternd & "55traddress")%>"></td>
      </tr>
      <tr>
        <td align="right">邮政编码：</td>
        <td colspan="3"><input name="zip" id="zip" type="text" class="input2 size4" value="<%=application(siternd & "55trzip")%>"></td>
      </tr>
      <tr>
        <td align="right">联系电话：</td>
        <td colspan="3"><input name="tel" id="tel" type="text" class="input2 size4" value="<%=application(siternd & "55trtel")%>"></td>
      </tr>
      <tr>
        <td align="right">传真号码：</td>
        <td colspan="3"><input name="fax" id="fax" type="text" class="input2 size4" value="<%=application(siternd & "55trfax")%>"></td>
      </tr>
      <tr>
        <td align="right">QQ号码：</td>
        <td colspan="3"><input name="qq" id="qq" type="text" class="input2 size4" value="<%=application(siternd & "55trqq")%>"></td>
      </tr>
      <tr>
        <td align="right">email邮箱：</td>
        <td colspan="3"><input name="mail" id="mail" type="text" class="input2 size4" value="<%=application(siternd & "55trmail")%>"></td>
      </tr>
      <tr class="tr1" >
        <td class="td_title" colspan="5">■ 会员注册设置</td>
      </tr>
      <tr>
        <td align="right">注册会员需审核：</td>
        <td colspan="3"><label><input name="opencheckuser" id="opencheckuser" type="checkbox" value="1" <%if application(siternd & "55tropencheckuser")=1 then%>checked="checked"<%end if%>>
        勾选即为开启此功能</label></td>
      </tr>
      <tr>
        <td align="right">开启会员注册功能：</td>
        <td colspan="3"><label><input name="openuserreg" id="openuserreg" type="checkbox" value="1" <%if application(siternd & "55tropenuserreg")=1 then%>checked="checked"<%end if%>>
        勾选即为开启此功能</label></td>
      </tr>
      <tr>
        <td align="right">禁止注册的会员名：<br>请用英文逗号 , 隔开</td>
        <td colspan="3"><textarea name="forbidusername" rows="2" class="input3 size5" id="forbidusername"><%=CHTMLEncode_fan(application(siternd & "55trforbidusername"))%></textarea></td>
      </tr>
      <tr>
        <td align="right">会员注册页面显示的提示信息：</td>
        <td colspan="3"><textarea name="regnotice" rows="2" class="input3 size5" id="regnotice"><%=CHTMLEncode_fan(application(siternd & "55trregnotice"))%></textarea></td>
      </tr>
      <tr>
        <td align="right">注册会员初始积分：</td>
        <td colspan="3"><input name="initialpoints" id="initialpoints" type="text" class="input2 size6" value="<%=application(siternd & "55trinitialpoints")%>">
          分</td>
      </tr>
      <tr>
        <td align="right">每日签到赠送积分数：</td>
        <td colspan="3"><input name="signpoints" id="signpoints" type="text" class="input2 size6" value="<%=application(siternd & "55trsignpoints")%>">
        分</td>
      </tr>
      <tr>
        <td align="right">每日登陆赠送积分数：</td>
        <td colspan="3"><input name="loginpoints" id="loginpoints" type="text" class="input2 size6" value="<%=application(siternd & "55trloginpoints")%>">
        分</td>
      </tr>
      <tr class="tr1" >
        <td class="td_title" colspan="5">■ 留言发文功能设置</td>
      </tr>
      <tr>
        <td align="right">开启留言功能：</td>
        <td colspan="3"><label><input name="openguestbook" id="openguestbook" type="checkbox" value="1" <%if application(siternd & "55tropenguestbook")=1 then%>checked="checked"<%end if%>>
        勾选即为开启此功能</label></td>
      </tr>
      <tr>
        <td align="right">开启评论功能：</td>
        <td colspan="3"><label><input name="opencomment" id="opencomment" type="checkbox" value="1" <%if application(siternd & "55tropencomment")=1 then%>checked="checked"<%end if%>>
        勾选即为开启此功能</label></td>
      </tr>
      <tr>
        <td align="right">开启会员发布文章功能：</td>
        <td colspan="3"><label><input name="openarticle" id="openarticle" type="checkbox" value="1" <%if application(siternd & "55tropenarticle")=1 then%>checked="checked"<%end if%>>
        勾选即为开启此功能</label></td>
      </tr>
      <tr>
        <td align="right">开启留言/评论审核功能：</td>
        <td colspan="3"><label><input name="opencheckcomment" id="opencheckcomment" type="checkbox" value="1" <%if application(siternd & "55tropencheckcomment")=1 then%>checked="checked"<%end if%>>
        勾选即为开启此功能</label></td>
      </tr>
      <tr>
        <td align="right">开启会员发布文章审核功能：</td>
        <td colspan="3"><label><input name="opencheckarticle" id="opencheckarticle" type="checkbox" value="1" <%if application(siternd & "55tropencheckarticle")=1 then%>checked="checked"<%end if%>>
        勾选即为开启此功能</label></td>
      </tr>
      <tr>
        <td align="right">发布文章赠送积分数：</td>
        <td colspan="3"><input name="publisharticlepoints" id="publisharticlepoints" type="text" class="input2 size6" value="<%=application(siternd & "55trpublisharticlepoints")%>">
分</td>
      </tr>
      <tr>
        <td align="right">留言页面提示信息：</td>
        <td colspan="3"><textarea name="guestnotice" rows="2" class="input3 size5" id="guestnotice"><%=CHTMLEncode_fan(application(siternd & "55trguestnotice"))%></textarea></td>
      </tr>
      <tr>
        <td align="right">评论页面提示信息：</td>
        <td colspan="3"><textarea name="commentnotice" rows="2" class="input3 size5" id="commentnotice"><%=CHTMLEncode_fan(application(siternd & "55trcommentnotice"))%></textarea></td>
      </tr>
      <tr>
        <td align="right">会员发布文章页面提示信息：</td>
        <td colspan="3"><textarea name="articlenotice" rows="2" class="input3 size5" id="articlenotice"><%=CHTMLEncode_fan(application(siternd & "55trarticlenotice"))%></textarea></td>
      </tr>
      <tr>
        <td align="right">&nbsp;</td>
        <td colspan="3"><input type="hidden" value="edit" name="action" id="action"><input type="submit" value=" 提交修改 " name="submit1" id="submit1" class="button1"></td>
      </tr>
      <tr>
        <td align="right">&nbsp;</td>
        <td colspan="3">&nbsp;</td>
      </tr>
</table>
</form>
<script type="text/javascript" >
showtable("mytable")
</script>
<!--#include file="../inc/ifr.asp"-->
</body>
</html>
<%'勿删改此行%>
