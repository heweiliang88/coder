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
<%
If InStr(","&adminlimitidrange&",", ",")<1 Then contrl_message "您没有访问权限，请尝试其他版块。", "", "", "", ""

PATH_INFO=trim(Request.ServerVariables("PATH_INFO"))
if instr(PATH_INFO,"/")=1 then
adpath=split(PATH_INFO,"/")(1)
elseif instr(PATH_INFO,"/")>1 then
adpath=split(PATH_INFO,"/")(0)
else
adpath=""
end if
%>
<title>
离线上传安装插件、模板、升级包 <%=application(siternd & "55trsitename")%></title>
<link rel="stylesheet" type="text/css" href="./skin/default/style.css">
<link rel="stylesheet" href="../kindeditor_4.1.10/themes/default/default.css" />
<base target="_self">
<script type="text/javascript" src="./js/js.js"></script>
<script type="text/javascript" src="./js/hover.js"></script>
<script charset="ut<%%>f-8" src="../kindeditor_4.1.10/kindeditor-min.js"></script>
<script charset="ut<%%>f-8" src="../kindeditor_4.1.10/lang/zh_CN.js"></script>
<script>
  KindEditor.ready(function (K) {
    var editor = K.editor({
      allowFileManager: true,
      extraFileUploadParams: {
      isupapp:1,
	  PATH_INFO:'<%=adpath%>'
      }
    });
    K('#app1').click(function () {
      editor.loadPlugin('insertfile', function () {
        editor.plugin.fileDialog({
          fileUrl: K('#inspath').val(),
          clickFn: function (url, title) {
            var spurl = url.split("/");
            var fori = spurl.length - 1;
			//alert (fori);
            var tmpurl = ''
            for (var i = 0; i < 1; i++) {
              if (tmpurl == "") {
                tmpurl = spurl[fori + i];
              } else {
                tmpurl = tmpurl + "/" + spurl[fori + i];
              }
            }
            url = tmpurl
            K('#inspath').val(url);
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

<script>
function checkinspath(frm){
	var inspathstr=frm.inspath.value.trim()
  if (inspathstr == "" || inspathstr.length<10 ) {
//	  alert (inspathstr.length)
    alert("文件上传不能为空或不正确，请先上传从官网应用中心下载的安装包再试。");
    frm.inspath.focus();
    frm.submit1.disabled = 0;
    return false;
  }

	if (typeof(frm.appid)!="undefined"){
  var patrn = /^\+?[1-9][0-9]*$/;
  if (!patrn.test(frm.appid.value.trim())) {
    //		if(frm.orderid.value.trim())
    alert("官网ID为大于0的整数，请输入");
    frm.appid.focus();
    frm.submit1.disabled = 0;
    return false;
  }
	}
  frm.submit();
  frm.submit1.disabled = 0;
  return false;
}

</script>
</head>
<body style="background:#fff; " class="w1">
<table width="100%" align="center" border="0" cellpadding="3" cellspacing="1" id="mytable" class="size3" >
  <tr class="tr1" >
    <td class="td_title" colspan="11"><!--#include file="incgoback.asp"-->■
      离线上传安装插件、模板、升级包</td>
  </tr>
  <form name="form1" id="form1" method="get" onSubmit="return checkinspath(this)" action="uncompress.asp" >
    <tr class="">
      <td align="right">&nbsp;</td>
      <td colspan="3" align="left">&nbsp;</td>
    </tr>
    <tr class="">
      <td align="right">文件上传：</td>
      <td colspan="3" align="left"><input type="text" id="inspath" value="" class="input2 size19" name="inspath">
        <input type="button" id="app1" value="上传本地文件" class="button8 mar2">
</td>
    </tr>
    <tr class="">
      <td align="right" nowrap>该应用对应官网的ID：</td>
      <td colspan="3" align="left"><input type="text" id="appid" value="" class="input2 size12" name="appid">
      <span style="color:#F00;">请务必填写真实的ID，ID查看方法，在应用详情页面，“应用ID”后面的数字就是。ID用于安装时获取一些必要的应用信息，在线安装的情况下不需要填写ID</span></td>
    </tr>
    <tr class="">
      <td colspan="4" align="center"><input type="hidden" value="<%=adpath%>" name="PATH_INFO" id="PATH_INFO">
        <input type="submit" value=" 立即安装 " name="submit1" id="submit1" class="button1" onClick="editor.sync()"></td>
    </tr>
    <tr class="">
      <td colspan="4" align="left">小提示：<br>1、上传安装的功能可以上传安装插件、模板、升级包，上传成功后点击立即安装即可自动进行安装，安装过程中按照提示进行数据备份等即可。<br>
     2、官网应用中心地址是：<a href="http://www.55tr.com/gwapplist.asp" target="_blank">www.55tr.com/gwapplist.asp</a><br>如果您的网站部署在外网的时可以直接在后台左侧“获取插件”“获取模板”功能中自动安装</td>
    </tr>
  </form>
</table>
<script type="text/javascript" >
showtable("mytable")
</script>
<!--#include file="../inc/ifr.asp"-->
</body>
</html>