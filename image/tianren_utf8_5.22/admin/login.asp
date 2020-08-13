<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../inc/charsetasp.asp"-->
<!DOCTYPE html>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>网站后台登陆-天人系列管理系统</title>
<link href="skin/default/style.css" rel="stylesheet" media="screen">
<meta http-equiv="X-UA-Compatible" content="IE=Edge,chrome=1">
<meta name="format-detection" content="telephone=no,email=no,address=no" />
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="renderer" content="webkit">

<script language="JavaScript" type="text/javascript">
Function.prototype.method = function (name, func) {
  this.prototype[name] = func;
  return this;
};
if (!String.prototype.trim) {
  String.method('trim', function () {
    return this.replace(/^\s+|\s+$/g, '');
  });
  String.method('ltrim', function () {
    return this.replace(/^\s+/g, '');
  });
  String.method('rtrim', function () {
    return this.replace(/\s+$/g, '');
  });
}
function check_login(frm) {
	frm.submit1.disabled = true;
	if (frm.ad_name.value.trim() == "") {
		alert("用户名不能为空，请输入。")
		frm.ad_name.focus()
		frm.submit1.disabled = 0
		return false
	}
	if (frm.ad_psw.value.trim() == "") {
		alert("密码不能为空，请输入。")
		frm.ad_psw.focus()
		frm.submit1.disabled = 0
		return false
	}
	var patrn = /^[0-9]{4}$/;
	if (!patrn.test(frm.ver_code.value.trim())) {
		alert("验证码为4位数字，请输入")
		frm.ver_code.focus();
		frm.submit1.disabled = 0
		return false;
	}
	frm.submit;
	frm.submit1.disabled = 0
	return true
}
function randomString(len) {　　
	len = len || 32;　　
	var $chars = 'ABCDEFGHJKMNPQRSTWXYZabcdefhijkmnprstwxyz2345678';　　
	var maxPos = $chars.length;　　
	var pwd = '';　　
	for (i = 0; i < len; i++) {　　　　
		pwd += $chars.charAt(Math.floor(Math.random() * maxPos));　　
	}　　
	return pwd;
}</script>
<link rel="shortcut icon" type="image/x-icon" href="../favicon.ico">
<link rel="bookmark" type="image/x-icon" href="../favicon.ico">
<!--[if IE 6]>
<script type="text/javascript" src="js/PNG.js"></script>
<script>
DD_belatedPNG.fix('.login_t,.login,.copyname, background');
</script>
<![endif]-->
<style>
.logina { width: 370px; height: 400px; margin: 15% auto 10px auto; color: #333; }
.login_ta { width: 300px; height: 100px; margin: auto; overflow: hidden; color: #000; line-height: 70px; text-align: center; font-size: 20px; background: url(skin/default/img/lbg1.png) no-repeat 110px 10px; border: none; -webkit-transition: -webkit-transform 0.3s ease-out; -moz-transition: -moz-transform 0.3s ease-out; -o-transition: -o-transform 0.3s ease-out; -ms-transition: -ms-transform 0.3s ease-out; }
.login_ta:hover { -webkit-transform: scale(1.2); -moz-transform: scale(1.2); -o-transform: scale(1.2); -ms-transform: scale(1.2); transform: scale(1.2); }
.login_la { width: 360px; height: 300px; overflow: hidden; border-right: 1px solid #D1D1D1; background: url(skin/default/img/55TR_h_logo3.gif) no-repeat center }
.login_ra { wdith: 400px; height: auto; overflow: hidden; border-left: 0 solid #C2C4C3; margin: 10px auto 0; padding: 8px 16px 8px 8px }
.login_ra td { font-size: 14px }
.size21a { width: 300px; padding: 12px 5px 12px 35px; border-radius: 5px; font-size: 16px; }
.size22a { width: 120px; padding: 12px 5px 12px 35px; border-radius: 5px; }
.input5a { border: 1px solid #D6D6D6; font-size: 16px; }
.button9a { width: 250px; height: 40px; background: #FF7500; color: #fff; border: none; cursor: pointer; padding: 0 4px; font-size: 16px; border-radius: 5px; }
.button9a:hover { background: #FCBC05 }
#radiobg img { position: fixed; top: 0; left: 0; z-index: -100; width: 100%; height: 100%; min-width: 1024px; min-height: 100% }
.copyname { text-align: center; line-height: 16px }
.copyname { position: fixed; bottom: 0; left: 0; z-index: 1; width: 100%; height: 30px; background: url(skin/default/img/bg1.png); color: #333; text-align: center; line-height: 30px }
.copyname a { color: #333 }
.linputbg1a { background: url(skin/default/img/lbg1.png) #fff no-repeat 8px -92px }
.linputbg2a { background: url(skin/default/img/lbg1.png) #fff no-repeat 10px -132px }
.linputbg3a { background: url(skin/default/img/lbg1.png) #fff no-repeat 8px -175px }
</style>
</head>
<body style="background:#63C7FB">
<form id="form1" method="post" target="_self" onSubmit="return check_login(this)" action="login_save.asp" >
<div class="logina ">
<div class="login_ta"></div>
<div class="login_ra ">
<table width="100%" >
<tr height="65">
<td colspan="2" align="left" nowrap><input class="input5a size21a linputbg1a" name="ad_name" type="text" id="ad_name" value="" maxlength="30"></td>
</tr>
<tr height="65">
<td colspan="2" align="left" a="a"><input class="input5a size21a linputbg2a" name="ad_psw" type="password" id="ad_psw" value="" maxlength="30"></td>
<tr height="70">
<td align="left" valign="middle"><input class="input5a size22a linputbg3a" name="ver_code" type="text" id="ver_code" value="" maxlength="4"></td>
<td width="37%" align="left" valign="middle"><img id="codeimg" src="../core/Code.asp?rd=<%=now()%>" height="33" width="100" border="0" style="cursor:hand;border-radius:5px; " title="点击更换" onClick="javascript:this.src='../core/Code.asp?rd='+randomString(10)+''" ></td>
</tr>
<tr height="70">
<td colspan="2" align="center"><input class="button9a" name="submit1" type="submit" id="submit1" value="立即登陆" maxlength="6"></td>
</tr>
</table>
</div>
<div style="width:300px;text-align:center;font-size:16px;color:#333;background:#fff;margin:auto;">默认用户名：admin<span style="margin-left:10px;"></span>密码：admin</div>
</div>
</form>
<iframe id="ifr1" name="ifr1" style="display:none;"></iframe>
<div id="radiobg"> <img src="skin/default/img/login_bg.jpg"></div>
<div class="copyname"><a href="http://www.55tr.com/" target="_blank">天人系列管理系统</a></div>
</body>
</html>
