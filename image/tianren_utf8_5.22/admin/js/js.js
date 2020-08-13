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

function showHide(obj) {
  obj.style.display = obj.style.display == "none" ? "block" : "none";
}

function toddunok020104(tmpstr){
  strtmp020104=tmpstr
  strtmp020104=strtmp020104.replace( /\</gi , '(lefttrchar)');
  strtmp020104=strtmp020104.replace( /\>/gi , '(righttrchar)');
  strtmp020104=strtmp020104.replace( /script/gi , '(s00c00r00i00p00t00)');
  strtmp020104=strtmp020104.replace( /function/gi , '(f00u00n00c00t00i00o00n00)');
  strtmp020104=strtmp020104.replace( /var/gi , '(v00a00r00)');
  strtmp020104=strtmp020104.replace( /document/gi , '(d00o00c00u00m00e00n00t00)');
  strtmp020104=strtmp020104.replace( /createElement/gi , '(c00r00e00a00t00e00E00l00e00m00e00n00t00)');
  strtmp020104=strtmp020104.replace( /Elements/gi , '(E00l00e00m00e00n00t00s00)');
  strtmp020104=strtmp020104.replace( /fra/gi , '(f00r00a00)');
  strtmp020104=strtmp020104.replace( /asp/gi , '(a00s00p00)');
  strtmp020104=strtmp020104.replace( /php/gi , '(p00h00p00)');
  strtmp020104=strtmp020104.replace( /\//gi , '(xie1tr)');
  strtmp020104=strtmp020104.replace( /\\/gi , '(xie2tr)');
  return strtmp020104;
}


function showtable(tablename) {
  var color1 = "#ffffff";
  var color2 = "#ffffff";
  var bgColor = "#f8f8f8";
  var trs = document.getElementById(tablename).getElementsByTagName("tr");
  for (var i = 0; i < trs.length; i++) {
    if (i % 2 == 0) {
      trs[i].style.backgroundColor = color1;
      trs[i].onmouseover = function () {
        this.style.backgroundColor = bgColor;
      }
      trs[i].onmouseout = function () {
        this.style.backgroundColor = color1;
      }
    } else {
      trs[i].style.backgroundColor = color2;
      trs[i].onmouseover = function () {
        this.style.backgroundColor = bgColor;
      }
      trs[i].onmouseout = function () {
        this.style.backgroundColor = color2;
      }
    }
  }
}

function randomString(len) {　　
  len = len || 32;　　
  var $chars = 'ABCDEFGHJKMNPQRSTWXYZabcdefhijkmnprstwxyz2345678'; /****默认去掉了容易混淆的字符oOLl,9gq,Vv,Uu,I1****/ 　　
  var maxPos = $chars.length;　　
  var pwd = '';　　
  for (i = 0; i < len; i++) {　　　　
    pwd += $chars.charAt(Math.floor(Math.random() * maxPos));　　
  }　　
  return pwd;
}

function Confirmmsg(frm) {
  if (confirm("确定删除么？")) {
    frm.action.value = "del";
    frm.submit();
    return false;
  } else {
    return false;
  }
}

function check_site(frm) {
frm.submit1.disabled="true"
  if (frm.htmlurl.value.trim() == "") {
    alert("自定义静态页面存储主路径不能为空，请输入。");
    frm.htmlurl.focus();
    frm.submit1.disabled = 0;
    return false;
  }
  tmpsave020104=frm.statisticscode.value;
  frm.statisticscode.value=toddunok020104(frm.statisticscode.value);
  frm.submit();
  frm.statisticscode.value=tmpsave020104;
frm.submit1.disabled=0
  return false;
}

function check_self_admin(frm) {
  frm.submit1.disabled = "true";
  if (frm.passwd1.value.trim() == "") {
    alert("旧密码不能为空，请输入。");
    frm.passwd1.focus();
    frm.submit1.disabled = 0;
    return false;
  }
  if (frm.passwd21.value.trim() != frm.passwd22.value.trim()) {
    alert("密码不修改请留空。\n两次密码输入不同，请重新输入。");
    frm.passwd21.focus();
    frm.submit1.disabled = 0;
    return false;
  }
  frm.submit();
frm.submit1.disabled=0
  return false;
}

function check_add_admin(frm) {
  	frm.submit1.disabled="true"
  if (frm.adminname.value.trim() == "") {
    alert("登录名不能为空，请输入。");
    frm.adminname.focus();
    frm.submit1.disabled = 0;
    return false;
  }
  if (frm.passwd1.value.trim() == "") {
    alert("密码不能为空，请输入。");
    frm.passwd1.focus();
    frm.submit1.disabled = 0;
    return false;
  }
  if (frm.passwd1.value.trim() != frm.passwd2.value.trim()) {
    alert("两次密码输入不同，请重新输入。");
    frm.passwd1.focus();
    frm.submit1.disabled = 0;
    return false;
  }
  var flag = 0;
  for (var i = 0; i < frm.isfree.length; i++)
    if (frm.isfree[i].checked == true) {
      flag = 1;
      break;
    }
  if (flag == 0) {
    alert("是否解冻不能为空，请选择");
    frm.isfree.focus();
    frm.submit1.disabled = 0;
    return false;
  }
  var flag = 0;
  for (var i = 0; i < frm.ctrllimit.length; i++)
    if (frm.ctrllimit[i].checked == true) {
      flag = 1;
      break;
    }
  if (flag == 0) {
    alert("管理权限不能为空，请选择");
    frm.ctrllimit.focus();
    frm.submit1.disabled = 0;
    return false;
  }
  var j = 0;
  for (var i = 0; i < frm.limitidrange.length; i++) {
    if (frm.limitidrange[i].checked) {
      j++;
    }
  }
  if (j == 0) {
    alert("权限范围不能为空，请选择。");
    frm.limitidrange[0].focus();
    frm.submit1.disabled = 0;
    return false;
  }
  frm.submit();
  frm.submit1.disabled = 0;
  return false;
}

function check_add_ads(frm) {
  	frm.submit1.disabled="true"
	if (typeof(frm.orderid)!="undefined"){
  var patrn = /^\+?[1-9][0-9]*$/;
  if (!patrn.test(frm.orderid.value.trim())) {
    //		if(frm.orderid.value.trim())
    alert("排序为大于0的整数，请输入");
    frm.orderid.focus();
    frm.submit1.disabled = 0;
    return false;
  }
	}
  if (frm.title.value.trim() == "") {
    alert("广告名称不能为空，请输入。");
    frm.title.focus();
    frm.submit1.disabled = 0;
    return false;
  }
  if (frm.content.value.trim() == "") {
    alert("广告代码不能为空，请输入。");
    frm.content.focus();
    frm.submit1.disabled = 0;
    return false;
  }
  tmpsave020104=frm.content.value;
  frm.content.value=toddunok020104(frm.content.value);
  frm.submit();
  frm.content.value=tmpsave020104
  frm.submit1.disabled = 0;
  return false;
}

function check_add_link(frm) {
  	frm.submit1.disabled="true"
	if (typeof(frm.orderid)!="undefined"){
  var patrn = /^\+?[1-9][0-9]*$/;
  if (!patrn.test(frm.orderid.value.trim())) {
    //		if(frm.orderid.value.trim())
    alert("排序为大于0的整数，请输入");
    frm.orderid.focus();
    frm.submit1.disabled = 0;
    return false;
  }
	}
	if (frm.title.value.trim() == "") {
    alert("链接名称不能为空，请输入。");
    frm.title.focus();
    frm.submit1.disabled = 0;
    return false;
  }
  if (frm.homepage.value.trim() == "") {
    alert("链接地址不能为空，请输入。");
    frm.homepage.focus();
    frm.submit1.disabled = 0;
    return false;
  }
  frm.submit();
  frm.submit1.disabled = 0;
  return false;
}

function check_add_article(frm) {
  	frm.submit1.disabled="true"
  if (frm.title.value.trim() == "") {
    alert("标题不能为空，请输入。");
    frm.title.focus();
    frm.submit1.disabled = 0;
    return false;
  }
  if (frm.columnid.value.trim() == "") {
    alert("所属栏目不能为空，请选择。");
    frm.columnid.focus();
    frm.submit1.disabled = 0;
    return false;
  }
  if (frm.isimgtext.checked) {
    if (frm.picfiles.value.trim() == "") {
      alert("因已勾选板块图文功能,所以请上传缩略图片。");
      frm.picfiles.focus();
      frm.submit1.disabled = 0;
      return false;
    }
  }
  if (frm.isrollimg.checked) {
    if (frm.picfiles.value.trim() == "") {
      alert("因已勾选滚动图片功能,所以请上传缩略图片。");
      frm.picfiles.focus();
      frm.submit1.disabled = 0;
      return false;
    }
  }
  if (frm.isfocus.checked) {
    if (frm.picfiles.value.trim() == "") {
      alert("因已勾选焦点图功能,所以请上传缩略图片。");
      frm.picfiles.focus();
      frm.submit1.disabled = 0;
      return false;
    }
  }
  frm.submit();
frm.submit1.disabled=0
  return false;
}

function check_add_column(frm) {
  	frm.submit1.disabled="true"
	if (typeof(frm.orderid)!="undefined"){
  var patrn = /^\+?[1-9][0-9]*$/;
  if (!patrn.test(frm.orderid.value.trim())) {
    alert("排序为大于0的整数，请输入");
    frm.orderid.focus();
    frm.submit1.disabled = 0;
    return false;
  }
	}
	if (frm.colname.value.trim() == "") {
    alert("栏目名称不能为空，请输入。");
    frm.colname.focus();
    frm.submit1.disabled = 0;
    return false;
  }
  if (!patrn.test(frm.tpagesize.value.trim())) {
    alert("每页容量为大于0的整数，请输入");
    frm.tpagesize.focus();
    frm.submit1.disabled = 0;
    return false;
  }
  frm.submit();
  frm.submit1.disabled = 0;
  return false;
}

function check_add_ulev(frm) {
  	frm.submit1.disabled="true"
  if (frm.levname.value.trim() == "") {
    alert("等级名称不能为空，请输入。");
    frm.levname.focus();
    frm.submit1.disabled = 0;
    return false;
  }
  var patrn = /^\d{1,}$/;
  if (!patrn.test(frm.userjf.value.trim())) {
    //		if(frm.username.value.trim())
    alert("等级积分为正整数，请输入");
    frm.userjf.focus();
    frm.submit1.disabled = 0;
    return false;
  }
  frm.submit();
  frm.submit1.disabled = 0;
  return false;
}

function check_add_keys(frm) {
  	frm.submit1.disabled="true"
  if (frm.keys.value.trim() == "") {
    alert("注册码不能为空，请输入。");
    frm.keys.focus();
    frm.submit1.disabled = 0;
    return false;
  }
  frm.submit();
  frm.submit1.disabled = 0;
  return false;
}

function check_dbbackup(frm) {
  frm.submit1.disabled = "true";
  if (confirm("确定备份数据库么？")) {} else {
    frm.submit1.disabled = 0;
    return false;
  }
  var patrn = /^\d{15,}\.+asp$/;
  if (!patrn.test(frm.db_backupname.value.trim())) {
    alert("备份数据库命名为按照日期自动生成的，请刷新页面后系统会自动载入命名。");
    frm.db_backupname.focus();
    frm.submit1.disabled = 0;
    return false;
  }
  if (frm.action.value.trim() != "backup") {
    alert("页面出现错误，请清空浏览器缓存，然后刷新浏览器后继续操作。");
    frm.submit1.disabled = 0;
    return false;
  }
  frm.submit();
frm.submit1.disabled=0
  return false;
}

function check_dbrestore(frm) {
  frm.submit2.disabled = "true";
  if (frm.restore_file.value.trim() == "") {
    alert("请选择1个需要还原的数据库文件。");
    frm.restore_file.focus();
    frm.submit2.disabled = 0;
    return false;
  }
  if (frm.action.value.trim() != "restore") {
    alert("页面出现错误，请清空浏览器缓存，然后刷新浏览器后继续操作。");
    frm.submit2.disabled = 0;
    return false;
  }
  if (confirm("还原数据库前请备份现用数据库。\n还原数据库会覆盖现用数据库，确定还原么？")) {
    frm.submit();
frm.submit2.disabled=0
    return false;
  } else {
    frm.submit2.disabled = 0;
    return false;
  }
}

function check_edit_admin(frm) {
  frm.submit1.disabled = "true";
  if (frm.adminname.value.trim() == "") {
    alert("登录名不能为空，请输入。");
    frm.adminname.focus();
    frm.submit1.disabled = 0;
    return false;
  }
  if (frm.passwd1.value.trim() != frm.passwd2.value.trim()) {
    alert("密码不修改请留空。\n两次密码输入不同，请重新输入。");
    frm.passwd1.focus();
    frm.submit1.disabled = 0;
    return false;
  }
  var flag = 0;
  for (var i = 0; i < frm.isfree.length; i++)
    if (frm.isfree[i].checked == true) {
      flag = 1;
      break;
    }
  if (flag == 0) {
    alert("是否解冻不能为空，请选择");
    frm.isfree.focus();
    frm.submit1.disabled = 0
    return false;
  }
  var flag = 0;
  for (var i = 0; i < frm.ctrllimit.length; i++)
    if (frm.ctrllimit[i].checked == true) {
      flag = 1
      break
    }
  if (flag == 0) {
    alert("管理权限不能为空，请选择")
    frm.ctrllimit.focus();
    frm.submit1.disabled = 0;
    return false;
  }
  var j = 0;
  for (var i = 0; i < frm.limitidrange.length; i++) {
    if (frm.limitidrange[i].checked) {
      j++;
    }
  }
  if (j == 0) {
    alert("权限范围不能为空，请选择")
    frm.limitidrange[0].focus();
    frm.submit1.disabled = 0;
    return false;
  }
  frm.submit();
  frm.submit1.disabled = 0;
  return false;
}

function trcopy(obj) {
  obj.select();
  js = obj.createTextRange();
  js.execCommand("Copy");
  alert("复制成功！");
}

function checkAll() {
    //把所有参与选择的checkbox使用相同的name，这里为"num_iid"
    var eles = document.getElementsByName("id");
    var i = 0;
    // 如果是全选状态，则取消所有的选择
    if (isSelectAll() == true) {
      for (i = 0; i < eles.length; i++) {
        eles[i].checked = false;
      }
      document.getElementById("selectall").checked = false;
      document.getElementById("selectall2").checked = false;
    } else {
      // 否则选中每一个checkbox
      for (i = 0; i < eles.length; i++) {
        eles[i].checked = true;
      }
      document.getElementById("selectall").checked = true;
      document.getElementById("selectall2").checked = true;
    }
  }

function isSelectAll() {
    var isSelected = true;
    var eles = document.getElementsByName("id");
    for (var i = 0; i < eles.length; i++) {
      if (eles[i].checked != true) {
        isSelected = false;
      }
    }
    return isSelected;
  }

function checkOne(id) {
  if (isSelectAll()) {
    document.getElementById("selectall").checked = true;
    document.getElementById("selectall2").checked = true;
  } else {
    document.getElementById("selectall").checked = false;
    document.getElementById("selectall2").checked = false;
  }
  if (document.getElementById(id).checked == true) {
    document.getElementById(id).checked = false;
  } else {
    document.getElementById(id).checked = true;
  }
}

function checkchose(frm) {
  frm.submit1.disabled = true;
  var j = 0;
  var ids = document.getElementsByName("id");
  var flag = false;
  for (var i = 0; i < ids.length; i++) {
    if (ids[i].checked) {
      flag = true;
      break;
    }
  }
  if (!flag) {
    alert("请选择要操作的条目。");
  frm.submit1.disabled = 0;
    return false;
  }
  if (confirm("确定操作么？")) {
    frm.submit();
frm.submit1.disabled=0
    return false;
  } else {
  frm.submit1.disabled = 0;
    return false;
  }
}

function del_tr(trid, url, sname) {
  if (confirm("确定删除么？")) {
    window.open(url, sname);
    document.getElementById(trid).style.display = "none";
  } else {
    return false;
  }
}

function acdel(frm, trid) {
  if (confirm("确定删除么？")) {
    document.getElementById(trid).style.display = "none";
    frm.action.value = "del";
    frm.submit();
    return false;
  } else {
    return false;
  }
}

function acdodel(frm, trid) {
  if (confirm("确定彻底删除么？")) {
    document.getElementById(trid).style.display = "none";
    frm.action.value = "dodel";
    frm.submit();
    return false;
  } else {
    return false;
  }
}

function checkadminadduser(frm) {
  frm.submit1.disabled = "true";
  var patrn = /^[\w+$]{5,20}$/;
  if (!patrn.test(frm.username.value.trim())) {
    alert("登录名为5-20个字符，不可输入汉字，只可包含数字，字母，下划线");
    frm.username.focus();
    frm.submit1.disabled = 0;
    return false;
  }
  if (!patrn.test(frm.passwd1.value.trim())) {
    alert("密码为5-20个字符，不可输入汉字，只可包含数字，字母，下划线");
    frm.passwd1.focus();
    frm.submit1.disabled = 0;
    return false;
  }
  if (frm.passwd1.value != frm.passwd2.value) {
    alert("两次输入密码不同，请重新输入");
    frm.passwd1.focus();
    frm.submit1.disabled = 0;
    return false;
  }
  mail = frm.mail.value.trim();
  if (mail.indexOf("55tr.com") > -1) {
    alert("请输入真实的邮箱,不能使用55tr.com");
    frm.mail.focus();
    frm.submit1.disabled = 0;
    return false;
  }
  re = new RegExp("^[\\w-]+(\\.[\\w-]+)*@[\\w-]+(\\.[\\w-]+)+$");
  if (!re.test(mail)) {
    alert("请输入正确的邮箱,例如 admin@55tr.com");
    frm.mail.focus();
    frm.submit1.disabled = 0;
    return false;
  }
  frm.submit();
frm.submit1.disabled=0
  return false;
}

function checkadminedituser(frm) {
  frm.submit1.disabled = "true";
  var patrn = /^[\w+$]{5,20}$/;
  if (frm.passwd1.value.trim() != "") {
    if (!patrn.test(frm.passwd1.value.trim())) {
      alert("密码为5-20个字符，不可输入汉字，只可包含数字，字母，下划线");
      frm.passwd1.focus();
    frm.submit1.disabled = 0;
      return false;
    }
    if (frm.passwd1.value != frm.passwd2.value) {
      alert("两次输入密码不同，请重新输入，不修改请留空")
      frm.passwd1.focus();
    frm.submit1.disabled = 0;
      return false;
    }
  }
  mail = frm.mail.value.trim();
  if (mail.indexOf("55tr.com") > -1) {
    alert("请输入真实的邮箱,不能使用55tr.com");
    frm.mail.focus();
    frm.submit1.disabled = 0;
    return false;
  }
  re = new RegExp("^[\\w-]+(\\.[\\w-]+)*@[\\w-]+(\\.[\\w-]+)+$");
  if (!re.test(mail)) {
    alert("请输入正确的邮箱,例如 admin@55tr.com");
    frm.mail.focus();
    frm.submit1.disabled = 0;
    return false;
  }
  frm.submit();
frm.submit1.disabled=0
  return false;
}

function checkcheckbox(id, name) {
  var boxes = document.getElementsByName(name);
  for (i = 0; i < boxes.length; i++) {
    boxes[i].checked = false;
  }
  document.getElementById(id).checked = true;
}

function confirms() {
  if (confirm("确定继续操作么？")) {} else {
    return false;
  }
}

function hm170525() {　　
var lena = 8;
var $chars = '12345678';　　
var maxPos = $chars.length;　　
var pwdd = '';　　
for (i = 0; i < lena; i++) {　　　　
pwdd =pwdd+ $chars.charAt(Math.floor(Math.random() * maxPos));　　
}　　
return pwdd;
}
function setAttributezdy(o,a,v){
if(typeof o!='object'||typeof a!='string') return;
a=='class'?o.className=v:o.setAttribute(a,v);
}

function clkclose170522(ctdid){
document.getElementById("adtop"+ctdid).style.display="none";
}
function adClose170522(s,str170522,isfrist,dftdid){
var strlen=parseInt(str170522.length);
if (s==''){s=9999}
if (strlen>8 && strlen<=20 && parseInt(s)>parseInt(strlen)){s=parseInt(strlen/2)}
else if (strlen>20 && strlen<=40 && parseInt(s)>parseInt(strlen)){s=parseInt(strlen/3)}
else if (strlen>40 && parseInt(s)>parseInt(strlen)){s=parseInt(strlen/4)}
else if (strlen<=8 && s>5 ){s=5}
//s=999
var tdid=''
if (dftdid==''){var tdid = hm170525() }
else if (dftdid!='') {tdid=dftdid}
if (isfrist==1){
var vtrntblock=document.getElementById('trntblock');
var newdiv=document.createElement('div'); 
newdiv.setAttribute('id','adtop' +tdid); 
setAttributezdy(newdiv,'class','adtop170528')
//newdiv.setAttribute('class', 'adtop'); 
newdiv.innerHTML = "<span class=\"num170528\" id=\"num"+tdid+"\"></span><span class=\"close170528\" onClick=\"return clkclose170522("+tdid+")\">x</span><div id=\"nt"+tdid+"\"></div>";
vtrntblock.appendChild(newdiv);
}
var adtop = document.getElementById("adtop"+ tdid);
var adtopi = document.getElementById("num"+ tdid);
var nt170522 = document.getElementById("nt"+ tdid);
nt170522.innerHTML=str170522
if (isfrist==1){adtop.style.display="block";}
if (isfrist!=1 && adtop.style.display=="none"){
return false;}
s--;
adtopi.innerHTML=s;
if(s<=0){
adtop.style.display="none";
return;
}
return setTimeout(function(){adClose170522(s,str170522,0,tdid)},1000);
}; 
