<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>网站后台顶部</title>
<link rel="stylesheet" type="text/css" href="./skin/default/style.css">
<link rel="shortcut icon" type="image/x-icon" href="../favicon.ico">
<link rel="bookmark" type="image/x-icon" href="../favicon.ico">
<script>
function showtophelp(obj)
{obj.style.display ="block" ;}
function hidetophelp(obj)
{obj.style.display ="none" ;}
function showhidehelp(obj)
{obj.style.display = obj.style.display == "none" ? "block" : "none";}

</script>
</head>
<body>
<div class="top1">
  <div class="top1_logo fl"> </div>
  <div class="top1bt fr">
<a href="javascript:void(0)" target="_self" class="topabg7" onclick="showhidehelp(tophelpdiv)" onMouseOut="hidetophelp(tophelpdiv)">
帮助
</a>
<a href="javascript:window.parent.frames['main'].history.go(-1)" target="_self" class="topabg4">
后退
</a>
<a href="javascript:window.parent.frames['main'].location.reload()" target="_self" class="topabg5">
刷新
</a>
<a href="javascript:window.parent.frames['main'].history.go(1)" target="_self" class="topabg6">
前进
</a>
<a href="../" target="_blank" class="topabg1">
前台
</a>
<a href="main.asp" target="main" class="topabg2">
后台
</a>
 <a href="logout.asp" target="_top" onClick="if(!confirm('真的要退出吗?')) return false;" class="topabg3">
退出
</a>
<div class="tophelpdiv" id="tophelpdiv"  onMouseOver="showtophelp(tophelpdiv)" onMouseOut="hidetophelp(tophelpdiv)">
<iframe src="https://www.55tr.com/inc/trinfo.asp?s=1&t=1" width="100%" height="100%" align="center" scrolling="no" frameborder="0" style="background-color=transparent" allowtransparency="true" ></iframe>
</div> 
  </div>
</div>
</body>
</html>
