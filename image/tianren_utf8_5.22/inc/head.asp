<script>
function showHide0123(obj)
{obj.style.display = obj.style.display == "none" ? "block" : "none";}
</script>
<%'body1此行代码供安装插件用勿删改%>
<div class="trtopbar" id="trtopbar178888">
<div class="container trtop" id="trrolltext178888">
<%call tr_rolltextnd("",5,60,0,0)%>
<div class="trtop2 trfr clearfix" id="trtop2">
<form action="<%=funpath%>trsearch.asp" target="_blank" name="formsc1" id="formsc1" method="get" onsubmit="return checksearch(this)">
<input name="keywords" type="text" class="trsearchs trfl" id="keywords"  onfocus="this.value='';" onblur="if(this.value.length < 1) this.value='请输入关键词';" value="请输入关键词" maxlength="30" />
<button type="submit" class="trsearchbt trfl "> <span class="glyphicon glyphicon-search" aria-hidden="true"></span> </button>
</form>
</div>
<%'body6此行代码供安装插件用勿删改%>
<div class="trtop3 trfr" id="trtoplink178888"><!--#include file="crphoneewm.asp"--> 
<span id="unamequit"><a href="<%=funpath%>user/login.asp" id="uname">登陆</a> | <a href="<%=funpath%>user/reg.asp" id="quit" >注册</a> |</span> <a href="<%=funpath%>user/book.asp">留言</a> | <a onclick="SetHome(window.location)" href="javascript:void(0);" target="_self">设首页</a> | <a onclick="AddFavorite(window.location,document.title)" href="javascript:void(0);" target="_self">加收藏</a> <script type="text/javascript">changeloginreg()</script> </div>
</div>
</div>
<%'body2此行代码供安装插件用勿删改%>
<nav class="navbar navbar-default trnavbar">
<div class="container trlogoother " id="trlogoother">
<%'body7此行代码供安装插件用勿删改%>
<div class="trlogor trad1" id="trlogoright178888"><script type="text/javascript"> _55tr_com('tr1')</script></div>
<div class="trlogo clearfix" id="trlogodiv">
<table id="trlogotb">
<tr>
<td align="left" valign="middle"><a href="<%=funpath%>" title=""><img src="<%=funpath%><%=application(siternd & "55trlogopic")%>" alt=""/></a></td>
</tr>
</table>
<%'body8此行代码供安装插件用勿删改%>
<button type="button" class="navbar-toggle trmnavbtn" data-toggle="collapse" data-target="#navbar-collapse"> <span class="sr-only">手机导航</span> <span class="icon-bar"></span> <span class="icon-bar"></span> <span class="icon-bar"></span> </button>
<%'body9此行代码供安装插件用勿删改%>
</div>
</div>
<script type="text/javascript">
function trautoLogoPlace(){
if (document.body.clientWidth>=1200 ){
otrlogotb=document.getElementById("trlogotb")
trlogoother=document.getElementById("trlogoother")
trlogodiv=document.getElementById("trlogodiv")
otrlogotb.style.height =trlogoother.offsetHeight+'px'
otrlogotb.style.width =(trlogodiv.offsetWidth)+'px'
}
}
trautoLogoPlace()
//window.onresize=function(){trautoLogoPlace()}
</script>
<%'body3此行代码供安装插件用勿删改%>
<div class="trnav" id="trnav178888">
<%call tr_headnavnd(500)%>
</div>
</nav>
<%'body4此行代码供安装插件用勿删改%>
<div class="container">
<div class="trad3"> 
<script type="text/javascript"> _55tr_com('tr3')</script> 
</div>
<div class="trad3 clearfix"> 
<script type="text/javascript"> _55tr_com('tr9')</script> 
</div>
</div>
<%'body5此行代码供安装插件用勿删改%>
