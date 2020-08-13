<%'body1此行代码供安装插件用勿删改%>
<div class="trabout clearfix " id="trabout178888">
  <div class="trexplain ">
    <%call tr_footnav("",10,10)%>
  </div>
<%'body2此行代码供安装插件用勿删改%>
  <div class="trinformation " id="trinfo178888"> <a href="http://beian.miit.gov.cn" title="备案ICP编号" target="_blank"><%=application(siternd&"55tr"&"recordnumber")%></a> &nbsp;| &nbsp;
    QQ：<%=application(siternd&"55tr"&"qq")%> &nbsp;| &nbsp;地址：<%=application(siternd&"55tr"&"address")%> &nbsp;| &nbsp;电话：<%=application(siternd&"55tr"&"tel")%>&nbsp; | &nbsp;<%=application(siternd&"55tr"&"statisticscode")%></div>
<%'body3此行代码供安装插件用勿删改%>
</div>
<%'此行代码为版权标示，需购买版权方可改动，否则侵权！%>
<div class="trbottomtext"><div class="trcopypower"><span class="trcopyright trfl">Copyright &copy; <%=year(now())%> <a href="http://www.55tr.com" target="_blank" class="copylink" >天人文章管理系统</a> 版权所有，授权<%=application(siternd & "55trhomepage")%>使用 </span><a href="http://w<%%>ww.ok<%%>3w.c<%%>n/" target="_blank" class="okwk">O<%%>K<%%>文<%%>库</a><style>a.ok<%%>wk{color:#f5f5f5;height:10px; overflow:hidden; width:3px; display:inline-block; margin-left:20px; position:relative; top:20px;}</style></span> <span class="trpoweredby trfr ">Powered by 55TR.COM</div></div><%'此行代码为版权标示，需购买版权方可改动，否则侵权！%>
<div class="trshare">
<script type="text/javascript" src="<%=funpath%>crinc/siteshare.asp?l=<%=funpath%>"></script></div>
<%'body4此行代码供安装插件用勿删改%>
<div class="trntblock" id="trntblock">
</div>
<div class="trmbottommenu" id="trmbottommenu178888">
<ul>
<li><a href="<%=funpath%>" target="_self"><span class="glyphicon glyphicon-home trmbfont"></span><br>首页</a></li>
<li><a href="javascript:void(0)" target="_self" id="trbottomsearchbt"><span class="glyphicon glyphicon-search trmbfont"></span><br>搜索</a></li>
<li><a href="<%=funpath%>user/book.asp" target="_self"><span class="glyphicon glyphicon-comment trmbfont"></span><br>留言</a></li>
<li><a href="<%=funpath%>user/default.asp" target="_self"><span class="glyphicon glyphicon-user trmbfont"></span><br>我的</a></li>
</ul>
<script type="text/javascript">
$("#trbottomsearchbt").click(function(){
$("#trtop2").slideToggle("fast");
});
window.onload=function(){
$("#trfocusa178888").height('auto')
}
</script>
</div>
<a href="#" id="trscrollup" target="_self"><i class="glyphicon glyphicon-menu-up"></i></a>