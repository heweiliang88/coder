<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../inc/charsetasp.asp"-->
<!DOCTYPE html>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<%
'版权声明：
'天人系列管理系统，以下简称“本程序、程序”可免费用于个人非商业用途，其他用途需付费购买版权，请按照如下约定执行：
'**本程序可免费用于个人非商业用途，除个人非商业外的用途均需要付费购买版权，并在官网登记域名作为正规版权记录，并可获得电子版版权证书(不是绑定域名，而是在官网数据库中登记域名旨在法律维权时避开已购买版权的域名)，侵权使用必将承担法律风险并追究高额赔偿
'**购买版权将获得稳定健全的升级服务、免费安装插件、模板、配套工具免费使用、官方优惠优先获得等特权。并可用于企业、政府、校园、单位、团体、个人、机构、部门、机关、公司及其他商业用途
'**购买版权后，除源代码中版权声明，后台左下角的“源程序：55tr.com”字样（这两种版权字样都不会在前台显示，只作为程序版权归属）以外所有代码都可以修改及删除，包括前台页面底部的官网链接及版权等代码都可修改。
'**购买版权后可用本程序为别人或自己提供建设网站，修改源码，二次开发的服务，但禁止发布或传播本程序及本程序的任何形态！禁止出售本程序及本程序的任何形态。说白了就是不能将程序出售或传播，否则侵权。（例如：小李与客户约定使用本程序制作某中学的官网，并收取某中学的网站制作服务费用1万元，并购买版权，此行为不侵权。小赵与客户约定使用本程序制作某小学的官网，并收取某小学的网站制作服务费用1万元，未购买版权，此行为侵权(因为未购买版权)，请终止！小孙使用本程序建立了个人漫画网站供用户浏览并获取广告收益，未购买版权，此行为不侵权。小张使用本程序制作为企业网站进行出售，并购买版权，此行为侵权（因为传播出售了程序），请终止！小王又将本程序制作为小说网站源码在网络免费发布传播提供下载，并购买版权，此行为侵权（因为传播出售了程序），请终止！）
'**不购买版权的个人非商业用途依然享受升级、安装插件、模板的权利，可修改源代码，但不得修改源代码中“官网链接”、“版权声明”、“应用中心”相关的代码，前台页面底部的“Copyright 2016 天人文章管理系统 版权所有，授权xxx使用 Powered by 55TR.COM”字样不得删除,否则请购买版权
'**不购买版权的除个人非商业外的用途请及时购买版权或终止使用本程序，否则侵权
'**计算机软件著作权登记证书登记号：2016SR204636，侵权必究，作者QQ81962480 欢迎洽谈基于此程序开发、修改业务。
'**版权声明会存在修改的偶然性，请以官网公布的对应程序版本版权声明为准！具体的版权声明、特权、及功能介绍，请浏览：http://www.55tr.com/gwappshow.asp?id=11

Const dbpath = "../"
%>
<!--#include file="admin_fun.asp"-->
<%
If InStr(","&adminlimitidrange&",", ",")<1 Then contrl_message "您没有访问权限，请尝试其他版块。", "", "", "", ""'instr逗号的作用是只要有这个管理员就行，不管他是否具备权限
%>
<title>后台菜单页 -<%=application(siternd & "55trsitetitle")%>%></title>
<link rel="stylesheet" type="text/css" href="./skin/default/style.css">
<script type="text/javascript" src="./js/js.js"></script>
<base target="main" />
<!--[if lte IE 6]>
<style type="text/css">
body { behavior:url("../js/csshover3.htc"); }
</style>
<![endif]-->
<link rel="shortcut icon" type="image/x-icon" href="../favicon.ico">
<link rel="bookmark" type="image/x-icon" href="../favicon.ico">
</head>
<body target="main" style="background:#fff; margin:3px;">
<div class="left1">
<div class="managestr "> <a href="admin_info.asp" target="main"><%=lefts(adminname,20)%></a> </div>
<div class="left_block mar1">
<div class="left_tt font6 menubg1" onClick='showHide(items1)'>
<div class="fl">个人管理</div>
<div class="down_arr fr"></div>
</div>
<div class="left_nr" style="display:none" id="items1">
<ul>
<a href="../" target="_blank">前台首页</a> <a href="main.asp" target="main">后台首页</a> <a href="serverinfo.asp" target="main">服务器端</a> <a href="admin_info.asp" target="main">修改资料</a> <a href="logout.asp" target="_top" onClick="if(!confirm('真的要退出吗？')) return false;">安全退出</a>
</ul>
</div>
</div>
<%if instr(","&adminlimitidrange&",",",6,")>0 then %>

<div class="left_block mar1">
<div class="left_tt font6 menubg2" onClick='showHide(items2)'>
<div class="fl">应用中心</div>
<div class="down_arr fr"></div>
</div>
<div class="left_nr" style="display:none" id="items2">
<ul>
<a href="ifmapp.asp?lx=2" target="main">获取模板</a> <a href="mytem.asp" target="main">我的模板</a> <a href="ifmapp.asp?lx=1" target="main">获取插件</a> <a href="myapp.asp" target="main">我的插件</a> <a href="keys.asp" target="main">注册码</a> <a href="upfileup.asp" target="main">上传安装</a> <a href="main.asp" target="main">程序升级</a>
</ul>
</div>
</div>
<%end if%>
<%if instr(","&adminlimitidrange&",",",1,")>0 then %>
<div class="left_block mar1">
<div class="left_tt font6 menubg3" onClick='showHide(items3)'>
<div class="fl">网站管理</div>
<div class="down_arr fr"></div>
</div>
<div class="left_nr" style="display:none" id="items3">
<ul>
<a href="site_set.asp" target="main">站点设置</a> <a href="admin.asp" target="main">管理员</a> <a href="ads.asp" target="main">广告管理</a> <a href="database_info.asp" target="main">数据库</a> <a href="link.asp" target="main">友情链接</a> <a href="log.asp" target="main">后台日志</a> <a href="recache.asp" target="main">更新缓存</a> <a href="share.asp" target="main">分享按钮</a>
</ul>
</div>
</div>
<%end if%>
<%if instr(","&adminlimitidrange&",",",2,")>0 then %>
<div class="left_block mar1">
<div class="left_tt font6 menubg4" onClick='showHide(items4)'>
<div class="fl">文章管理</div>
<div class="down_arr fr"></div>
</div>
<div class="left_nr" style="display:none" id="items4">
<ul>
<a href="column.asp" target="main">栏目管理</a> <a href="article.asp" target="main">新增文章</a> <a href="article_list.asp" target="main">文章列表</a> <a href="article_list.asp?isdel=1" target="main"><font color="#0000ff">回收站</font> </a> <a href="comment_list.asp?iscom=1" target="main">评论管理</a> <a href="comment_list.asp?isdel=1&iscom=1" target="main"><font color="#0000ff">回收站</font></a> <a href="html.asp" target="main" style="color:#ff0000;">整站静态</a>
</ul>
</div>
</div>
<%end if%>
<%if instr(","&adminlimitidrange&",",",5,")>0 then %>
</p>
<div class="left_block mar1">
<div class="left_tt font6 menubg5" onClick='showHide(items5)'>
<div class="fl">采集管理</div>
<div class="down_arr fr"></div>
</div>
<div class="left_nr" style="display:none" id="items5">
<ul>
<a href="cj/SK_GetArticle.asp?action=add&Colleclx=1" target="main">新增规则</a> <a href="cj/SK_GetArticle.asp?Colleclx=1" target="main">规则管理</a> <a href="cj/sk_checkDatabase.asp" target="main">数据入库</a> <a href="cj/sk_ItemHistroly.asp" target="main">历史记录</a>
</ul>
</div>
</div>
<%end if%>
<%if instr(","&adminlimitidrange&",",",3,")>0 then %>
<div class="left_block mar1">
<div class="left_tt font6 menubg6" onClick='showHide(items6)'>
<div class="fl">留言管理</div>
<div class="down_arr fr"></div>
</div>
<div class="left_nr" style="display:none" id="items6">
<ul>
<a href="comment_list.asp" target="main">全部留言</a> <a href="comment_list.asp?isdel=1" target="main"><font color="#0000ff">回收站</font></a>
</ul>
</div>
</div>
<%end if%>
<%if instr(","&adminlimitidrange&",",",4,")>0 then %>
<div class="left_block mar1">
<div class="left_tt font6 menubg7" onClick='showHide(items7)'>
<div class="fl">会员管理</div>
<div class="down_arr fr"></div>
</div>
<div class="left_nr" style="display:none" id="items7">
<ul>
<a href="user.asp?t=1" target="main">添加会员</a> <a href="userlev.asp?t=1" target="main">管理等级</a> <a href="user_list.asp" target="main">管理会员</a> <a href="user_list.asp?isdel=1" target="main"><font color="#0000ff">回收站</font> </a>
</ul>
</div>
</div>
<%end if%>
<%if instr(","&adminlimitidrange&",",",6,")>0 then %>
<div class="left_block mar1">
<div class="left_tt font6 menubgqtgn170820" onClick='showHide(itemsqtgn020820)'>
<div class="fl">高级功能</div>
<div class="down_arr fr"></div>
</div>
<div class="left_nr" style="display:block" id="itemsqtgn020820">
<ul>
<a href="theplug020116.asp?appid=103&toappid=103" target="main">自动采集</a>
<a href="theplug020116.asp?appid=174&toappid=174" target="main">熊掌号推送</a>
<a href="theplug020116.asp?appid=64&toappid=64" target="main">全能搜索引擎推送</a>
<a href="theplug020116.asp?appid=102&toappid=102" target="main">视频音频播放</a>
<a href="theplug020116.asp?appid=96&toappid=96" target="main">会员充值</a>
<a href="theplug020116.asp?appid=61&toappid=61" target="main">万能伪静态</a>
<a href="theplug020116.asp?appid=95&toappid=95" target="main">强制弹窗</a>
<a href="theplug020116.asp?appid=205&toappid=205" target="main">自动生成静态</a>
<a href="theplug020116.asp?appid=178&toappid=178" target="main">内容付费</a>
<a href="theplug020116.asp?appid=100&toappid=100" target="main">内容定时发布</a>
<a href="theplug020116.asp?appid=97&toappid=97" target="main">内容部分隐藏预览</a>
<a href="theplug020116.asp?appid=80&toappid=80" target="main">微信快捷登录</a>
<a href="theplug020116.asp?appid=78&toappid=78" target="main">QQ快捷登录</a>
<a href="theplug020116.asp?appid=79&toappid=79" target="main">微博快捷登录</a>
<a href="theplug020116.asp?appid=214&toappid=214" target="main">微信自定义分享</a>
<a href="theplug020116.asp?appid=209&toappid=209" target="main">足球篮球实时比分</a>
<a href="theplug020116.asp?appid=74&toappid=74" target="main">管理员栏目权限</a>
<a href="theplug020116.asp?appid=54&toappid=54" target="main">广告可视化编辑</a>
<a href="theplug020116.asp?appid=22&toappid=22" target="main">二维码打赏</a>
<a href="theplug020116.asp?appid=24&toappid=24" target="main">禁止复制与右键</a>
<a href="theplug020116.asp?appid=195&toappid=195" target="main">站群一键批量控制</a>
<a href="theplug020116.asp?appid=77&toappid=77" target="main">栏目及文章批量迁移</a>
<a href="theplug020116.asp?appid=193&toappid=193" target="main">站内搜索关键词记录</a>
<a href="theplug020116.asp?appid=124&toappid=103" target="main">美女图片采集</a>
<a href="theplug020116.asp?appid=222&toappid=103" target="main">F1赛车内容采集</a>
<a href="theplug020116.asp?appid=216&toappid=103" target="main">彩票内容采集</a>
<a href="theplug020116.asp?appid=208&toappid=103" target="main">体育内容采集</a>
<a href="theplug020116.asp?appid=219&toappid=103" target="main">棋牌内容采集</a>
<a href="theplug020116.asp?appid=213&toappid=103" target="main">养生内容采集</a>
<a href="theplug020116.asp?appid=212&toappid=103" target="main">母婴内容采集</a>
<a href="theplug020116.asp?appid=198&toappid=103" target="main">财经内容采集</a>
<a href="theplug020116.asp?appid=197&toappid=103" target="main">新闻内容采集</a>
<a href="theplug020116.asp?appid=221&toappid=103" target="main">明星娱乐内容采集</a>
</ul>
</div>
</div>
<%end if%>
<div class="version mar1">
<ul>
<li>天人文章管理系统</li>
<!--此行代码切勿改动-->
<li><%=trversion%></li>
<!--此行代码切勿改动--> 
<!--此行代码切勿改动-->
<li>源程序：<a target="_blank" href="http://www.55tr.com">55TR.COM</a></li>
<!--此行代码切勿改动-->
</ul>
</div>
</div>
<%
Set ctr_admin = Nothing
%>
</body>
</html>

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
