<script type="text/javascript">
<%
id = tr_killstr(request.querystring("id"), 1, 8, 1, "", "", "", "", 1)
pageno = tr_killstr(request.querystring("pageno"), 1, 8, 1, "", "", "", "", 1)
keywordsak = tr_killstr(request.querystring("keywords"), 2, 30, 2, "", "", "", "", 1)
HTTP_HOST=Request.ServerVariables("HTTP_HOST")
PATH_INFO=Request.ServerVariables("PATH_INFO")
if instr(1,PATH_INFO,"Default.asp",1)>0 or instr(1,PATH_INFO,"index.asp",1)>0 then
tourl=funpath&"m/default.asp"
end if
if instr(1,PATH_INFO,"show.asp",1)>0 then
tourl=funpath&"m/show.asp?id="&id&"&pageno="&pageno
end if
if instr(1,PATH_INFO,"list.asp",1)>0 then
tourl=funpath&"m/list.asp?id="&id&"&pageno="&pageno
end if
if instr(1,PATH_INFO,"trsearch.asp",1)>0 then
tourl=funpath&"m/trsearch.asp?keywords="&keywordsak&"&pageno="&pageno
end if
if instr(1,PATH_INFO,"book.asp",1)>0 then
tourl=funpath&"m/book.asp?pageno="&pageno
end if

if instr(1,PATH_INFO,"user/login.asp",1)>0 then
tourl=funpath&"m/login.asp"
end if
if instr(1,PATH_INFO,"user/reg.asp",1)>0 then
tourl=funpath&"m/reg.asp"
end if
if instr(1,PATH_INFO,"user/mynews.asp",1)>0 then
tourl=funpath&"m/mynews.asp"
end if
if instr(1,PATH_INFO,"user/unew.asp",1)>0 then
tourl=funpath&"m/unew.asp?id="&id&""
end if
if instr(1,PATH_INFO,"user/sign.asp",1)>0 then
tourl=funpath&"m/sign.asp"
end if
if instr(1,PATH_INFO,"user/default.asp",1)>0 then
tourl=funpath&"m/udefault.asp"
end if

%>
//		var newurl='<%=tourl%>';
//    (function (Switch) {
//        var switch_pc = window.location.hash;
//        if (switch_pc != "#pc") {
//            if (/iphone|ipod|ipad|ipad|Android|nokia|blackberry|webos|webos|webmate|bada|lg|ucweb|skyfire|sony|ericsson|mot|samsung|sgh|lg|philips|panasonic|alcatel|lenovo|cldc|midp|wap|mobile/i.test(navigator.userAgent.toLowerCase())) {
//                Switch.location.href = newurl;
//            }
//        }
//    })(window);
//    document.write('<meta name="mobile-agent" content="format=xhtml;url=' + newurl + '" /><link href="' + newurl + '" rel="alternate" media="only screen and (max-width: 1000px)" />');

</script>