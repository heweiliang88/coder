<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../../inc/charsetasp.asp"-->
<%
'采集内核：SK采集系统
'QQ81962480 天人文章管理系统
%>
<!doctype html>

<%
option explicit
response.buffer=true
%>
<!--#include file="inc/setup.asp"-->
<!--#include file="inc/cj_cls.asp"-->

<%
If InStr(","&adminlimitidrange&",", ",5,")<1 Then contrl_message "您没有访问权限，请尝试其他版块。", "", "", "", ""

Dim SqlItem,RsItem,Action,ItemID,FilterName,FilterID,FilterObject,FilterType,FilterContent,FisString,FioString,FilterRep,Flag,PublicTf
Dim FoundErr,MaxPerPage,CurrentPage,AllPage,HistrolyNum,i_His,lx,DelFlag
Action=Trim(Request("Action"))
lx=Request("radiobutton")
FilterID=Trim(Request("FilterID"))
DelFlag=Trim(Request.Form("DelFlag"))
MaxPerPage=20
If lx="" or lx=0 then lx=1
Call top()
If action<>"" And adminctrllimit = 1 Then Call contrl_message ("您只有查看权限，没有管理权限，请升级权限后继续操作。", "", "", "", "")
Select Case Action
Case "add"
	 If Trim(Request.Form("Saveok"))="ok" then
	 	Call Save_Data
	 Else 
	 	Call add_edit(1)
	 End if
Case "edit"
	 If Trim(Request.Form("Saveok"))="ok" then
	 	Call Save_Data
	 Else 
	 	Call add_edit(2)
	 End if
Case "flag"
	 flag=Trim(Request("flag"))
	 ConnItem.execute("update Filters set flag="& flag &"  where FilterID="&FilterID )
	 response.redirect request.ServerVariables("HTTP_REFERER")''返回来源页面有刷新
Case "del"
	 Select Case DelFlag
	 Case "删除所选记录"
	 Response.Write("sadfsadfasdfasdfasfasfd")
	 	if FilterID<>"" then 
			 ConnItem.execute("Delete From Filters Where FilterID in(" & FilterID & ")")
		end if
	 Case "删除全部记录"
	 	if lx=1 then ConnItem.execute("Delete From Filters where colleclx=1")
		if lx=3 then ConnItem.execute("Delete From Filters where colleclx=3")
		if lx=5 then ConnItem.execute("Delete From Filters where colleclx=5")
	 Case Else
	 	if FilterID<>"" then 
			 ConnItem.execute("Delete From Filters Where FilterID in(" & FilterID & ")")
			 response.redirect request.ServerVariables("HTTP_REFERER")''返回来源页面有刷新
		end if	
	 End Select
	 	Response.Redirect "sk_itemfilters.asp?radiobutton="& lx 
Case Else
	Select Case lx
	case 1
	   Call Main(1)'新闻
	   Call Show_Page
	case 3
	   Call Main(3)'图片
	   Call Show_Page
	case 5
	   Call Main(5)'软件
	   Call Show_Page
	case 6
	   Call Main(6)'自定
		Call Show_Page
	case else
	   Call Main(1)
	   Call Show_Page
	end select
End select
'关闭数据库链接
Call CloseConnItem()
%>
<%sub top()%>

<html>
<head>
<title>天人文章管理系统采集模块</title>
<meta charset="utf-8">
<link rel="stylesheet" type="text/css" href="../skin/default/style.css">
<SCRIPT language="javascript" type="text/javascript">
function showset(thisform)
{
		if(thisform.FilterType.selectedIndex==1)
		{
        	FilterType1.style.display = "none";
			FilterType2.style.display = "";
		}
		else
		{
        	FilterType1.style.display = "";
			FilterType2.style.display = "none";
	    }

}
</script>
<SCRIPT language="javascript" type="text/javascript">
function unselectall(thisform)
{
    if(thisform.chkAll.checked)
	{
		thisform.chkAll.checked = thisform.chkAll.checked&0;
    } 	
}

function CheckAll(thisform)
{
	for (var i=0;i<thisform.elements.length;i++)
    {
	var e = thisform.elements[i];
	if (e.Name != "chkAll"&&e.disabled!=true)
		e.checked = thisform.chkAll.checked;
    }
}
//-->
</script>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1"  class="tableBorder">
  <tr class='topbg'> 
    <td height="22" colspan="2" align="center" bgcolor="#F2F2F2" ><strong>采 集 系 统 过 滤 管 理</strong></td>
  </tr>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1"  class="tableBorder">
  <tr bgcolor="#F2F2F2" class="tdbg"> 
    <td width="65" height="30"><strong>管理导航：</strong></td>
    <td height="30"><a href="sk_ItemFilters.asp">管理首页</a> | <a href="?Action=add&radiobutton=<%=lx%>">添加新过滤</a></td>
  </tr>
</table>
<table width="100%" align="center"  border="0" cellspacing="0" cellpadding="0" class="tableBorder">
  <tr>
    <td height="30" align="center" bgcolor="#F2F2F2"> 选择类型：
      <input name="radiobutton" type="radio" value="1" <%if lx =1 then Response.Write "checked" %>  onclick="location.href='?radiobutton=1';" >
      新闻采集
      <input type="radio" name="radiobutton" value="3"  <%if lx =3  then Response.Write "checked" %> onClick="location.href='?radiobutton=3';">
      图片采集
      <input type="radio" name="radiobutton" value="5"  <%if lx =5  then Response.Write "checked" %> onClick="location.href='?radiobutton=5';">
    软件采集 </td>
  </tr>
</table>

<%
end sub 
'---------文章列表------------------
Sub Main(lx)
%>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="tableBorder" >
    <tr> 
      <td height="22" colspan="2" class="title"> <div align="center"><strong>所 有 记 录</strong></div></td>
    </tr>
</table>
 <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="tableBorder">
  <form name="form1" method="post" action="?action=del">
    <tr class="tdbg" style="padding: 0px 2px;"> 
      <td width="57" height="22" align="center" bgcolor="#EFEFEF" class="ButtonList">选择</td>
      <td width="176" align="center" bgcolor="#EFEFEF" class="ButtonList">所属采集项目</td>
      <td width="246" align="center" bgcolor="#EFEFEF" class="ButtonList">过滤名称</td>
      <td width="114" height="22" align="center" bgcolor="#EFEFEF" class="ButtonList">过滤类型</td>
      <td width="134" align="center" bgcolor="#EFEFEF" class="ButtonList">过滤属性</td>
	  <td width="30" align="center" bgcolor="#EFEFEF" class="ButtonList">状态</td>
      <td width="154" height="22" align="center" bgcolor="#EFEFEF" class="ButtonList">操作</td>
    </tr>
    <%                          
Set RsItem=server.createobject("adodb.recordset")
SqlItem="select * from Filters where colleclx=" & lx 
If Request("page")<>"" then
    CurrentPage=Cint(Request("Page"))
Else
    CurrentPage=1
End if 

SqlItem=SqlItem  &  " order by FilterID DESC"
RsItem.open SqlItem,ConnItem,1,1
If (Not RsItem.Eof) and (Not RsItem.Bof) then
   RsItem.PageSize=MaxPerPage
   Allpage=RsItem.PageCount
   If Currentpage>Allpage Then Currentpage=1
   HistrolyNum=RsItem.RecordCount
   RsItem.MoveFirst
   RsItem.AbsolutePage=CurrentPage
   i_His=0
   Do While not RsItem.Eof
%>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#cccccc'" style="padding: 0px 2px;"> 
      <td width="57" align="center"> <input type="checkbox" value="<%=RsItem("FilterID")%>" name="FilterID" onClick="unselectall(this.form)" style="border: 0px;background-color: #E1F4EE;">      </td>
      <td width="176" align="center"> 
	  <%dim rs : set RS=ConnItem.execute("select * from Item where ItemID=" & RsItem("ItemID"))
	  if rs.eof then
	  	Response.Write("没有项目!")
	  else
	  	Response.Write RS("ItemName") 
	  end if
	   rs.close : set rs=nothing%> </td>
      <td width="246" align="left"><%=RsItem("FilterName")%></td>
      <td width="114" align="center"><%if RsItem("FilterObject")=1 then Response.Write "标题过滤" else Response.Write "正文过滤" end if%></td>
      <td align="center"> <%if RsItem("FilterType")=1 then Response.Write "简单替换" else Response.Write "高级替换" end if%></td>
      <td width="30" align="center">
	  <%if RsItem("Flag")=1 then Response.Write "√" else Response.Write "<font color=""#FF0000"">×</font>" end if%>
	  </td>
	  <td width="154" align="center">
	  <%if RsItem("Flag")=0 then Response.Write "<a href=""?Action=flag&flag=1&FilterID="& RsItem("FilterID") &"&radiobutton="& lx &""">启用</a>" else Response.Write "<a href=""?Action=flag&flag=0&FilterID="& RsItem("FilterID") &"&radiobutton="& lx &""">禁用</a>" end if%>
	   &nbsp;<a href="?Action=edit&FilterID=<%=RsItem("FilterID")%>&radiobutton=<%=lx%>" >编辑</a> &nbsp;
      <a href="?Action=del&FilterID=<%=RsItem("FilterID")%>&radiobutton=<%=lx%>" onclick='return confirm("确定要删除此记录吗？");'>删除</a>      </td>
    </tr>
    <%         
           i_His=i_His+1
           If i_His > MaxPerPage Then
              Exit Do
           End If
        RsItem.Movenext         
   Loop         
%>
    <tr class="tdbg"> 
      <td colspan="8" height="30">
        <input name="Action" type="hidden" id="Action" value="ok"> 
		<input name="radiobutton" type="hidden" id="Action" value="<%=lx%>"> 
        <input name="chkAll" type="checkbox" id="chkAll" onclick="CheckAll(this.form)" value="checkbox" style="border: 0px;background-color: #E1F4EE;">
        全选 </td>
    </tr> 
    <tr class="tdbg"> 
      <td colspan="8" height="30" align="center">&nbsp;&nbsp;&nbsp;&nbsp; <input name="DelFlag" type="submit" class="lostfocus" style="cursor: hand;background-color: #cccccc;"  onclick='return confirm("确定要删除选中的记录吗？");' value="删除所选记录"> 
        &nbsp;&nbsp;&nbsp;&nbsp; <input name="DelFlag" type="submit" class="lostfocus" style="cursor: hand;background-color: #cccccc;" onclick='return confirm("确定要删除所有的记录吗？");' value="删除全部记录"> 
      &nbsp;&nbsp;
    </td></tr>
    <tr class="tdbg"> 
      <td colspan="8" height="30"> </td>
    </tr>
    <%Else%>
    <tr class="tdbg"> 
      <td colspan='9' class="tdbg" align="center"><br>
        系统中暂无记录！</td>
    </tr>
    <%End  If%>
    <%         
RsItem.Close         
Set RsItem=nothing           
%>
  </form>
</table>  
<%End Sub%>

<% '----------显示“上一页 下一页”等信息-----------
sub Show_Page() %>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="tableBorder" >
    <tr> 
      <td height="22" colspan="2" class="tdbg">
<%
Response.Write ShowPage("sk_checkDatabase.asp?Action="& Action,HistrolyNum,MaxPerPage,True,True," 个记录")
%>

      </td>
    </tr>
</table>
<%end sub

'--------------过滤添加和编辑--------------
sub add_edit(Type_1)
If Type_1=2 then Call Test
%>
<form method="post" action="?Action=<% if Type_1 =1 then Response.Write("add") else Response.Write("edit") end if %>" name="form1">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1"  class="tableBorder" >                                
    <tr>                                 
      <td height="22" colspan="2" class="title"> <div align="center"><strong><% if Type_1=1 then Response.Write "添 加 过 滤" else  Response.Write "编  辑 过 滤" end if %></strong></div></td>                                
    </tr>                                
    <tr class="tdbg">                                 
      <td width="150" class="tdbg"><strong> 过滤名称：</strong></td>                                
      <td class="tdbg"><input name="FilterName" type="text" id="FilterName" value="<%=FilterName%>" size="25" maxlength="30">                                 
        &nbsp;</td>                                
    </tr>                                
    <tr class="tdbg">                                 
      <td width="150" class="tdbg"><strong> 所属项目：</strong></td>                                
      <td class="tdbg"><%Call Admin_ShowItem_Option(ItemID,lx)%>      </td>                                
    </tr>                                
    <tr class="tdbg">                                 
      <td width="150" class="tdbg"><strong> 过滤对象：<strong></td>                               
      <td class="tdbg">                               
         <select name="FilterObject" id="FilterObject">                              
            <option value="1" <%if FilterObject=1 Then Response.Write "selected"%>>标题过滤</option>                              
            <option value="2" <%if FilterObject=2 or FilterObject="" or FilterObject=0  Then Response.Write "selected"%>>正文过滤</option>             
         </select>      </td>                               
    </tr>                               
    <tr class="tdbg">                                
      <td width="150" class="tdbg"><strong> 过滤类型：</strong></td>                               
      <td class="tdbg">                               
         <select name="FilterType" id="FilterType" onchange="showset(this.form)">                              
            <option value="1" <%if FilterType=1 Then Response.Write "selected"%> >简单替换</option>                              
            <option value="2" <%if FilterType=2 Then Response.Write "selected"%>>高级过滤</option>                              
         </select>      </td>                                
    </tr>     
    <tr class="tdbg">                                
      <td width="150" class="tdbg"><strong> 使用状态：</strong></td>                               
      <td class="tdbg">                            
         <select name="Flag" id="Flag">    
		 	<option value="0" <%If Flag=0 Then Response.Write "selected"%>>禁用</option>                           
            <option value="1" <%If Flag=1 or Flag="" Then Response.Write "selected"%>>启用</option>                            
         </select>      </td>                                
    </tr>                                                                         
</table>                                
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1"  class="tableBorder" id="FilterType1" style="display:<%if FilterType<>1 and FilterType<>"" Then Response.Write "none"%>">                                
    <tr class="tdbg">                                 
      <td width="150" class="tdbg"><strong> 内容：</strong></td>                               
      <td class="tdbg"><textarea name="FilterContent" cols="70" rows="5" class="lostfocus"><%=FilterContent%></textarea></td>                               
    </tr>                            
</table>                 
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1"  class="tableBorder" id="FilterType2" style="display:<%if FilterType<>2 Then Response.Write "none"%>">                                   
    <tr class="tdbg">                               
      <td width="150" class="tdbg"><strong> 开始标记：</strong></td>                              
      <td class="tdbg"><textarea name="FisString" cols="70" rows="5" class="lostfocus"><%=FisString%></textarea>                               
        &nbsp;</td>                                
    </tr>                                
    <tr class="tdbg">                                 
      <td width="150" class="tdbg"><strong> 结束标记：</strong></td>                               
      <td class="tdbg"><textarea name="FioString" cols="70" rows="5" class="lostfocus"><%=FioString%></textarea>                                
        &nbsp;</td>                                
    </tr>                         
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1"  class="tableBorder" >
    <tr class="tdbg"> 
      <td width="150" class="tdbg"><strong> 替换：</strong></td>
      <td class="tdbg"><textarea name="FilterRep" cols="70" rows="5" class="lostfocus" id="FilterRep"><%=FilterRep%></textarea></td>
    </tr>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1"  class="tableBorder" >
    <tr class="tdbg"> 
      <td colspan="2" align="center" class="tdbg">
        <input name="Saveok" type="hidden" id="Saveok" value="ok">
		<input name="radiobutton" type="hidden" id="Action" value="<%=lx%>"> 
        <input name="FilterID" type="hidden" id="FilterID" value="<%=FilterID%>">
        <input name="Cancel" type="button" id="Cancel" value="返&nbsp;&nbsp;回" onClick="window.location.href='sk_ItemFilters.asp'" style="cursor: hand;background-color: #cccccc;">
        &nbsp;
        <input  type="submit" name="Submit" value="确&nbsp;&nbsp;定" style="cursor: hand;background-color: #cccccc;"> 
      </td>
    </tr>
</table>
</form>
         
</body>         
</html>
<%End Sub%>

<%
Sub Save_Data
FilterID=Trim(Request("FilterID"))
FilterName=Trim(Request.Form("FilterName"))
ItemID=Trim(Request.Form("ItemID"))
FilterObject=Trim(Request.Form("FilterObject"))
FilterType=Trim(Request.Form("FilterType"))
FilterContent=Request.Form("FilterContent")
FisString=Request.Form("FisString")
FioString=Request.Form("FioString")
FilterRep=Request.Form("FilterRep")
Flag=Trim(Request.Form("Flag"))
'PublicTf=Trim(Request.Form("PublicTf"))
If Action<>"edit" and  Action<>"add" then exit sub
If Action="edit" then 
	If FilterID="" Then                                
	   FoundErr=True                                
	   ErrMsg=ErrMsg & "<br><li>参数错误！请从有效链接进入</li>"
	Else
	   FilterID=Clng(FilterID)                                 
	End If                                
End If
If FilterName="" Then                                
   FoundErr=True                                
   ErrMsg=ErrMsg & "<br><li>过滤名称不能为空</li>"                                 
End If                                
If ItemID="" Then                                
   FoundErr=True                                
   ErrMsg=ErrMsg & "<br><li>请选择过滤所属项目</li>"                                 
Else                                
   ItemID=Clng(ItemID)  
   If ItemID=0 Then  
      FoundErr=True                                
      ErrMsg=ErrMsg & "<br><li>请选择过滤所属项目</li>"   
   End If                                
End If       
If FilterObject="" Then                                
   FoundErr=True                                
   ErrMsg=ErrMsg & "<br><li>请选择过滤对象</li>"                                 
Else                                
   FilterObject=Clng(FilterObject)   
End If                                
                            
If FilterType="" Then                                
   FoundErr=True                                
   ErrMsg=ErrMsg & "<br><li>请选择过滤类型</li>"                                 
Else                                
   FilterType=Clng(FilterType)                                
   If FilterType=1 Then     
      If FilterContent="" Then                              
         FoundErr=True                                
         ErrMsg=ErrMsg & "<br><li>过滤的内容不能为空</li>"   
      End If   
   ElseIf FilterType=2 Then   
      If FisString="" or FioString="" Then   
         FoundErr=True                                
         ErrMsg=ErrMsg & "<br><li>开始/结束标记不能为空</li>"   
      End If   
   Else   
      FoundErr=True                                
      ErrMsg=ErrMsg & "<br><li>参数错误，请从有效链接进入</li>"   
   End If   
End If                     
If FoundErr<>True Then     
   if Action="add" then                            
   		SqlItem ="select top 1 *  from Filters"
   else 
   		SqlItem ="select top 1 *  from Filters Where FilterID=" & FilterID                                
   end if
   Set RsItem=server.CreateObject("adodb.recordset")                                
   RsItem.open SqlItem,ConnItem,2,3
   if Action="add" then RsItem.addnew                                           
   RsItem("FilterName")=FilterName                                
   RsItem("ItemID")=ItemID   
   RsItem("FilterObject")=FilterObject                                
   RsItem("FilterType")=FilterType                                
   If FilterType=1 Then                                
      RsItem("FilterContent")=FilterContent                                
   ElseIf FilterType=2 Then                                
      RsItem("FisString")=FisString                                
      RsItem("FioString")=FioString                                
   End If                                                
   RsItem("FilterRep")=FilterRep        
   RsItem("Flag")=Flag
   RsItem("colleclx")=lx  
   'RsItem("PublicTf")=PublicTf                        
   RsItem.Update                                
   RsItem.Close                                
   Set RsItem=Nothing 
   	Response.Redirect "SK_ItemFilters.asp?radiobutton=" &lx                         
Else                                
   Call WriteErrMsg(ErrMsg)                                
End If                  
End Sub

Sub Test
Set RsItem=server.createobject("adodb.recordset")         
SqlItem="select * from Filters Where FilterID=" & FilterID         
RsItem.open SqlItem,ConnItem,1,1      
If Not RsItem.Eof Then
   ItemID=RsItem("ItemID")
   FilterName=RsItem("FilterName")
   FilterObject=RsItem("FilterObject")
   FilterType=RsItem("FilterType")
   FilterContent=RsItem("FilterContent")
   FisString=RsItem("FisString")
   FioString=RsItem("FioString")
   FilterRep=RsItem("FilterRep")
   Flag=RsItem("Flag")
   'PublicTf=RsItem("PublicTf")
Else
   FoundErr=True
   ErrMsg=ErrMsg & "<br></li>参数错误，找不到该项目</li>"
End If
RsItem.Close
Set RsItem=Nothing
End Sub
%>