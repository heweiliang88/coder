<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../../inc/charsetasp.asp"-->
<%
'采集内核：SK采集系统
'QQ81962480 天人文章管理系统
%>
<!doctype html>
<html>
<head>
<title>天人文章管理系统采集模块</title>
<meta charset="utf-8">
<!--#include file="inc/setup.asp"-->

<%
If InStr(","&adminlimitidrange&",", ",5,")<1 Then contrl_message "您没有访问权限，请尝试其他版块。", "", "", "", ""

Dim Sql,SqlItem,RsItem,Action,FoundErr
Dim HistrolyID,ItemID,ChannelID,ClassID,SpecialID,ArticleID,Title,CollecDate,NewsUrl,Result
Dim  Arr_Histroly,Arr_ArticleID,i_Arr,Del,Flag
Dim MaxPerPage,CurrentPage,AllPage,HistrolyNum,i_His
MaxPerPage=20
FoundErr=False
Del=Trim(Request("Del"))
Action=Trim(Request("Action"))
If (action<>"" or Del<>"") And adminctrllimit = 1 Then Call contrl_message ("您只有查看权限，没有管理权限，请升级权限后继续操作。", "", "", "", "")
If Del="Del" Then
   Call DelHistroly()
End If
If FoundErr<>True Then
   Call Main()
else
  Call WriteErrMsg(ErrMsg)
End If
'关闭数据库链接
Call CloseConnItem()
%>
<%Sub Main%>
<link rel="stylesheet" type="text/css" href="../skin/default/style.css">
<script type="text/javascript" src="../js/hover.js"></script>
<link rel="shortcut icon" type="image/x-icon" href="../../favicon.ico">
<link rel="bookmark" type="image/x-icon" href="../../favicon.ico">
<style type="text/css">
.ButtonList { BORDER-RIGHT: #000000 2px solid; BORDER-TOP: #ffffff 2px solid; BORDER-LEFT: #ffffff 2px solid; CURSOR: default; BORDER-BOTTOM: #999999 2px solid; BACKGROUND-COLOR: #e6e6e6 }
.textbox{ width:50px; height:30px; line-height:30px; overflow:hidden; padding:0 5px; border:1px solid #BFE2E6; background:#86ECEE; margin-right:8px;}
</style>
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
</script>
</head>
<body style="background:#fff; " class="w1">
<%                             
Set RsItem=server.createobject("adodb.recordset")         
SqlItem="select * from Histroly"
If Action="Succeed"  Then
   SqlItem=SqlItem  &  " Where  Result=True"
   Flag="成功记录"
ElseIf Action="Failure"  Then
   SqlItem=SqlItem  &  " Where  Result=False"
   Flag="失败记录"
ElseIf  Action="LoseEf"  Then
   Flag="失效记录"
   Set Rs=server.createobject("adodb.recordset")
   Sql="Select ArticleID From SK_Article Where Deleted=False"
   Rs.open  Sql,ConnItem,1,1
   If (Not Rs.Eof) And (Not Rs.Bof) Then
      Do While not rs.eof
         Arr_ArticleID=Arr_ArticleID & "," & CStr(rs("ArticleID"))
      Rs.MoveNext
      Loop
   End  If
   Rs.Close
   Set Rs=Nothing
   If Arr_ArticleID<>"" Then
      Arr_ArticleID=0 & Arr_ArticleID
   Else
      Arr_ArticleID=0
   End If
   SqlItem=SqlItem  &  " Where ArticleID Not In("  &  Arr_ArticleID & ")"
Else
   Flag="所有记录"
End  If
%>
<table width="100%" align="center" border="0" cellpadding="3" cellspacing="1" id="mytable" class="size3" >
  <tr class="tr1" >
    <td class="td_title" colspan="17"><!--#include file="../incgoback.asp"-->■ 历史记录 [<%=Flag%>] | <a href="Sk_ItemHistroly.asp">管理首页</a>&nbsp;&nbsp;<a href="Sk_ItemHistroly.asp?Action=Succeed">成功记录</a>&nbsp;&nbsp;<a href="Sk_ItemHistroly.asp?Action=Failure">失败记录</a>&nbsp;&nbsp;<a href="Sk_ItemHistroly.asp?Action=LoseEf">失效记录</a></td>
  </tr>
  <form name="form1" method="post" action="?">
    <tr class="tdbg" style="padding: 0px 2px;">
      <td width="56" height="22" align="center"  >选择</td>
      <td width="93" height="22" align="center"  >操作</td>
      <td width="214" align="center"  >项目名称</td>
      <td width="435" align="center" >原文标题</td>
      <td width="120" height="22" align="center"  >栏目</td>
      <td width="87" align="center"  >状态</td>
    </tr>
    <%
If Request("page")<>"" then
    CurrentPage=Cint(Request("Page"))
Else
    CurrentPage=1
End if 
SqlItem=SqlItem  &  " order by HistrolyID DESC"
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
    <tr class="tdbg" >
      <td width="56" align="center" ><input type="checkbox" value="<%=RsItem("HistrolyID")%>" name="HistrolyID" onClick="unselectall(this.form)" ></td>
      <td width="93" align="center" ><a href="Sk_ItemHistroly.asp?Action=<%=Action%>&Del=Del&HistrolyID=<%=RsItem("HistrolyID")%>" onclick='return confirm("确定要删除此记录吗？");'>删除</a></td>
      <td width="214" align="center" ><%Call Admin_ShowItem_Name(RsItem("ItemID"))%></td>
      <td width="435" align="left" ><a href="<%=RsItem("NewsUrl")%>" target="_blank" title="<%=RsItem("NewsUrl")%>"><%=RsItem("Title")%></a></td>
      <td width="120" align="center" ><%Call Admin_ShowClass_Name(RsItem("ChannelID"),RsItem("ClassID"))%></td>
      <td width="87" align="center" ><%If RsItem("Result")=True Then
           Response.write "成功"
        ElseIf RsItem("Result")=False Then
           Response.Write "<font color=red>失败</font>"
        Else
           Response.Write "<font color=red>异常</font>"
        End If
      %></td>
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
      <td colspan="8" height="30"><input name="Del" type="hidden" id="Del" value="Del">
        <input name="Action" type="hidden" id="Action" value="<%=Action%>">
        <label>
          <input name="chkAll" type="checkbox" id="chkAll" onclick="CheckAll(this.form)" value="checkbox" >
          全选</label></td>
    </tr>
    <tr class="tdbg">
      <td colspan="8" height="30" align="center"><input type="submit" value="清除选中记录" name="DelFlag" onclick='return confirm("确定要清除所选记录吗？");' class="button4">
        &nbsp;&nbsp;
        <input type="submit" value="清除失败记录" name="DelFlag" onclick='return confirm("确定要清除所有失败记录吗？");' class="button4" >
        &nbsp;&nbsp;
        <input type="submit" value="清除失效记录" name="DelFlag"  onclick='return confirm("确定要清除所有失效记录吗？");' class="button3">
        &nbsp;&nbsp;
        <input type="submit" value="清空所有记录" name="DelFlag" onclick='return confirm("确定要清除所有记录吗？");' class="button3"></td>
    </tr>
    <%Else%>
    <tr class="tdbg">
      <td colspan='9' class="tdbg" align="center">
        系统中暂无历史记录！</td>
    </tr>
    <%End  If%>
    <%         
RsItem.Close         
Set RsItem=nothing           
%>
  </form>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="tableBorder" >
  <tr>
    <td height="22" colspan="2" class="tdbg"><%
Response.Write ShowPage("Sk_ItemHistroly.asp?Action="& Action,HistrolyNum,MaxPerPage,True,True," 个记录 ")
%></td>
  </tr>
</table>
</body>
</html>
<%End Sub%>
<%Sub DelHistroly

Dim DelFlag
DelFlag=Trim(Request("DelFlag"))
HistrolyID=Trim(Request("HistrolyID"))
If HistrolyID<>"" Then
   HistrolyID=Replace(HistrolyID," ","")
End If
If DelFlag="清除选中记录" Then
   If HistrolyID="" Then
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>请选择要删除的记录</li>"
   Else
      HistrolyID=Replace(HistrolyID," ","")
      SqlItem="Delete From [Histroly] Where HistrolyID in(" & HistrolyID & ")"
   End If
ElseIf DelFlag="清除失败记录" Then
   SqlItem="Delete From [Histroly] Where Result=False"
ElseIf DelFlag="清除失效记录" Then
   Set Rs=server.createobject("adodb.recordset")
   Sql="Select ArticleID From SK_Article Where Deleted=False"
   Rs.open Sql,ConnItem,1,1
   If (Not Rs.Eof) And (Not Rs.Bof) Then
      Do While not rs.eof
         Arr_ArticleID=Arr_ArticleID & "," & CStr(rs("ArticleID"))
      Rs.MoveNext
      Loop
   End  If
   Rs.Close
   Set Rs=Nothing
   If Arr_ArticleID<>"" Then
      Arr_ArticleID=0 & Arr_ArticleID
      SqlItem="Delete From [Histroly] Where ArticleID Not In(" & Arr_ArticleID & ")"
   Else
      SqlItem="Delete From [Histroly]"
   End If
ElseIf DelFlag="清空所有记录" Then
   SqlItem="Delete From [Histroly]"
Else
   If HistrolyID="" Then
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>请选择要删除的记录</li>"
   Else
      HistrolyID=Replace(HistrolyID," ","")
      SqlItem="Delete From [Histroly] Where HistrolyID In(" & HistrolyID & ")"
   End If
End if

If FoundErr<>True Then
   ConnItem.Execute(SqlItem)
End If
End Sub
%>
<script type="text/javascript" >
showtable("mytable")
</script>