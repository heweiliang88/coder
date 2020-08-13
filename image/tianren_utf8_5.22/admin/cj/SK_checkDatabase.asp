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

Dim cmsrs,Sql,SqlItem,RsItem,Action,FoundErr,Num,SuccNum,ErrNum,Frs,RSNum,Inum,ii
Dim HistrolyID,ItemID,ChannelID,ClassID,SpecialID,ArticleID,Title,CollecDate,NewsUrl,Result
Dim  Arr_Histroly,Arr_ArticleID,i_Arr,Del,Flag,NewsID,DelFlag
Dim MaxPerPage,CurrentPage,AllPage,HistrolyNum,i_His,lx,radiobutton,lb,rslb

Dim oUrl, oTmp, oI
oUrl = Request.ServerVariables("PATH_INFO")
oTmp = Split(oUrl,"/")
oUrl = ""
For oI=0 To Ubound(oTmp) - 3
	oUrl = oUrl & oTmp(oI) & "/"
Next
%>
<link rel="stylesheet" type="text/css" href="../skin/default/style.css">
<script type="text/javascript" src="../js/hover.js"></script>
<link rel="shortcut icon" type="image/x-icon" href="../../favicon.ico">
<link rel="bookmark" type="image/x-icon" href="../../favicon.ico">
<style type="text/css">
.ButtonList { BORDER-RIGHT: #000000 2px solid; BORDER-TOP: #ffffff 2px solid; BORDER-LEFT: #ffffff 2px solid; CURSOR: default; BORDER-BOTTOM: #999999 2px solid; BACKGROUND-COLOR: #e6e6e6 }
.textbox{ width:50px; height:30px; line-height:30px; overflow:hidden; padding:0 5px; border:1px solid #BFE2E6; background:#86ECEE; margin-right:8px;}
</style>
<%

SuccNum=Trim(Request("SuccNum"))
ErrNum=Trim(Request("ErrNum"))
RSNum=Trim(Request("RSNum"))
Inum=Trim(Request("INum"))
HistrolyID=Trim(Request("HistrolyID"))
DelFlag=Trim(Request("DelFlag"))
if SuccNum="" or ErrNum="" or RSNum="" or Inum="" then
SuccNum=0 : ErrNum=0 : RSNum=0 : Inum=0
end if
MaxPerPage=20
FoundErr=False
Action=LCase(Trim(Request("Action")))
If action<>"" And adminctrllimit = 1 Then Call contrl_message ("您只有查看权限，没有管理权限，请升级权限后继续操作。", "", "", "", "")
lx=Request("radiobutton")
if lx="" or lx=0 then lx=1
'输出图片地址
if Trim(Request("Urlsc"))="ok" then
	if lx=3 then Set FRS = ConnItem.execute("select * from SK_photo")
	if lx=5 then Set FRS = ConnItem.execute("select * from SK_DownLoad")
	if lx=3  then
		while not FRS.eof
		   PhotoUrl= PhotoUrl & vbcrlf & frs("PhotoUrl")
			PicUrls=Split(frs("PicUrls"),"|||")
			for i=0 to Ubound(PicUrls)
			 pic_temp=Replace(PicUrls(i),"图片" & i+1 &"|","")
			 pic_1 = pic_1 & vbcrlf & pic_temp
			next 
			pic_2 = pic_2 & vbcrlf & pic_1
			pic_1=""
			Frs.movenext
			SuccNum=SuccNum+1
		wend
	end if
	Response.Write PhotoUrl
	Response.Write pic_2
	if lx =3 then call FSOSaveFile(PhotoUrl & pic_2,"photo.txt")
	if lx =5 then call FSOSaveFile(PhotoUrl & pic_2,"soft.txt")
	Frs.close
	set frs=nothing
	response.write  "<script>alert('提示:本次共输出 " & SuccNum &  " 条地址;');location.href='sk_checkdatabase.asp?radiobutton="& lx &"'</script>"  
		
end if 
select case Action
case "ok"
	select case DelFlag
	case "审核入库所选记录"
	If HistrolyID="" Then
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>请选择要发布的记录</li>"
	Else
		Call OpenConn()
      HistrolyID=Replace(HistrolyID," ","")
	  	 if lx=1 then	
			 Set FRS = Server.CreateObject("ADODB.RECORDSET")
			  FRS.Open "select * from SK_Article Where ArticleID in(" & HistrolyID & ")", ConnItem, 1, 3
			  If Not FRS.EOF Then
				  while not FRS.eof
				  InsertIntoBase 1,FRS
				  FRS.movenext
				  wend
			  End if
			  ConnItem.execute("Delete From SK_Article Where ArticleID in(" & HistrolyID & ")")
			  FRS.close
			  set FRS=nothing
		 end if
		Call CloseConn()
		  Call NumMsg()
	End if 
	case "审核入库全部记录"
		Call OpenConn()
		select case lx
		case 1
			SQLstr="select * from SK_Article order by ArticleID DESC"
		case 3
			SQLstr="select  * from SK_photo order by ID DESC"
		case 5
			SQLstr="select * from SK_download order by ID DESC"
		end select
		if SQLstr<>"" then
			Set FRS = Server.CreateObject("ADODB.RECORDSET")
			FRS.Open SQLstr, ConnItem, 1, 3
			  If Not FRS.EOF Then
				  while not FRS.eof
				  	if lx=1 then InsertIntoBase 1,FRS
					if lx=3 then InsertIntoBase_photo FRS 
					if lx=5 then InsertIntoBase_down FRS
					FRS.movenext
				  wend 
		 	  end if	  
				if lx=1 then ConnItem.execute("Delete From SK_Article")
				if lx=3 then ConnItem.execute("Delete From sk_photo")
				if lx=5 then ConnItem.execute("Delete From sk_download")
		 	  FRS.close
		 	  set FRS=nothing
		End if
		
		Call CloseConn()
		
		Response.Write  "<script>alert('提示:本次共操作 " & SuccNum + ErrNum & " 篇文章\n其中成功入库 " & SuccNum & " 篇,重复而不允许入库 "&ErrNum & " 篇;');location.href='sk_checkdatabase.asp?radiobutton="& lx &"'</script>"  
	case "删除所选记录"
		if HistrolyID<>""  then 
			if lx=1 then ConnItem.execute("Delete From SK_Article Where ArticleID in(" & HistrolyID & ")")
		end if
		If Request("page")<>"" then
			CurrentPage=Cint(Request("Page"))
		Else
			CurrentPage=1
		End if 
		If CurrentPage=0 then CurrentPage=1
		response.write "<meta http-equiv=""refresh"" content=""0;url=sk_checkdatabase.asp?radiobutton="& lx &"&page="& CurrentPage &""">"
	case "删除全部记录"
		if lx=1 then ConnItem.execute("Delete From SK_Article")
		response.write "<meta http-equiv=""refresh"" content=""0;url=sk_checkdatabase.asp?radiobutton="& lx &""">"
		response.end
	end select
case "del"
		Response.Flush()
		if HistrolyID<>""  then 
			Select Case lx
			Case 1
				ConnItem.execute("Delete From SK_Article Where ArticleID in(" & HistrolyID & ")")
			End Select	
		End if
		If Request("page")<>"" then
			CurrentPage=Cint(Request("Page"))
		Else
			CurrentPage=1
		End if 
		If CurrentPage=0 then CurrentPage=1
		response.write "<meta http-equiv=""refresh"" content=""0;url=sk_checkdatabase.asp?radiobutton="& lx &"&page="& CurrentPage &""">"
Case else
    call top()
	select case lx
	case 1
	   Call Main1()'新闻
	case else
	   Call Main1()
	end select
end select

if FoundErr=True then  Call WriteErrMsg(ErrMsg)
sub NumMsg()
		   response.write  "<script>alert('提示:本次共操作 " & SuccNum + ErrNum & " 篇文章\n其中成功入库 " & SuccNum & " 篇,重复而不允许入库 "&ErrNum & " 篇;');</script>" 
		   response.write "<meta http-equiv=""refresh"" content=""0;url=sk_checkdatabase.asp?radiobutton="& lx &""">"	   
end sub
%>
<%sub top()%>
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
<body style="background:#fff; " class="w1">
			<%
%>
<%end sub %>
<%
'---------文章------------------
Sub Main1
%>
<!--列表-->
<table width="100%" align="center" border="0" cellpadding="3" cellspacing="1" id="mytable" class="size3" >
  <tr class="tr1" >
    <td class="td_title" colspan="17"><!--#include file="../incgoback.asp"-->■ 已采集数据管理 </td>
  </tr>
	<form name="form1" method="post" action="sk_checkDatabase.asp">
		<tr class="tdbg" style="padding: 0px 2px;">
			<td width="57" height="22" align="center" >选择</td>
			<td width="90" height="22" align="center" >操作</td>
			<td width="142" align="center" >文章来源</td>
			<td width="358" align="center" >新闻标题</td>
			<td width="110" height="22" align="center" >栏目</td>
		</tr>
		<%                          
Set RsItem=server.createobject("adodb.recordset")
SqlItem="select * from SK_Article" 
If Request("page")<>"" then
    CurrentPage=Cint(Request("Page"))
Else
    CurrentPage=1
End if 

SqlItem=SqlItem  &  " order by ArticleID DESC"
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
			<td width="57" align="center"><input type="checkbox" value="<%=RsItem("ArticleID")%>" name="HistrolyID" onClick="unselectall(this.form)" ></td>
			<td width="90" align="center"><a href="sk_checkDatabase.asp?Action=Del&Page=<%= CurrentPage %>&HistrolyID=<%=RsItem("ArticleID")%>&radiobutton=<%=lx%>" onclick='return confirm("确定要删除此记录吗？");'>删除</a></td>
			<td width="142" align="center"><%=left(RsItem("CopyFrom"),8)%></td>
			<td width="358" align="left"><a href="show.asp?id=<% = RsItem("ArticleID") %>&lx=1" target="_blank"><%=RsItem("Title")%></a></td>
			<td width="110" align="center"><%Call Admin_ShowClass_Name(RsItem("ChannelID"),RsItem("ClassID"))%></td>
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
			<td colspan="8" height="30"><input name="Action" type="hidden" id="Action" value="ok">
				<input name="radiobutton" type="hidden" id="Action" value="<%=lx%>">
				<input name="page" type="hidden" value="<%=CurrentPage %>">
				<label><input name="chkAll" type="checkbox" id="chkAll" onclick="CheckAll(this.form)" value="checkbox" >
				全选</label> </td>
		</tr>
		<tr class="tdbg">
			<td colspan="8" height="30" align="center"><input name="DelFlag" type="submit" class="button4" onclick='return confirm("为避免重复入库后所选数据将删除，确定要审核入库所选记录吗？");' value="审核入库所选记录">&nbsp;&nbsp;&nbsp;&nbsp;
				<input name="DelFlag" type="submit" class="button4" onclick='return confirm("为避免重复入库后数据将删除，确定要审核入库全部记录吗？");' value="审核入库全部记录">
				&nbsp;&nbsp;&nbsp;&nbsp;
				<input name="DelFlag" type="submit" class="button3" onclick='return confirm("确定要删除选中的记录吗？");' value="删除所选记录">
				&nbsp;&nbsp;&nbsp;&nbsp;
				<input name="DelFlag" type="submit" class="button3" onclick='return confirm("确定要删除所有的记录吗？");' value="删除全部记录">
				&nbsp;&nbsp; </td>
		</tr>
		<tr class="tdbg">
			<td colspan="8" height="30">小提示：为避免重复录入数据，入库后的数据自动删除！ </td>
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
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="tableBorder" >
	<tr>
		<td height="22" colspan="2" class="tdbg"><%
Response.Write ShowPage("sk_checkDatabase.asp?Action="& Action,HistrolyNum,MaxPerPage,True,True," 个记录 ")
%></td>
	</tr>
</table>

<!--列表-->
</body>
</html>
<%End Sub
'---------文章------------------
%>
<%Sub InsertIntoBase(collectlx,FRS)
Response.Flush()
Dim oMaxID
set rslb=Conn.Execute("select top 1 * from tr_column Where ID=" & FRS("ClassID")) 
if not rslb.eof then
	   Set CMSRS=Server.CreateObject("ADODB.RECORDSET")
	   CMSRS.Open "Select top 1 * From tr_Article Where Title='" & FRS("Title") & "' And columnid=" & FRS("ClassID") & " and isdel=0 ",Conn,1,3
	   If CMSRS.EOF Then
	   	   CMSRS.addnew
		   
		   	oMaxID = Conn.Execute("select max(ID) from tr_Article")(0)
			If IsNull(oMaxID) Then
				oMaxID = 0
			End If
			
		   	CMSRS("ID") =  oMaxID+1
			CMSRS("ChannelID") = 1
			CMSRS("columnid") = FRS("ClassID")
			CMSRS("queueid") = Conn.Execute("select queueid from tr_column where ID=" & FRS("ClassID"))(0)
			CMSRS("columnname") = Conn.Execute("select colname from tr_column where ID=" & FRS("ClassID"))(0)
			CMSRS("Title") = FRS("Title")
			CMSRS("TitleColor") = ""
			CMSRS("jumpurl") = ""
			CMSRS("Keywords") = FRS("Keyword")
			CMSRS("Descriptions") = ""
			CMSRS("pgcount") = UBound(Split(FRS("Content"), "[NextPage]")) + 1
			CMSRS("Content") = Replace(Replace(FRS("Content"), "../../upfiles/", oUrl & "upfiles/"),"[NextPage]","<55tr.com>page</55tr.com>")

			CMSRS("Author") = FRS("Author")
			CMSRS("sources") = FRS("CopyFrom")
			CMSRS("AddTime") = FRS("UpdateTime")
			CMSRS("Inputer") = "55tr"
			CMSRS("IsPass") = 1
'			CMSRS("IsPic") = 0
			If Not IsNull(FRS("picpath")) Then
			CMSRS("PicFiles") = Replace(FRS("picpath"),"../../",oUrl)
			else
			CMSRS("PicFiles") = ""
			end if
			CMSRS("mpicfiles") = ""
'			If CMSRS("PicFile")<>"" Then CMSRS("IsPic") = 1
			CMSRS("isfocus") = 0
			CMSRS("isroll") = 0
			CMSRS("isindextop") = 0
			CMSRS("islisttop") = 0
			CMSRS("iscommend") = 0
			CMSRS("isrollimg") = 0
			CMSRS("isabout") = 0
			CMSRS("isdel") = 0
			CMSRS("iscontribute") = 0
			CMSRS("isimgtext") = 0
			CMSRS("isgavepoints") = 0
'			CMSRS("IsDelete") = 0
'			CMSRS("IsMove") = 0
'			CMSRS("IsPlay") = 0
'			CMSRS("IsUserAdd") = 0
			CMSRS("limitmore") = 0
			CMSRS("minuspoints") = 0
			CMSRS("limittype") = 0
			CMSRS("limitpoints") = 0
			CMSRS("contributor") = ""
			CMSRS("clicks") = FRS("Hits")
			
		   CMSRS.update
		   SuccNum = SuccNum + 1
	   Else
	       ErrNum = ErrNum + 1
	   End if
	   Inum=Inum+1
	   CMSRS.close
	   Set CMSRS = Nothing
end if
rslb.close
set rslb=nothing
end sub

'===============================================
'函数名：FSOSaveFile
'作  用：生成文件
'参  数： Content内容,路径 注意虚拟目录
'===============================================
Sub FSOSaveFile(Content, LocalFileName)
    Dim FSO, FileObj
    Set FSO = Server.CreateObject("Scripting.FileSystemObject")
    Set FileObj = FSO.CreateTextFile(Server.MapPath(LocalFileName), True) '创建文件
    FileObj.Write Content
    FileObj.Close     '释放对象
    Set FileObj = Nothing
    Set FSO = Nothing
End Sub
'Call CloseConn()
Call CloseConnItem()
%>
<script type="text/javascript" >
showtable("mytable")
</script>
