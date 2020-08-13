<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../../inc/charsetasp.asp"-->
<%
'采集内核：SK采集系统
'QQ81962480 天人文章管理系统
%>
<!--#include file="inc/setup.asp"-->
<!--#include file="inc/cj_cls.asp"-->
<html>
<head>
<title>SK采集系统---定时采集</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" type="text/css" href="../skin/default/style.css">
<script LANGUAGE="JavaScript"> 
<!-- 
function openwin() { 
window.open ("sk_Timing.asp?action=GoTiming", "newwindow", "height=400, width=500, top=0, left=0, toolbar=no, menubar=no, scrollbars=no, resizable=yes, location=no, status=no")
} 
--> 
</script>
<%
If InStr(","&adminlimitidrange&",", ",5,")<1 Then contrl_message "您没有访问权限，请尝试其他版块。", "", "", "", ""
dim action
action=Trim(Request("action"))
If action<>"" And adminctrllimit = 1 Then Call contrl_message ("您只有查看权限，没有管理权限，请升级权限后继续操作。", "", "", "", "")
If action="SaveTiming" then
   Call SaveTiming()
Else
	If action="GoTiming" then
    	Call GoTiming()
	Else
		Call main()	
	End if
End if
'关闭数据库链接
Call CloseConnItem()

Sub main()%>
<% dim RsItem,SetWeekday_style
Set RsItem=ConnItem.execute("select top 1 * from SK_timing")
if RsItem("Timing_SetDate")=0 then SetWeekday_style="display:none"
%>
<script language="JavaScript">
<!--
var tID=0;
function ShowTabs(ID){
  if(ID!=tID){
    TabTitle[tID].className='title5';
    TabTitle[ID].className='title6';
    Tabs[tID].style.display='none';
    Tabs[ID].style.display='';
    tID=ID;
  }
}
        function ShowTimingType(num){ 
                switch(num){
                case "0":     
                        document.myform.Timing_SetWeekday.style.display='none';
                        //document.myform.Timing_SetDay.style.display='none';
                        break;
                case "1":     
                        document.myform.Timing_SetWeekday.style.display='';
                        //document.myform.Timing_SetDay.style.display='none';
                        break;
                default:
                        alert("错误参数调用！");
                        break;
                }
        }
//-->
</script>
<br>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="tableBorder">
  <tr>
    <td height="38" align="center" bgcolor="#F5F5F5"><strong>定时采集设置</strong></td>
  </tr>
  <tr>
  
  <form name='myform' method='post' action='SK_Timing.asp?action=SaveTiming'>
    <td height="35" align="center" bgcolor="#F5F5F5">
        定时启动时间: 
        <select name='Timing_SetDate' id='SetTiming_Date' onChange="javascript:ShowTimingType(this.options[this.selectedIndex].value)">
          <option value='0' <% if RsItem("Timing_SetDate")=0 then Response.Write "selected" %>>每天</option>
          <option value='1' <% if RsItem("Timing_SetDate")=1 then Response.Write "selected" %>>每周</option>
        </select>
        <select name='Timing_SetWeekday' id='Timing_SetWeekday'  style='<%= SetWeekday_style %>'>
          <option value='1' <% if RsItem("Timing_SetWeekday")=1 then Response.Write "selected" %>>星期日</option>
          <option value='2' <% if RsItem("Timing_SetWeekday")=2 then Response.Write "selected" %>>星期一</option>
          <option value='3' <% if RsItem("Timing_SetWeekday")=3 then Response.Write "selected" %>>星期二</option>
          <option value='4' <% if RsItem("Timing_SetWeekday")=4 then Response.Write "selected" %>>星期三</option>
          <option value='5' <% if RsItem("Timing_SetWeekday")=5 then Response.Write "selected" %>>星期四</option>
          <option value='6' <% if RsItem("Timing_SetWeekday")=6 then Response.Write "selected" %>>星期五</option>
          <option value='7' <% if RsItem("Timing_SetWeekday")=7 then Response.Write "selected" %>>星期六</option>
        </select>
        <select name="Timing_Time" id="Timing_Time"> 
          <%dim i,Time_S : Time_S=CDate("00:00:00")
		 for i=1 to 48
		 if  CStr(RsItem("Timing_Time"))=CStr(Time_S) then 
			Response.Write "<option value="""& Time_S &""" selected>"& Time_S &"</option>"
		 else
		 	Response.Write "<option value="""& Time_S &""">"& Time_S &"</option>"
		 end if	
		 Time_S = CDate(Time_S) + CDate("00:30:00")
		 next  
		 %>
        </select>
        &nbsp;  &nbsp;  &nbsp; 
        <input name="Flag" type="checkbox" value="yes" <% if RsItem("Flag")=1 then Response.Write "checked"%>>
        是否包括号区域采集 
        &nbsp;  &nbsp;  &nbsp;
        <input name="Submit2" type="submit" value="保存设置">
   </td>
	</form>
  </tr>
  
  <tr>
    <td height="35" align="center" bgcolor="#F5F5F5"><input name="Submit" type="button" onClick="openwin()" class="alert" value="开始启动"></td>
  </tr>
</table>
<% end sub
SK.GeU_WUrl 
Sub SaveTiming
   dim Timing_SetDate,Timing_SetWeekday,Timing_Time,RsItem,Flag
   Timing_SetDate=Trim(Request.Form("Timing_SetDate"))
   Timing_SetWeekday=Trim(Request.Form("Timing_SetWeekday"))
   Timing_Time=Trim(Request.Form("Timing_Time"))
   Flag=Trim(Request.Form("Flag"))
   IF Timing_SetWeekday="" then Timing_SetWeekday=1
   IF Flag="yes" then  Flag=1 else Flag=0 end if
   Set RsItem=server.CreateObject("adodb.recordset")                                
   RsItem.open "select *  from SK_timing Where id=1",ConnItem,1,3
   RsItem("Timing_SetDate")=Timing_SetDate
   RsItem("Timing_SetWeekday")=Timing_SetWeekday
   RsItem("Timing_Time")=Timing_Time
   RsItem("Flag")=Flag 
   RsItem.update
   RsItem.close
   set RsItem=nothing
   Response.Write "<script>alert('设置成功');location.href='SK_timing.asp'</script>"'关闭窗口
End sub

Sub GoTiming()
	Dim Timing_SetDate,Timing_SetWeekday,Timing_Time,RsItem,Flag
	Dim Flag_1,Timing_SetDate_1,NowTiem,Collecdate
	Collecdate=Trim(Request("Collecdate"))
	NowTiem=Hour(Now()) &":"&  Minute(Now()) &":"& Second(Now()) 
	Set RsItem=ConnItem.execute("select top 1 *  from SK_timing")
	Timing_SetDate=RsItem("Timing_SetDate")
    Timing_SetWeekday=RsItem("Timing_SetWeekday")
    Timing_Time=RsItem("Timing_Time")
    Flag=RsItem("Flag")
	Select Case Timing_SetDate
	Case 0
	    If CDate(NowTiem)>=CDate(Timing_Time) Then
			If CStr(Day(now()))<>CStr(Collecdate) then
				Collecdate=Day(now())
				Call Timing_window(Collecdate)
			End If
		End if 
		Timing_SetDate_1="每天的 " & Timing_Time
	Case 1
		If Timing_SetWeekday = Weekday(Now()) And CDate(NowTiem)>=CDate(Timing_Time) Then
			If CStr(Day(now()))<>CStr(Collecdate) then
				Collecdate=Day(now())
				Call Timing_window(Collecdate)
			End If
		End if 
		Timing_SetDate_1="每周的星期" & Timing_SetWeekday &"   " & Timing_Time 
	End Select 
	If Flag=1 Then Flag_1="启动" Else Flag_1="无"  End If
	Response.Write "<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"" class=""tableBorder"">"
	Response.Write "<tr>"
	Response.Write "<td height=""30"" align=""center"" bgcolor=""#F5F5F5""><strong>定时采集设置</strong></td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td height=""100"" align=""left"" bgcolor=""#F5F5F5"">"
	Response.Write "&nbsp;&nbsp;定时区域采集：<FONT style='font-size:12px' color='red'>"& Flag_1 &"</font><br><br>"
	Response.Write "&nbsp;&nbsp;当前的日期是：<FONT style='font-size:12px' color='red'>"& now() &" </FONT><br><br>"
	Response.Write "&nbsp;&nbsp;您设定的时间是: <FONT style='font-size:12px' color='red'>"& Timing_SetDate_1 &" 执行</FONT><br></td></tr><tr></table>"
	Response.Write "<meta http-equiv=""refresh"" content=5;url=""SK_Timing.asp?Action=GoTiming&Collecdate="& Collecdate &""">"
	RsItem.close : set RsItem=nothing 
End sub

Sub Timing_window(Collecdate)'开始定时采集
Dim RsItem,ItemID
Set RsItem=ConnItem.execute("select  ItemID  from Item where Timing=1")
If RsItem.Eof then
	Response.Write "没有定时项目!"
Else
Do While Not RsItem.Eof
	ItemID=ItemID &","& RsItem(0) 
	RsItem.movenext
Loop
	ItemID=Right(ItemID,len(ItemID)-1)
	Response.Redirect "Sk_Collection.asp?ItemID="& ItemID &"&ItemNum=1&ListNum=1&NewsSuccesNum=0&NewsFalseNum=0&ImagesNumAll=0&NewsNum_i=0&Collecdate=" & Collecdate
End if
RsItem.close
Set RsItem=nothing
End Sub
%>