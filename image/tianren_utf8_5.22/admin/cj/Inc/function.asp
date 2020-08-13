﻿<%

'*************************************************************************************
	'函数名:DisClassTable
	'作  用:递归显示所有分类
'*************************************************************************************
Private Sub DisClassTable(ChannelId,ParentID)
	Dim opRs,cTmp,cLen,cCount
	Set opRs = Server.CreateObject("Adodb.RecordSet")
	Sql = "select ClassID,ClassName,ParentPath from SK_Class where ChannelId=" & ChannelId & " and ParentID=" & ParentID & " order by OrderID"
	opRs.Open Sql,ConnItem,0,1
	Do While Not opRs.Eof
		Response.Write("<TR class=""tdbg"">")
        Response.Write("<TD bgcolor=""#F2F2F2"" class=tdbg>" & opRs("ClassID") & "</TD>")
        Response.Write("<TD bgcolor=""#F2F2F2"" class=tdbg>")
		cTmp = Split(opRs("ParentPath"),",")
		cLen = Ubound(cTmp) - 2
		For cCount=1 To cLen
			Response.Write("│&nbsp;")
		Next
		Response.Write("├" & opRs("ClassName"))
		Response.Write("</TD>")
        Response.Write("</TR>")		
		Call DisClassTable(ChannelId,opRs("ClassID"))
		opRs.MoveNext
	Loop
	opRs.Close
	Set opRs = Nothing
End Sub	

'*************************************************************************************
	'函数名:InitClassSelectOption
	'作  用:分类ID下拉列表选择菜单
'*************************************************************************************
Private Sub InitClassSelectOption(ChannelId,ParentID,ChkID)
	Dim opRs,cTmp,cLen,cCount
	Set opRs = Server.CreateObject("Adodb.RecordSet")
	Sql = "select ClassID,ClassName,ParentPath from SK_Class where ChannelId=" & ChannelId & " and ParentID=" & ParentID & " order by OrderID"
	opRs.Open Sql,ConnItem,0,1
	Do While Not opRs.Eof
		Response.Write("<option value=""" & opRs("ClassID") & """")
		If Cint(ChkID) = opRs("ClassID") Then Response.Write(" selected=""selected""")
		Response.Write(">")
		cTmp = Split(opRs("ParentPath"),",")
		cLen = Ubound(cTmp) - 2
		For cCount=1 To cLen
			Response.Write("│&nbsp;")
		Next
		Response.Write("├" & opRs("ClassName") & "</option>")
		
		Call InitClassSelectOption(ChannelId,opRs("ClassID"),ChkID)
		opRs.MoveNext
	Loop
	opRs.Close
	Set opRs = Nothing
End Sub

'*************************************************************************************
	'函数名:SK_Class_Channel_edit
	'作  用:频道管理--修改频道
'*************************************************************************************
	Function SK_Class_Channel_edit(className,ParentPath,classID)
		dim sql
		if className="" or ParentPath="" then
			ErrMsg="<font color=red>请填写完整！</font>"
		Else
			sql = "select top 1 * from SK_class where classID=" & classID 
			Set Rs = Server.CreateObject("adodb.recordset")
			Rs.Open SQL, ConnItem, 1,3
			Rs("className")=className
			rs("ParentPath")=ParentPath
			rs.update
			set rs=nothing
			response.redirect("sk_class.asp?Colleclx="& Colleclx)
			response.end
		end if	
		SK_Class_Channel_edit=ErrMsg
	End Function
'*************************************************************************************
'函数名:SK_Class_small_add
'作  用:频道管理--增加栏目
'*************************************************************************************
	Function SK_Class_small_add(ClassName,ParentPath,ParentID)
		dim sql,Ordermax
		if ClassName="" or ParentPath="" then
		'Response.Write  ClassName & ParentPath & ("asdfsadf") & ParentID
				ErrMsg="<font color=red>请填写完整！</font>"
		Else
			set rs=ConnItem.execute("select top 1 * from SK_class where classID=" & ParentID)
			Ordermax=rs("OrderID")+1
			depth=rs("depth")+1			
			
			Set Rs=ConnItem.execute("select * from SK_class where OrderID >=" & Ordermax)
			while not rs.eof
			Response.Write rs("classid")
			ConnItem.execute("update SK_class set OrderID="& rs("OrderID")+1 &" where classid="& rs("classid") )
			rs.movenext
			wend
			
			sql = "select top 1 * from SK_class order by ClassID desc"
			Set Rs = Server.CreateObject("adodb.recordset")
			Rs.Open SQL, ConnItem, 1,3
			rs.addnew
			if Colleclx<>0 then
				rs("ChannelID")=Colleclx
			Else
				rs("ChannelID")=0
			end if
			Rs("className")=className
			rs("ParentPath")=ParentPath
			rs("ParentID")=ParentID
			rs("OrderID")=Ordermax
			rs("depth")=depth
			rs.update
			set rs=nothing
			response.redirect("sk_class.asp?Colleclx="& Colleclx)
			response.end
				
				if ConnItem.execute("select count(ClassID) from sk_Class where ClassDir='"& ClassDir &"'")(0)>0  then
				ErrMsg="<font color=red>目录以存在！请用别的目录</font>"
					Else
					sql = "insert into sk_Class (ChannelID,ClassName,Readme,ClassDir,ParentPath) values ('" & ChannelID & "','" & ClassName &"','" & Readme &"','" & ClassDir &"','" & ClassDir &"')"
					
					ConnItem.execute(sql)
					response.redirect("sk_smallclass.asp?ChannelID=" & ChannelID)
					response.end
	  			end if
		end if		
		SK_Class_small_add=ErrMsg
	End Function
'==================================================
'过程名：SK_Showclass
'作  用：显示频道栏目分类
'==================================================	
sub SK_Showclass(FolderID,ChannelID) 
if ChannelID=0 then
Response.Write"<option selected value='555555555'>5555555555555555</option>"

Else
			Dim RS,FolderName,TreeStr,ID		
			Set  RS=Server.CreateObject("ADODB.Recordset")
			FolderID = Trim(FolderID)
			If Not IsNumeric(ChannelID) Then Return
			RS.Open ("select ID,FolderName from KS_Class Where ChannelID=" & ChannelID & " AND tj=1 Order BY FolderOrder ASC"),conn, 1, 1
			Do While Not RS.EOF
			 ID = Trim(RS(0))
			 FolderName = Trim(RS(1))
				 If FolderID = ID Then
				   TreeStr = TreeStr & "<option selected value='" & ID & "'>" & FolderName & "</option>"
				 Else
				  TreeStr = TreeStr & "<option value='" & ID & "'>" & FolderName & " </option>"
				 End If
			  TreeStr = TreeStr & ReturnSubList(ID, FolderID)
			RS.MoveNext
		   Loop
		   RS.Close:Set RS = Nothing
		   if TreeStr="" then
		   Response.Write "<option value='0'> 你没设分类 </option>"
		   Else
		   Response.Write TreeStr
		   end if
		   
		   'KSCache.add ReturnTree,dateadd("n",1000000,now)
end if
end sub
'**************************************************
'函数名：ReturnSubClass
'作  用：查找并返子批量采集项目ID。
'参  数：ParentID ----父节点ID,   K ----选择项ID
'返回值：子树
'************************************************** 
Public Function ReturnSubClass(ParentID,k)
	Dim  SubRS, SpaceStr, Num,FolderName,ID,RS
	'-----------------------------------------
	IF k=1 then
		Set RS = Server.CreateObject("ADODB.RECORDSET")
		RS.Open ("Select * From Item Where  ClassID ='"& ParentID &"' order by ItemID DESC" ), connItem, 1, 1
		Do while Not Rs.Eof
			SpaceStr= SpaceStr & RS("ItemID") & ","
			Rs.Movenext
		Loop
		Rs.Close : Set Rs=Nothing
		k=0
	End If
	'-----------------------------------------
	Set SubRS = Server.CreateObject("ADODB.RECORDSET")
	SubRS.Open ("Select ClassID,ClassName,ParentID from SK_Class Where  ParentID =" & ParentID & " "), connItem, 1, 1
	Do While Not SubRS.EOF
		ID = Trim(SubRS(0))
		'-------------------------------------
		Set RS = Server.CreateObject("ADODB.RECORDSET")
		RS.Open ("Select * From Item Where  ClassID ='"& ID &"' order by ItemID DESC" ), connItem, 1, 1
		Do while Not Rs.Eof
			SpaceStr= SpaceStr & RS("ItemID") & ","
			Rs.Movenext
		Loop
		Rs.Close : Set Rs=Nothing
		'-------------------------------------
		SpaceStr= SpaceStr & ReturnSubClass(ID,k)
		SubRS.MoveNext
	Loop 
	SubRS.Close : Set SubRS=Nothing
	ReturnSubClass=SpaceStr
End Function 
'**************************************************
'函数名：ReturnSubList
'作  用：查找并返子树数据。
'参  数：ParentID ----父节点ID,   FolderID ----选择项ID
'返回值：子树
'**************************************************
Public Function ReturnSubList(ParentID, FolderID)
	  Dim SubTypeList, SubRS, SpaceStr, k, Total, Num,FolderName, ID,TJ
	  Set SubRS = Server.CreateObject("ADODB.RECORDSET")
	  SubRS.Open ("Select count(ID) AS total from KS_Class Where  TN='" & ParentID & "'"), conn, 1, 1
	  Total = SubRS("Total")
	  SubRS.Close
	  SubRS.Open ("Select ID,FolderName,TJ from KS_Class Where  TN='" & ParentID & "' Order BY FolderOrder ASC"), conn, 1, 1
	  Num = 0
	  Do While Not SubRS.EOF
	   Num = Num + 1
	   SpaceStr = ""
		TJ = CInt(SubRS(2))
		For k = 1 To TJ - 1
		  If k = 1 And k <> TJ - 1 Then
		  SpaceStr = SpaceStr & "&nbsp;&nbsp;│"
		  ElseIf k = TJ - 1 Then
			If Num = Total Then
				 SpaceStr = SpaceStr & "&nbsp;&nbsp;└ "
			Else
				 SpaceStr = SpaceStr & "&nbsp;&nbsp;├ "
			End If
		  Else
		   SpaceStr = SpaceStr & "&nbsp;&nbsp;│"
		  End If
		Next
	  ID = Trim(SubRS(0))
	  FolderName = Trim(SubRS(1))
	  If FolderID = ID Then
	   SubTypeList = SubTypeList & "<option selected value='" & ID & "'>" & SpaceStr & FolderName & "</option>"
	  Else
	   SubTypeList = SubTypeList & "<option value='" & ID & "'>" & SpaceStr & FolderName & "</option>"
	  End If
	   SubTypeList = SubTypeList & ReturnSubList(ID, FolderID)
	  SubRS.MoveNext
	 Loop
	  SubRS.Close:Set SubRS = Nothing
	  ReturnSubList = SubTypeList
End Function
'*************************************************************************************
	'函数名:GetFileID
	'作  用:生成文件ID号,6位随机+文件
	'lx=采集类型
	'参  数:无
'*************************************************************************************
	Function GetFileID(dir,filename,lx)
		GetFileID = dir + MakeRandom(6) + filename
		Exit Function
		  Dim RSC,TempUrlArray
          Set RSC=Server.CreateObject("ADODB.RECORDSET")
			 '6位随机+文件
				Do While True
				 GetFileID = dir + MakeRandom(6) + filename
				 select case lx
				 case 1
				 'RSC.Open "Select all from sk_photo Where PicUrls='" & GetFileID & "'", ConnItem, 1, 1
				 case 2
				 'RSC.Open "Select PicUrls from sk_photo Where PicUrls='" & GetFileID & "'", ConnItem, 1, 1
				 case 3
				 RSC.Open "Select PicUrls from sk_photo Where PicUrls LIKE '%" & GetFileID & "%'", ConnItem, 1, 1
				 case 6
				 RSC.Open "Select FileUrls from sk_all Where PicUrls LIKE '%" & GetFileID & "%'", ConnItem, 1, 1
				 end select
				 
				  If RSC.EOF And RSC.BOF Then
				  Exit Do
				  End If
				Loop
		RSC.Close
		Set RSC = Nothing
	End Function
'*************************************************************************************
	'函数名:GetClassID
	'作  用:生成新目录或频道的ID号
	'参  数:无
'*************************************************************************************
	Function GetClassID()
		  Dim RSC
          Set RSC=Server.CreateObject("ADODB.RECORDSET")
			 '生成目录ID 年+10位随机
				Do While True
				 GetClassID = Year(Now()) & MakeRandom(10)
				 RSC.Open "Select ID from KS_Class Where ID='" & GetClassID & "'", ConnItem, 1, 1
				  If RSC.EOF And RSC.BOF Then
				  Exit Do
				  End If
				Loop
		RSC.Close
		Set RSC = Nothing
	End Function
	'**************************************************
	'函数名：GetFileName
	'作  用：构造文件名。
	'参  数：ArticleFsoType  ----生成类型
	'        addDate   -----添加时间,GetFileNameType--扩展名
	'**************************************************
	Public Function GetFileName(ArticleFsoType, AddDate, GetFileNameType)
		Dim N
		Randomize
	 Select Case ArticleFsoType
	  Case 1
	  GetFileName = Year(AddDate) & "/" & Month(AddDate) & "-" & Day(AddDate) & "/" & GetFileNameType  '年/月-日/随机数+扩展名
	  Case 2
	  GetFileName = Year(AddDate) & "/" & Month(AddDate) & "/" & Day(AddDate) & "/" & GetFileNameType '年/月/日/随机数+扩展名
	  Case 3
	  GetFileName = Year(AddDate) & "-" & Month(AddDate) & "-" & Day(AddDate) & "/" & GetFileNameType '年-月-日/随机数+扩展名
	  Case 4
	  GetFileName = Year(AddDate) & "/" & Month(AddDate) & "/" & GetFileNameType '年/月/随机数+扩展名
	  Case 5
	  GetFileName = Year(AddDate) & "-" & Month(AddDate) & "/"  & GetFileNameType '年-月/随机数+扩展名
	  Case 6
	  GetFileName = Year(AddDate) & Month(AddDate) & Day(AddDate) & "/"  & GetFileNameType '年月日/随机数+扩展名
	  Case 7
	  GetFileName = Year(AddDate) & "/"  & GetFileNameType '年/随机数+扩展名
	  Case 8
	  GetFileName = Year(AddDate) & Month(AddDate) & Day(AddDate)  & GetFileNameType '年+月+日+随机数+扩展名
	  Case 9
	  GetFileName =  GetFileNameType
	  Case 10
	  GetFileName =  GetFileNameType '随机字符
	  Case Else
	   GetFileName = Year(AddDate) & Month(AddDate) & Day(AddDate) & GetFileNameType '12位随机数+扩展名
	End Select
	End Function
	'*************************************************************************************
	'函数名:GetInfoID_CMS
	'作  用:生成文章,图片或下载等的唯一ID
	'参  数:ChannelID--频道ID
	'*************************************************************************************
	Public Function GetInfoID_CMS(ChannelID)
	   On Error Resume Next
	   Dim RSC, TableNameStr
       Set RSC=Server.CreateObject("ADODB.RECORDSET")
	   Select Case ChannelID
		 Case 1
		   TableNameStr = "Select NewsID From KS_Article Where NewsID='"
		 Case 2
		   TableNameStr = "Select PicID From KS_Photo Where PicID='"
		 Case 3
		   TableNameStr = "Select DownID From KS_DownLoad Where DownID='"
		 Case 4
		   TableNameStr = "Select FlashID From KS_Flash Where FlashID='"
	   End Select
	   
	   Do While True
		 GetInfoID_CMS = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Now(), "-", ""), " ", ""), ":", ""), "PM", ""), "AM", ""), "上午", ""), "下午", "") & MakeRandom(3)
			RSC.Open TableNameStr & GetInfoID_CMS & "'", conn, 1, 1
				  If RSC.EOF And RSC.BOF Then Exit Do
	   Loop
		RSC.Close:Set RSC = Nothing
	End Function
	'*************************************************************************************
	'函数名:GetInfoID
	'作  用:生成文章,图片或下载等的唯一ID
	'参  数:ChannelID--频道ID
	'*************************************************************************************
	Function GetInfoID(ChannelID)
	   On Error Resume Next
	   Dim RSC, TableNameStr
       Set RSC=Server.CreateObject("ADODB.RECORDSET")
	   Select Case ChannelID
		 Case 1
		    TableNameStr = "Select ArticleID From SK_Article Where ArticleID='"
		 Case 2
		    TableNameStr = "Select FlashID From sk_Flash Where FlashID='"
		 Case 3
			TableNameStr = "Select PicID From sk_Photo Where PicID='"
	   End Select
	   Do While True
		 GetInfoID = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Now(), "-", ""), " ", ""), ":", ""), "PM", ""), "AM", ""), "上午", ""), "下午", "") & MakeRandom(3)
			RSC.Open TableNameStr & GetInfoID & "'", ConnItem, 1, 1
				If RSC.EOF And RSC.BOF Then
					Exit Do
				End If
	   Loop
		RSC.Close
		Set RSC = Nothing
	End Function
'=============================================================================
'参  数： Tvar  ----日期   sType---选择显示类型
'返回值：成功返回随机数
'作  用：时间格式处理, 
'==============================================================================
	Public Function Format_Time(Tvar,sType)
        dim Tt,sYear,sMonth,sDay,sHour,sMinute,sSecond
        If Not IsDate(Tvar) Then Format_Time = "" : Exit Function
        Tt			= Tvar
        sYear		= Year(Tt)
        sMonth		= Right("0" & Month(Tt),2)
		sDay		= Right("0" & Day(Tt),2)
		sHour		= Right("0" & Hour(Tt),2)
		sMinute		= Right("0" & Minute(Tt),2)
		sSecond		= Right("0" & Second(Tt),2)
        Select Case sType
        Case 1	'2005-10-01 23:45:45
			Format_Time = sYear & "-" & sMonth & "-" & sDay & " " & sHour & ":" & sMinute & ":" & sSecond
        Case 2	'年-月-日 时:分:秒
			Format_Time = sYear & "年" & sMonth & "月" & sDay & "日 " & sHour & "时" & sMinute & "分" & sSecond & "秒"
        Case 3	'10-01 23:45
			Format_Time = sMonth & "-" & sDay & " " & sHour & ":" & sMinute
        Case 4	'2005-10-01
			Format_Time = sYear & "-" & sMonth & "-" & sDay
        Case 5	'2005年10月01日
			Format_Time = sYear & "年" & sMonth & "月" & sDay & "日"
        Case 6	'10-01
			Format_Time = sMonth & "-" & sDay
        Case 7	'20051001234545
			Format_Time = sYear & sMonth & sDay & sHour & sMinute & sSecond		
		Case 8	'20051001234545
			Format_Time = sYear & sMonth & sDay & sHour & sMinute
        Case Else
			Format_Time = Tt
        End Select
    End Function
'===================================================
'小金通用采集系统
'函数名：MakeRandom
'作  用：生成指定位数的随机数
'参  数： maxLen  ----生成位数
'返回值：成功返回随机数
'===================================================
	Function MakeRandom(ByVal maxLen)
	
	  Dim strNewPass
	  Dim whatsNext, upper, lower, intCounter
	  Randomize
	
	 For intCounter = 1 To maxLen
	   upper = 57
	   lower = 48
	   strNewPass = strNewPass & Chr(Int((upper - lower + 1) * Rnd + lower))
	 Next
	   MakeRandom = strNewPass
	End Function
'===================================================
'小金通用采集系统
'过程名：km_template()
'作  用：显示栏目
'===================================================
sub km_template()
	 set rs= ConnItem.execute("Select * From ks_template")
	 if rs.eof then
	 	response.write "暂无分类"
	 Else 
	
	 	Response.Write "<table width=100% border=0 cellspacing=0 cellpadding=0>"
		Response.Write "<tr><td width='40%'>ID号</td><td width='60%'>名子</td></tr>"
		do while not rs.eof
		 if rs("ChannelID")<>0 then		
			Response.Write "<tr height='22'><td><input name='km_asdfsf' type='text'value="&rs("TemplateID")&" size='15' /></td><td>&nbsp;"&rs("TemplateName")&"--模板</td></tr>"
		  end if
		rs.movenext
		loop
		Response.Write "</table>"
	 end if
	 rs.close
	 set rs=nothing
end sub	
'==================================================
'小金通用采集系统
'过程名：km_class()
'作  用：显示输出目标栏目
'==================================================
sub km_class()
	 set rs= ConnItem.execute("Select * From ks_class")
	 if rs.eof then
	 	response.write "暂无分类"
	 Else 
	
	 	Response.Write "<table width=100% border=0 cellspacing=0 cellpadding=0>"
		Response.Write "<tr><td width='40%'>ID号</td><td width='60%'>名子</td></tr>"
		do while not rs.eof
		 if rs("tn")=0 then
		Response.Write "<tr height='22'><td><font color='#FF0000'>"&rs("id")&"</font></td><td><font color='#FF0000'>&nbsp;"&rs("FolderName")&"--频道</font></td></tr>"
			Else
			Response.Write "<tr height='22'><td><input name='km_asdfsf' type='text'value="&rs("id")&" size='15' /></td><td>&nbsp;"&rs("FolderName")&"--栏目</td></tr>"
		  end if
		rs.movenext
		loop
		Response.Write "</table>"
	 end if
	 rs.close
	 set rs=nothing
end sub	 

'==================================================
'小金通用采集系统
'函数名：sk_dir_get()
'作  用：读取目录
'==================================================
Function sk_dir_get(ClassID,lx)
Dim strInstallDir,CacheTemp,rs,strtemp,strChannelDir,Arr_Path,i,PathTemp,SaveTf
strInstallDir=trim(request.ServerVariables("SCRIPT_NAME"))
strInstallDir=left(strInstallDir,instrrev(lcase(strInstallDir),"/")-1)
'strtemp = strtemp & strInstallDir 

set rs = ConnItem.execute("select top 1 ArticleDir,flashdir,photoDir from SK_Config")
if lx =1 then strtemp = strtemp & rs("ArticleDir")
if lx =2 then strtemp = strtemp  & rs("flashdir")
if lx =3 then strtemp = strtemp  & rs("photoDir") 
rs.close
set rs = ConnItem.execute("select top 1 * from SK_Class where ClassID=" & ClassID)
strtemp = strtemp & rs("ClassDir") &"/"
sk_dir_get = strtemp
rs.close
set rs=nothing
end function
'==================================================
'小金通用采集系统
'函数名：sk_dir()
'PicUrls=远程文件地址
'作  用：建立保存目录
'==================================================
Function sk_dir(ClassID,lx,FileUrl)
Dim strInstallDir,CacheTemp,rs,strtemp,strChannelDir,Arr_Path,i,PathTemp,SaveTf
strInstallDir=trim(request.ServerVariables("SCRIPT_NAME"))
strInstallDir=left(strInstallDir,instrrev(lcase(strInstallDir),"/")-1)
strtemp = strtemp & strInstallDir 

set rs = ConnItem.execute("select top 1 ArticleDir,flashdir,photoDir from SK_Config")
if lx =1 then strtemp = strtemp & rs("ArticleDir")
if lx =2 then strtemp = strtemp  & rs("flashdir")
if lx =3 then strtemp = strtemp  & rs("photoDir") 
rs.close
set rs = ConnItem.execute("select top 1 * from SK_Class where ClassID=" & ClassID)
strtemp = strtemp & rs("ClassDir") &"/"
Call SaveRemoteFile(strtemp,FileUrl)'保存远程文件
sk_dir = strtemp
rs.close
set rs=nothing
end function
'==================================================
'SK采集系统
'函数名：Sk_GetSaveDir()
'lx=类型1=新闻 2=flash 3=图片 4=音乐 5=软件 6=自定
'作  用：读取文件保存目录设置
'==================================================
Function Sk_GetSaveDir(lx)
	Dim strInstallDir,CacheTemp,rs,strtemp,strChannelDir,Arr_Path,Tempi,PathTemp,SaveTf,TempUrlArray,Ranfilestr
	strInstallDir=trim(request.ServerVariables("SCRIPT_NAME"))
	strInstallDir=left(strInstallDir,instrrev(lcase(strInstallDir),"/")-1)
	set rs = ConnItem.execute("Select top 1 Dir,MaxFileSize,FileExtName,Timeout from SK_Cj where ID=" & lx)
	strtemp = strtemp  & rs("Dir")
	Sk_GetSaveDir = strtemp & SaveFileUrl
	rs.close
	set rs=nothing
end function
'==================================================
'小金通用采集系统
'函数名：Sk_SaveFile()
'Lx=频道
'FileUrl=远程文件地址
'作  用：按设置频道保存远程文件替换地址
'==================================================
Function Sk_SaveFile(lx,FileUrl)
Dim strInstallDir,CacheTemp,rs,strtemp,strChannelDir,Arr_Path,Tempi,PathTemp,SaveTf,TempUrlArray,Ranfilestr,Ranfilestr1
strInstallDir=trim(request.ServerVariables("SCRIPT_NAME"))
strInstallDir=left(strInstallDir,instrrev(lcase(strInstallDir),"/")-1)
'strtemp = strtemp & strInstallDir 
FileUrl=replace(replace(FileUrl,"""","")," ","")
Select Case lx
Case 1
	Set rs = ConnItem.execute("Select top 1 Dir,MaxFileSize,FileExtName,Timeout from SK_Cj where ID=1")
Case 2

Case 3
	Set rs = ConnItem.execute("Select top 1 Dir,MaxFileSize,FileExtName,Timeout from SK_Cj where ID=2")
Case 5
	Set rs = ConnItem.execute("Select top 1 Dir,MaxFileSize,FileExtName,Timeout from SK_Cj where ID=3")
Case 6
End Select
strtemp = strtemp & rs("Dir")
strtemp = strtemp & SaveFileUrl

      Arr_Path=Split(strtemp,"/")
      PathTemp=""
      For Tempi=0 To Ubound(Arr_Path)
         If Tempi=0 Then
            PathTemp=Arr_Path(0) & "/"
         ElseIf Tempi=Ubound(Arr_Path) Then
            Exit For
         Else
            PathTemp=PathTemp & Arr_Path(Tempi) & "/"
         End If
         If CheckDir(PathTemp)=False Then
            If MakeNewsDir(PathTemp)=False Then
               SaveTf=False
               Exit For
            End If
         End If
      Next
TempUrlArray=Split(FileUrl,"/")  
Ranfilestr=GetFileID(strtemp,TempUrlArray(Ubound(TempUrlArray)),3)'生成文件名
'Call SaveRemoteFile(Ranfilestr,FileUrl)'保存远程文件
If SaveRemoteFile(Ranfilestr,FileUrl)<>False then'保存远程文件
Ranfilestr1=Ranfilestr
if Thumb_WaterMark=1 then call SKThumb.AddWaterMark(Ranfilestr)'水印
Sk_SaveFile = Ranfilestr1
Else
Sk_SaveFile = False
End if
Rs.close
Set Rs=nothing
End function
'==================================================
'小金通用采集系统
'过程名：SaveFile()
'作  用：远程保存-
'==================================================
Sub SaveFile()
dim Savelx,ChannelID,ClassID,id,FoundErr,ErrMsg,sql,rs,strChannelDir,pici,picii,SaveErr,TempArray,i,lx,n
Savelx=Request("Savelx")
	ChannelID=Request("ChannelID")
	ClassID=Request("ClassID")
	lx=0
		If ChannelID="" or ChannelID=0 Then
		   FoundErr=True
		   ErrMsg=ErrMsg & "<br><li>未指定频道</li>"
		Else
		   ChannelID=Clng(ChannelID)
		End If
		If ClassID="" or ClassID = 0  Then
		   FoundErr=True
		   ErrMsg=ErrMsg & "<br><li>未指定栏目</li>"
		Else
		   ClassID=CLng(ClassID)
		  set rs=ConnItem.execute("select * From SK_Class Where ClassID=" & ClassID)
		   If rs.bof and rs.eof then
				 FoundErr=True
				 ErrMsg=ErrMsg & "<br><li>找不到指定的栏目</li>"
			End If
			strChannelDir=rs("ClassDir")
			rs.close
			Set rs=Nothing


		End if
		if FoundErr=True then'错误信息
			Response.Write ErrMsg
		call Main()	
		Else
		if ConnItem.execute("select count(ArticleID) from PE_Article where ClassID="& ClassID & " and Passed=true and SaveFile<>True")(0) >0 then lx=1
		if ConnItem.execute("select count(ID) from SK_Flash where TID="& ClassID &" and Verific=1 and SaveFile=0")(0) >0 then lx=2
		if ConnItem.execute("select count(ID) from sk_photo where TID="& ClassID &" and Verific=1 and SaveFile=0" )(0) >0 then lx=3
		if lx=0 then
				Response.Write "没找到所要的数据.检查是否审合 或 你以经保存过了"
		Else
		
			Response.Redirect "savefile.asp?lx="& lx &"&ClassID="& ClassID 
		end if
			call Main()	
		end if
end sub
'--------------------------SQL函数集------------------------
'===================================================
'小金通用采集系统
'作  用：SQL计算记录集总数
'===================================================
'==================================================
'过程名：Admin_ShowChannel_Name
'作  用：显示频道名称
'参  数：ChannelID ------频道ID
'==================================================
Sub Admin_ShowChannel_Name(ChannelID)   
   Dim Sqlc,Rsc,TempStr
   ChannelID=Clng(ChannelID)
   Sqlc ="select top 1 ChannelName from sk_Channel Where ChannelID=" & ChannelID   
   Set Rsc=server.CreateObject("adodb.recordset")   
   Rsc.open Sqlc,ConnItem,1,1   
   If Rsc.Eof and Rsc.Bof then   
      TempStr="无指定频道"   
   Else   
      TempStr=Rsc("ChannelName")
   End if   
   Rsc.Close   
   Set Rsc=Nothing
   response.write TempStr   
End Sub  

'==================================================
'过程名：Admin_ShowChannel_Option
'作  用：显示频道选项
'参  数：ChannelID ------频道ID
'==================================================
'--ipq改
Sub Admin_ShowChannel_Option(ChannelID)   
   Dim Sqlc,Rsc,ChannelName,TempStr
   ChannelID=Clng(ChannelID)
   Sqlc ="select ChannelID,ChannelName from sk_Channel where ModuleType=1 order by ChannelID asc"   
   Set Rsc=server.CreateObject("adodb.recordset")   
   Rsc.Open Sqlc,ConnItem,1,1  
   TempStr="<option value=""0"">请选择频道</option>"    
   If Rsc.Eof and Rsc.Bof Then
      TempStr=TempStr & "<option value=""0"">请添加频道</option>"   
   Else
      Do while not Rsc.Eof   
         TempStr=TempStr & "<option value=" & """" & Rsc("ChannelID") & """" & "" 
         If ChannelID=Rsc("ChannelID") Then
            TempStr=TempStr & "selected"
			End If
		 TempStr=TempStr & ">" & Rsc("ChannelName")
         TempStr=TempStr & "</option>"  
      Rsc.Movenext   
      Loop   
   End if
   Rsc.Close   
   Set Rsc=Nothing   
   Response.Write TempStr   
End sub 
'--ipq改
Sub Admin_ShowChannel_Opin(ChannelID)   
   Dim Sqlc,Rsc,ChannelName,TempStr
   ChannelID=Clng(ChannelID)
   Sqlc ="select ChannelID,ChannelName from sk_Channel where ModuleType=1 order by ChannelID asc"   
   Set Rsc=server.CreateObject("adodb.recordset")   
   Rsc.Open Sqlc,ConnItem,1,1  
   TempStr="<option value=""0"" selected>请选择频道</option>"    
   If Rsc.Eof and Rsc.Bof Then
      TempStr=TempStr & "<option value=""0"">请添加频道</option>"   
   Else
      Do while not Rsc.Eof   
         TempStr=TempStr & "<option value=" & """" & Rsc("ChannelID") & """" & "" 
         If ChannelID=Rsc("ChannelID") Then
            TempStr=TempStr & ""
         End If
         TempStr=TempStr & ">" & Rsc("ChannelName")
         TempStr=TempStr & "</option>"  
      Rsc.Movenext   
      Loop   
   End if
   Rsc.Close   
   Set Rsc=Nothing   
   Response.Write TempStr   
End sub 

'==================================================
'过程名：Admin_ShowClass_Name
'作  用：显示栏目名称
'参  数：ChannelID ------频道ID
'参  数：ClassID ------栏目ID
'==================================================
Sub Admin_ShowClass_Name(ChannelID,ClassID)   
   Dim SqlC,RsC,TempStr
   Sqlc ="Select top 1 colname from tr_column Where  id=" & ClassID 
   'Sqlc ="Select top 1 FolderName from ks_Class Where ID='"& ClassID &"'"   
   Set RsC=server.CreateObject("adodb.recordset")   
   RsC.Open SqlC,conn,1,1   
   If RsC.Eof And RsC.Bof Then   
      TempStr="无指定栏目"   
   Else   
      TempStr=RsC("colname")
   End if   
   RsC.Close   
   Set RsC=Nothing
   Response.Write TempStr   
End Sub  

'==================================================
'过程名：Admin_ShowSpecial_Name
'作  用：显示专题名称
'参  数：ChannelID ------频道ID
'参  数：SpecialID ------专题ID
'==================================================
Sub Admin_ShowSpecial_Name(ChannelID,SpecialID)   
   Dim Sqlc,Rsc,TempStr
   ChannelID=Clng(ChannelID)
   SpecialID=Clng(SpecialID)
   Sqlc ="select top 1 SpecialName from SK_Special Where ChannelID=" & ChannelID & " and SpecialID=" & SpecialID   
   Set Rsc=server.CreateObject("adodb.recordset")   
   Rsc.open Sqlc,ConnItem,1,1   
   If Rsc.Eof and Rsc.Bof then   
      TempStr="无指定专题"   
   Else   
      TempStr=Rsc("SpecialName")
   End if   
   Rsc.Close   
   Set Rsc=Nothing
   Response.Write TempStr   
End Sub  

'==================================================
'过程名：Admin_ShowClass_Option
'作  用：显示栏目选项
'参  数：ChannelID ------频道ID
'参  数：ClassID ------栏目ID
'==================================================
sub Admin_ShowClass_Option(ChannelID,ClassID)
	dim rsClass,sqlClass,strTempC,tmpDepth,i
	dim arrShowLine(20)
	ChannelID=Clng(ChannelID)
	ClassID=Clng(ClassID)
	for i=0 to ubound(arrShowLine)
		arrShowLine(i)=False
	next
        StrTempC=""
	sqlClass="Select * From SK_Class where ChannelID="& ChannelID & " order by RootID,OrderID"
	set rsClass=ConnItem.execute(sqlClass)
	if rsClass.bof and rsClass.eof then 
		strTempC="<option value='0'>请先添加栏目</option>"
	Else
                do while not rsClass.eof
                        tmpDepth=rsClass("Depth")
			if rsClass("NextID")>0 then
				arrShowLine(tmpDepth)=True
			Else
				arrShowLine(tmpDepth)=False
			end if
	        strTempC=StrTempC & "<option value='" & rsClass("ClassID") & "'"
		        if rsClass("ClassID")=ClassID then
			        strTempC=strTempC & "selected"
		        end if
			strTempC=strTempC & ">" &rsClass("ClassName")
			strTempC=strTempC & "</option>"
		rsClass.movenext
		loop
	end if
	rsClass.close
	set rsClass=nothing
	response.write strTempc
end sub

'==================================================
'过程名：Admin_ShowSpecial_Option
'作  用：显示专题选项
'参  数：ChannelID ------频道ID
'参  数：SpecialID ------专题ID
'==================================================
sub Admin_ShowSpecial_Option(ChannelID,SpecialID)
    ChannelID=Clng(ChannelID)
    SpecialID=Clng(SpecialID)
    Dim TempStr
	TempStr="<select name='SpecialID' id='SpecialID'><option value='0'"
	if SpecialID=0 then
		TempStr=TempStr & " selected"
	end if
	TempStr=TempStr & ">不属于任何专题</option>"
	                
	dim sqlSpecial,rsSpecial
        sqlSpecial = "select * from SK_Special where ChannelID=" & ChannelID
	set rsSpecial=server.CreateObject("adodb.recordset")
	rsSpecial.open sqlSpecial,ConnItem,1,1
	do while not rsSpecial.eof
		if rsSpecial("SpecialID")=SpecialID then
			TempStr=TempStr & "<option value='" & rsSpecial("SpecialID") & "' selected>" & rsSpecial("SpecialName") & "</option>"
		Else
			TempStr=TempStr & "<option value='" & rsSpecial("SpecialID") & "'>" & rsSpecial("SpecialName") & "</option>"
		end if
	rsSpecial.movenext
	loop
	rsSpecial.close
        set rsSpecial = nothing
        Response.write TempStr
end sub



'==================================================
'过程名：Admin_ShowTemplate_Option
'作  用：显示设计模板选项
'参  数：TemplateID ------设计模板ID
'参  数：ChannelID-----
'==================================================
sub Admin_ShowTemplate_Option(ChannelID,TemplateType,TemplateID)
	dim sqlTemplate,rsTemplate,TempStr
    ChannelID=Clng(ChannelID)
    TempLateType=Clng(TempLateType)
    TempLateID=Clng(TempLateID)
	TempStr="<select name='TemplateID' id='TemplateID'><option value='0'>系统默认内容页模板</option>"
	sqlTemplate="select * from PE_Template where TemplateType=" & TemplateType & " And ChannelID=" & ChannelID
	set rsTemplate=server.CreateObject("adodb.recordset")
	rsTemplate.open sqlTemplate,ConnItem,1,1
	if rsTemplate.bof and rsTemplate.eof then
	  	TempStr= TempStr & "<option value='0'>请你先添加模板</option>"
	Else
	  	do while not rsTemplate.eof
	  		if rsTemplate("TemplateID")=TemplateID then
				TempStr= TempStr & "<option value='" & rsTemplate("TemplateID") & "' selected>" & rsTemplate("TemplateName") & "</option>"
			Else		
				TempStr= TempStr & "<option value='" & rsTemplate("TemplateID") & "'>" & rsTemplate("TemplateName") & "</option>"
	  		end if		
			rsTemplate.movenext
	  	loop
	end if
	rsTemplate.close
	set rsTemplate=nothing
    TempStr= TempStr & "</select>"
    Response.Write TempStr
end sub

'==================================================
'过程名：Admin_ShowItem_Name
'作  用：显示项目名称
'参  数：ItemID ------项目ID
'==================================================
Sub Admin_ShowItem_Name(ItemID)   
   Dim Sqlc,Rsc,TempStr
   ItemID=Clng(ItemID)
   Sqlc ="select top 1 ItemName from Item Where ItemID=" & ItemID   
   Set Rsc=server.CreateObject("adodb.recordset")   
   Rsc.open Sqlc,ConnItem,1,1   
   If Rsc.Eof and Rsc.Bof then   
      TempStr="无指定项目"   
   Else   
      TempStr=Rsc("ItemName")
   End if   
   Rsc.Close   
   Set Rsc=Nothing
   Response.Write TempStr   
End Sub  


'==================================================
'过程名：Admin_ShowItem_Option
'作  用：显示项目选项
'参  数：ItemID ------项目ID
'==================================================
Sub Admin_ShowItem_Option(ItemID,lx)   
   Dim SqlI,RsI,TempStr
   ItemID=Clng(ItemID)
   if lx="" or lx=0 then
   		SqlI ="select ItemID,ItemName from Item order by ItemID desc"
   Else
   		SqlI ="select ItemID,ItemName from Item where Colleclx="& lx &" order by ItemID desc"   
   end if
   Set RsI=server.CreateObject("adodb.recordset")   
   RsI.Open SqlI,ConnItem,1,1
   TempStr="<select Name=""ItemID"" ID=""ItemID"">"   
   If RsI.Eof and RsI.Bof Then
      TempStr=TempStr & "<option value=""0"">请添加项目</option>"   
   Else   
      TempStr=TempStr & "<option value=""0"">请选择项目</option>"
      Do while not RsI.Eof   
         TempStr=TempStr & "<option value=" & """" & RsI("ItemID") & """" & "" 
         If ItemID=RsI("ItemID") Then
            TempStr=TempStr & " Selected"
         End If
         TempStr=TempStr & ">" & RsI("ItemName")
         TempStr=TempStr & "</option>"  
      RsI.Movenext   
      Loop   
   End if
   RsI.Close   
   Set RsI=Nothing   
   TempStr=TempStr & "</select>"
   Response.Write TempStr   
End sub   
'==================================================
'函数名：GetHttpPage
'作  用：获取网页源码
'参  数：HttpUrl ------网页地址,Cset 编码
'==================================================
Function GetHttpPage(ByVal URL, ByVal Cset)
On Error Resume Next
Dim Http
 If IsNull(URL)=True Or Len(URL)<18 Or URL="$False$" Then
      GetHttpPage="$False$"
      Exit Function
   End If
   Set Http=server.createobject("MSXML2.XMLHTTP")
   Http.open "GET",URL,False
   Http.Send()
   If Http.Readystate<>4 then
      Set Http=Nothing 
      GetHttpPage="$False$"
      Exit function
   End if
   GetHTTPPage=bytesToBSTR(Http.responseBody,Cset)
   Set Http=Nothing
   
   If Err.number<>0 then
   	  If IsNull(URL)=True Or Len(URL)<18 Or URL="$False$" Then
      GetHttpPage="$False$"
      Exit Function
   End If
   Set Http=server.createobject("MSXML2.XMLHTTP")
   'On Error Resume Next
   Http.open "GET",URL,False
   Http.Send()
   If Http.Readystate<>4 or Http.Status > 300 then
      Set Http=Nothing 
      GetHttpPage="$False$"
      Exit function
   End if
   GetHTTPPage=bytesToBSTR(Http.responseBody,Cset)
   Set Http=Nothing
   	  Err.Clear
   End If
   
End Function
'==================================================
'函数名：BytesToBstr
'作  用：将获取的源码转换为中文
'参  数：Body ------要转换的变量
'参  数：Cset ------要转换的类型
'==================================================
Function BytesToBstr(Body,Cset)
   Dim Objstream
   Set Objstream = Server.CreateObject("adodb.stream")
   objstream.Type = 1
   objstream.Mode =3
   objstream.Open
   objstream.Write body
   objstream.Position = 0
   objstream.Type = 2
   objstream.Charset = Cset
   BytesToBstr = objstream.ReadText 
   objstream.Close
   set objstream = nothing
End Function

'==================================================
'函数名：PostHttpPage
'作  用：登录
'==================================================
Function PostHttpPage(RefererUrl,PostUrl,PostData) 
    Dim xmlHttp 
    Dim RetStr      
    Set xmlHttp = CreateObject("Msxml2.XMLHTTP")  
    xmlHttp.Open "POST", PostUrl, False
    XmlHTTP.setRequestHeader "Content-Length",Len(PostData) 
    xmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    xmlHttp.setRequestHeader "Referer", RefererUrl
    xmlHttp.Send PostData 
    If Err.Number <> 0 Then 
        Set xmlHttp=Nothing
        PostHttpPage = "$False$"
        Exit Function
    End If
    PostHttpPage=bytesToBSTR(xmlHttp.responseBody,"utf-8")
    Set xmlHttp = nothing
End Function 

'==================================================
'函数名：UrlEncoding
'作  用：转换编码
'==================================================
Function UrlEncoding(DataStr)
    Dim StrReturn,Si,ThisChr,InnerCode,Hight8,Low8
    StrReturn = ""
    For Si = 1 To Len(DataStr)
        ThisChr = Mid(DataStr,Si,1)
        If Abs(Asc(ThisChr)) < &HFF Then
            StrReturn = StrReturn & ThisChr
        Else
            InnerCode = Asc(ThisChr)
            If InnerCode < 0 Then
               InnerCode = InnerCode + &H10000
            End If
            Hight8 = (InnerCode  And &HFF00)\ &HFF
            Low8 = InnerCode And &HFF
            StrReturn = StrReturn & "%" & Hex(Hight8) &  "%" & Hex(Low8)
        End If
    Next
    UrlEncoding = StrReturn
End Function

'==================================================
'函数名：GetBody
'作  用：截取字符串
'参  数：ConStr ------将要截取的字符串
'参  数：StartStr ------开始字符串
'参  数：OverStr ------结束字符串
'参  数：IncluL ------是否包含StartStr
'参  数：IncluR ------是否包含OverStr
'==================================================
Function GetBody(ConStr,StartStr,OverStr,IncluL,IncluR)
   If ConStr="$False$" or ConStr="" or IsNull(ConStr)=True Or StartStr="" or IsNull(StartStr)=True Or OverStr="" or IsNull(OverStr)=True Then
      GetBody="$False$"
      Exit Function
   End If
   Dim ConStrTemp
   Dim Start,Over
   ConStrTemp=Lcase(ConStr)
   StartStr=Lcase(StartStr)
   OverStr=Lcase(OverStr)
   Start = InStrB(1, ConStrTemp, StartStr, vbBinaryCompare)
   If Start<=0 then
      GetBody="$False$"
      Exit Function
   Else
      If IncluL=False Then
         Start=Start+LenB(StartStr)
      End If
   End If
   Over=InStrB(Start,ConStrTemp,OverStr,vbBinaryCompare)
   If Over<=0 Or Over<=Start then
      GetBody="$False$"
      Exit Function
   Else
      If IncluR=True Then
         Over=Over+LenB(OverStr)
      End If
   End If
   GetBody=MidB(ConStr,Start,Over-Start)
End Function


'==================================================
'函数名：GetArray
'作  用：提取链接地址，以$Array$分隔
'参  数：ConStr ------提取地址的原字符
'参  数：StartStr ------开始字符串
'参  数：OverStr ------结束字符串
'参  数：IncluL ------是否包含StartStr
'参  数：IncluR ------是否包含OverStr
'==================================================
Function GetArray(Byval ConStr,StartStr,OverStr,IncluL,IncluR)
On Error Resume Next
   If ConStr="$False$" or ConStr="" Or IsNull(ConStr)=True or StartStr="" Or OverStr="" or  IsNull(StartStr)=True Or IsNull(OverStr)=True Then
      GetArray="$False$"
      Exit Function
   End If
   Dim TempStr,TempStr2,objRegExp,Matches,Match,Templisturl
   TempStr=""
   Set objRegExp = New Regexp 
   objRegExp.IgnoreCase = True 
   objRegExp.Global = True
   objRegExp.Pattern = "("&StartStr&").+?("&OverStr&")"
   Set Matches =objRegExp.Execute(ConStr) 
   For Each Match in Matches
      if Templisturl =Match.Value then
	  Else
	  TempStr=TempStr & "$Array$" & Match.Value
	  Templisturl=Match.Value
	  end if
   Next 
   Set Matches=nothing

   If TempStr="" Then
      GetArray="$False$"
      Exit Function
   End If
   TempStr=Right(TempStr,Len(TempStr)-7)
   If IncluL=False then
      objRegExp.Pattern =StartStr
      TempStr=objRegExp.Replace(TempStr,"")
   End if
   If IncluR=False then
      objRegExp.Pattern =OverStr
      TempStr=objRegExp.Replace(TempStr,"")
   End if
   Set objRegExp=nothing
   Set Matches=nothing
   
   TempStr=Replace(TempStr,"""","")
   TempStr=Replace(TempStr,"'","")
   TempStr=Replace(TempStr," ","")
   'TempStr=Replace(TempStr,"(","")
   'TempStr=Replace(TempStr,")","")

   If TempStr="" then
      GetArray="$False$"
   Else
      GetArray=TempStr
   End if
End Function


'==================================================
'函数名：DefiniteUrl
'作  用：将相对地址转换为绝对地址
'参  数：PrimitiveUrl ------要转换的相对地址
'参  数：ConsultUrl ------当前网页地址
'==================================================
Function DefiniteUrl(Byval PrimitiveUrl,Byval ConsultUrl)
   Dim ConTemp,PriTemp,Pi,Ci,PriArray,ConArray
   If PrimitiveUrl="" or ConsultUrl="" or PrimitiveUrl="$False$" or ConsultUrl="$False$" Then
      DefiniteUrl="$False$"
      Exit Function
   End If
  
   If Left(Lcase(ConsultUrl),7)<>"http://" and  Left(Lcase(ConsultUrl),8)<>"https://" Then
 
      ConsultUrl= "http://" & ConsultUrl
   End If
   ConsultUrl=Replace(ConsultUrl,"\","/")
   ConsultUrl=Replace(ConsultUrl,"://",":\\")
   PrimitiveUrl=Replace(PrimitiveUrl,"\","/")

   If Right(ConsultUrl,1)<>"/" Then
      If Instr(ConsultUrl,"/")>0 Then
         If Instr(Right(ConsultUrl,Len(ConsultUrl)-InstrRev(ConsultUrl,"/")),".")>0 then   
         Else
            ConsultUrl=ConsultUrl & "/"
         End If
      Else
         ConsultUrl=ConsultUrl & "/"
      End If
   End If
   ConArray=Split(ConsultUrl,"/")
  
   If Left(LCase(PrimitiveUrl),7) = "http://" or  Left(LCase(PrimitiveUrl),8) = "https://" then
 
      DefiniteUrl=Replace(PrimitiveUrl,"://",":\\")
   ElseIf Left(PrimitiveUrl,1) = "/" Then
      DefiniteUrl=ConArray(0) & PrimitiveUrl
   ElseIf Left(PrimitiveUrl,2)="./" Then
      PrimitiveUrl=Right(PrimitiveUrl,Len(PrimitiveUrl)-2)
      If Right(ConsultUrl,1)="/" Then   
         DefiniteUrl=ConsultUrl & PrimitiveUrl
      Else
         DefiniteUrl=Left(ConsultUrl,InstrRev(ConsultUrl,"/")) & PrimitiveUrl
      End If
   ElseIf Left(PrimitiveUrl,3)="../" then
      Do While Left(PrimitiveUrl,3)="../"
		 PrimitiveUrl=Right(PrimitiveUrl,Len(PrimitiveUrl)-3)
         Pi=Pi+1
      Loop            
      For Ci=0 to (Ubound(ConArray)-1-Pi)
         If DefiniteUrl<>"" Then
            DefiniteUrl=DefiniteUrl & "/" & ConArray(Ci)
         Else
            DefiniteUrl=ConArray(Ci)
         End If
      Next
      DefiniteUrl=DefiniteUrl & "/" & PrimitiveUrl
   Else
      If Instr(PrimitiveUrl,"/")>0 Then
         PriArray=Split(PrimitiveUrl,"/")
         If Instr(PriArray(0),".")>0 Then
            If Right(PrimitiveUrl,1)="/" Then
               DefiniteUrl="http:\\" & PrimitiveUrl
            Else
               If Instr(PriArray(Ubound(PriArray)-1),".")>0 Then 
                  DefiniteUrl="http:\\" & PrimitiveUrl
               Else
                  DefiniteUrl="http:\\" & PrimitiveUrl & "/"
               End If
            End If      
         Else
            If Right(ConsultUrl,1)="/" Then   
               DefiniteUrl=ConsultUrl & PrimitiveUrl
            Else
               DefiniteUrl=Left(ConsultUrl,InstrRev(ConsultUrl,"/")) & PrimitiveUrl
            End If
         End If
      Else
         If Instr(PrimitiveUrl,".")>0 Then
            If Right(ConsultUrl,1)="/" Then
               If right(LCase(PrimitiveUrl),3)=".cn" or right(LCase(PrimitiveUrl),3)="com" or right(LCase(PrimitiveUrl),3)="net" or right(LCase(PrimitiveUrl),3)="org" Then
                  DefiniteUrl="http:\\" & PrimitiveUrl & "/"
               Else
                  DefiniteUrl=ConsultUrl & PrimitiveUrl
               End If
            Else
               If right(LCase(PrimitiveUrl),3)=".cn" or right(LCase(PrimitiveUrl),3)="com" or right(LCase(PrimitiveUrl),3)="net" or right(LCase(PrimitiveUrl),3)="org" Then
                  DefiniteUrl="http:\\" & PrimitiveUrl & "/"
               Else
                  DefiniteUrl=Left(ConsultUrl,InstrRev(ConsultUrl,"/")) & "/" & PrimitiveUrl
               End If
            End If
         Else
            If Right(ConsultUrl,1)="/" Then
               DefiniteUrl=ConsultUrl & PrimitiveUrl & "/"
            Else
               DefiniteUrl=Left(ConsultUrl,InstrRev(ConsultUrl,"/")) & "/" & PrimitiveUrl & "/"
            End If         
         End If
      End If
   End If
   If Left(DefiniteUrl,1)="/" then
     DefiniteUrl=Right(DefiniteUrl,Len(DefiniteUrl)-1)
   End if
   If DefiniteUrl<>"" Then
      DefiniteUrl=Replace(DefiniteUrl,"//","/")
      DefiniteUrl=Replace(DefiniteUrl,":\\","://")
   Else
      DefiniteUrl="$False$"
   End If
End Function
'==================================================
'函数名：ReplaceSaveRemoteFile
'作  用：替换、保存远程图片
'参  数：ConStr ------ 要替换的字符串
'参  数：SaveTf ------ 是否保存文件，False不保存，True保存
'参  数: TistUrl------ 当前网页地址
'==================================================
Function ReplaceSaveRemoteFile(ConStr,strInstallDir,strChannelDir,SaveTf,TistUrl)
   If ConStr="$False$" or ConStr="" or strChannelDir="" Then
      ReplaceSaveRemoteFile=ConStr
      Exit Function
   End If
   Dim TempStr,TempStr2,TempStr3,Re,Matches,Match,Tempi,TempArray,TempArray2

   Set Re = New Regexp 
   Re.IgnoreCase = True 
   Re.Global = True
   Re.Pattern ="<img.+?[^\>]>"
   Set Matches =Re.Execute(ConStr) 
   For Each Match in Matches
      If TempStr<>"" then 
         TempStr=TempStr & "$Array$" & Match.Value
      Else
         TempStr=Match.Value
      End if
   Next
   If TempStr<>"" Then
      TempArray=Split(TempStr,"$Array$")
      TempStr=""
      For Tempi=0 To Ubound(TempArray)
         Re.Pattern ="src\s*=\s*.+?\.(gif|jpg|bmp|jpeg|psd|png|svg|dxf|wmf|tiff)"
         Set Matches =Re.Execute(TempArray(Tempi)) 
         For Each Match in Matches
            If TempStr<>"" then 
               TempStr=TempStr & "$Array$" & Match.Value
            Else
               TempStr=Match.Value
            End if
         Next
      Next
   End if
   If TempStr<>"" Then
   	  IncludePic=1'图片新闻
      Re.Pattern ="src\s*=\s*"
      TempStr=Re.Replace(TempStr,"")
   End If
   Set Matches=nothing
   Set Re=nothing
   If TempStr="" or IsNull(TempStr)=True Then
      ReplaceSaveRemoteFile=ConStr
      Exit function
   End if
   TempStr=Replace(TempStr,"""","")
   TempStr=Replace(TempStr,"'","")
   TempStr=Replace(TempStr," ","")
   
   Dim RemoteFileurl,SavePath,PathTemp,DtNow,strFileName,strFileType,ArrSaveFileName,RanNum,Arr_Path
   DtNow=Now()
   If SaveTf=True then
 '***********************************
      SavePath= strChannelDir & year(DtNow) & right("0" & month(DtNow),2) & "/"
	  response.write "链接路径：" & savepath & "<br>"
      Arr_Path=Split(SavePath,"/")
      PathTemp=""
      For Tempi=0 To Ubound(Arr_Path)
         If Tempi=0 Then
            PathTemp=Arr_Path(0) & "/"
         ElseIf Tempi=Ubound(Arr_Path) Then
            Exit For
         Else
            PathTemp=PathTemp & Arr_Path(Tempi) & "/"
         End If
         If CheckDir(PathTemp)=False Then
            If MakeNewsDir(PathTemp)=False Then
               SaveTf=False
               Exit For
            End If
         End If
      Next
   End If

   '去掉重复图片开始
   TempArray=Split(TempStr,"$Array$")
   TempStr=""
   For Tempi=0 To Ubound(TempArray)
      If Instr(Lcase(TempStr),Lcase(TempArray(Tempi)))<1 Then
         TempStr=TempStr & "$Array$" & TempArray(Tempi)
      End If
   Next
   TempStr=Right(TempStr,Len(TempStr)-7)
   TempArray=Split(TempStr,"$Array$")
   '去掉重复图片结束

   '转换相对图片地址开始
   TempStr=""
   For Tempi=0 To Ubound(TempArray)
      TempStr=TempStr & "$Array$" & DefiniteUrl(TempArray(Tempi),TistUrl)
   Next
   TempStr=Right(TempStr,Len(TempStr)-7)
   TempStr=Replace(TempStr,Chr(0),"")
   TempArray2=Split(TempStr,"$Array$")
   TempStr=""
   '转换相对图片地址结束
	'图片替换/保存
   Set Re = New Regexp
   Re.IgnoreCase = True 
   Re.Global = True
   For Tempi=0 To Ubound(TempArray2)
      RemoteFileUrl=TempArray2(Tempi)
      If RemoteFileUrl<>"$False$" And SaveTf=True Then'保存图片
         ArrSaveFileName = Split(RemoteFileurl,".")
	 strFileType=Lcase(ArrSaveFileName(Ubound(ArrSaveFileName)))'文件类型
         If strFileType="asp" or strFileType="asa" or strFileType="aspx" or strFileType="cer" or strFileType="cdx" or strFileType="exe" or strFileType="rar" or strFileType="zip" then
            UploadFiles=""
            ReplaceSaveRemoteFile=ConStr
            Exit Function
         End If

         Randomize
         RanNum=Int(900*Rnd)+100
	 strFileName = year(DtNow) & right("0" & month(DtNow),2) & right("0" & day(DtNow),2) & right("0" & hour(DtNow),2) & right("0" & minute(DtNow),2) & right("0" & second(DtNow),2) & ranNum & "." & strFileType
         Re.Pattern =TempArray(Tempi)
		 
	 If SaveRemoteFile(SavePath & strFileName,RemoteFileUrl)=True Then
'********************************
            PathTemp=SavePath & strFileName 
            ConStr=Re.Replace(ConStr,PathTemp)
            Re.Pattern=strInstallDir & strChannelDir 
            UploadFiles=UploadFiles & "|" & Re.Replace(SavePath &strFileName,"")
			Response.Flush()
			response.write " &nbsp;&nbsp;&nbsp;图片保存地址：" & PathTemp & "<br>"
			if Thumb_WaterMark=1 then call SKThumb.AddWaterMark(PathTemp)'水印
         Else
            PathTemp=RemoteFileUrl
            ConStr=Re.Replace(ConStr,PathTemp)
            'UploadFiles=UploadFiles & "|" & RemoteFileUrl
         End If
      ElseIf RemoteFileurl<>"$False$" and SaveTf=False Then'不保存图片
         Re.Pattern =TempArray(Tempi)
         ConStr=Re.Replace(ConStr,RemoteFileUrl)
         UploadFiles=UploadFiles & "|" & RemoteFileUrl
      End If
   Next   
   Set Re=nothing
   If UploadFiles<>"" Then
      UploadFiles=Right(UploadFiles,Len(UploadFiles)-1)
   End If
   ReplaceSaveRemoteFile=ConStr
End function
	'================================================
	'函数名：ReSaveRemoteFile
	'作  用：查找文件保存替换
	'参  数：Str   ----原字符串
	'参  数：url   ----当然网站URL
	'参  数：Dir -----保存目录
	'参  数：InSave ------是否保存,True,False
	'返回值：格式化取后的字符串
	'================================================
	Public Function ReSaveRemoteFile(ByVal str, ByVal URL, ByVal Dir,InSave)
		Dim s_Content
		Dim re
		Dim ContentFile, ContentFileUrl
		Dim strTempUrl,strFileUrl,DirTemp,PathTemp,FileTemp,Tempi,TempUrlArray,Arr_Path
		Dim sAllowExtName
		sAllowExtName="rm|swf"
		
		s_Content = str
		On Error Resume Next
		Set re = New RegExp
		re.IgnoreCase = True
		re.Global = True
		re.Pattern = "((src=|href=)((\S)+[.]{1}(" & sAllowExtName & ")))"
		Set ContentFile = re.Execute(s_Content)
		Dim sContentUrl(), n, i, bRepeat
		n = 0
		For Each ContentFileUrl In ContentFile
			strFileUrl = Replace(Replace(Replace(Replace(ContentFileUrl.Value, "src=", "", 1, -1, 1), "href=", "", 1, -1, 1), "'", ""), Chr(34), "")
			If n = 0 Then
				n = n + 1
				ReDim sContentUrl(n)
				sContentUrl(n) = strFileUrl
			Else
				bRepeat = False
				For i = 1 To UBound(sContentUrl)
					If UCase(strFileUrl) = UCase(sContentUrl(i)) Then
						bRepeat = True
						Exit For
					End If
				Next
				If bRepeat = False Then
					n = n + 1
					ReDim Preserve sContentUrl(n)
					sContentUrl(n) = strFileUrl
				End If
			End If
		Next
		If n = 0 Then
			ReSaveRemoteFile = s_Content
			Exit Function
		End If
		For i = 1 To n 
			strTempUrl = sContentUrl(i) : strTempUrl = FormatRemoteUrl(strTempUrl,URL)'得到文件地址
			Response.Write(strTempUrl)
			IF InSave=True then
				Arr_Path=Split(Dir,"/")
				'----------建目录-----------------------
				  For Tempi=0 To Ubound(Arr_Path)
					 If Tempi=0 Then
						PathTemp=Arr_Path(0) & "/"
					 ElseIf Tempi=Ubound(Arr_Path) Then
						Exit For
					 Else
						PathTemp=PathTemp & Arr_Path(Tempi) & "/"
					 End If
					 If CheckDir(PathTemp)=False Then
						If MakeNewsDir(PathTemp)=False Then
						   SaveTf=False
						   Exit For
						End If
					 End If
				  Next
				 '------------------------------------------------------
				TempUrlArray=Split(strTempUrl,"/")
				'----------检查文件是否存在.如果存在换文件名------------------
				Do while True 
					FileTemp=Dir &  MakeRandom(5) & TempUrlArray(Ubound(TempUrlArray))'生成随机文件名
					If CheckFile(FileTemp)=False then
						Exit Do
					end if
				loop 
				'-------------------------------------------------------------------
				Response.Write(FileTemp)
				If SaveRemoteFile(FileTemp,strTempUrl)=True then
					Response.Write("保存成功")&"<Br>"
					s_Content = Replace(s_Content,sContentUrl(i),FileTemp, 1, -1, 1)'替换地址	
				Else
					Response.Write("保存失败")&"<Br>"
				End if
			Else
				s_Content = Replace(s_Content,sContentUrl(i),strTempUrl, 1, -1, 1)'替换地址		
			End If	
		Next
		Set re = Nothing
		PictureExist = True
		ReSaveRemoteFile = s_Content
		Exit Function
	End Function
'================================================
'函数名：FormatRemoteUrl
'作  用：格式化成当前网站完整的URL-将相对地址转换为绝对地址
'参  数： url ----Url字符串
'参  数： CurrentUrl ----当然网站URL
'返回值：格式化取后的Url
'================================================
	Public Function FormatRemoteUrl(ByVal URL,ByVal CurrentUrl)
		Dim strUrl
		If Len(URL) < 2 Or Len(URL) > 255 Or Len(CurrentUrl) < 2 Then
			FormatRemoteUrl = vbNullString
			Exit Function
		End If
		CurrentUrl = Trim(Replace(Replace(Replace(Replace(Replace(CurrentUrl, "'", vbNullString), """", vbNullString), vbNewLine, vbNullString), "\", "/"), "|", vbNullString))
		URL = Trim(Replace(Replace(Replace(Replace(Replace(URL, "'", vbNullString), """", vbNullString), vbNewLine, vbNullString), "\", "/"), "|", vbNullString))	
		If InStr(9, CurrentUrl, "/") = 0 Then
			strUrl = CurrentUrl
		Else
			strUrl = Left(CurrentUrl, InStr(9, CurrentUrl, "/") - 1)
		End If

		If strUrl = vbNullString Then strUrl = CurrentUrl
		Select Case Left(LCase(URL), 6)
			Case "http:/", "https:", "ftp://", "rtsp:/", "mms://"
				FormatRemoteUrl = URL
				Exit Function
		End Select

		If Left(URL, 1) = "/" Then
			FormatRemoteUrl = strUrl & URL
			Exit Function
		End If
		
		If Left(URL, 3) = "../" Then
			Dim ArrayUrl
			Dim ArrayCurrentUrl
			Dim ArrayTemp()
			Dim strTemp
			Dim i, n
			Dim c, l
			n = 0
			ArrayCurrentUrl = Split(CurrentUrl, "/")
			ArrayUrl = Split(URL, "../")
			c = UBound(ArrayCurrentUrl)
			l = UBound(ArrayUrl) + 1
			
			If c > l + 2 Then
				For i = 0 To c - l
					ReDim Preserve ArrayTemp(n)
					ArrayTemp(n) = ArrayCurrentUrl(i)
					n = n + 1
				Next
				strTemp = Join(ArrayTemp, "/")
			Else
				strTemp = strUrl
			End If
			URL = Replace(URL, "../", vbNullString)
			FormatRemoteUrl = strTemp & "/" & URL
			Exit Function
		End If
		strUrl = Left(CurrentUrl, InStrRev(CurrentUrl, "/"))
		FormatRemoteUrl = strUrl & Replace(URL, "./", vbNullString)
		Exit Function
	End Function	
'================================================
'函数名：ReplaceTrim
'作  用：过滤掉字符中所有的tab和回车和换行
'================================================
	Public Function ReplaceTrim(ByVal strContent)
		On Error Resume Next
		Dim re
		Set re = New RegExp
		re.IgnoreCase = True
		re.Global = True
		re.Pattern = "(" & Chr(8) & "|" & Chr(9) & "|" & Chr(10) & "|" & Chr(13) & ")"
		strContent = re.Replace(strContent, vbNullString)
		Set re = Nothing
		ReplaceTrim = strContent
		Exit Function
	End Function
'==================================================
'函数名: CheckFile
'作  用：检查某一文件是否存在
'参  数：FileName ------ 文件地址 如:/swf/1.swf
'返回值：False  ----  True
'==================================================
	Public Function CheckFile(FileName)
		 On Error Resume Next
		 Dim FsoObj
		 Set FsoObj = Server.CreateObject("Scripting.FileSystemObject")
		  If Not FsoObj.FileExists(Server.MapPath(FileName)) Then
			  CheckFile = False
			  Exit Function
		  End If
		 CheckFile = True:Set FsoObj = Nothing
	End Function
'==================================================
'过程名：SaveRemoteFile
'作  用：保存远程的文件到本地
'参  数：LocalFileName ------ 本地文件名
'参  数：RemoteFileUrl ------ 远程文件URL
'==================================================
Function SaveRemoteFile(LocalFileName,RemoteFileUrl)
    SaveRemoteFile=True
	dim Ads,Retrieval,GetRemoteData	
	On Error Resume Next
	Set Retrieval = Server.CreateObject("Microsoft.XMLHTTP")
	With Retrieval
		.Open "Get", RemoteFileUrl, False, "", ""
		.Send
        If .Readystate<>4 or .Status > 300 then
            SaveRemoteFile=False
            Exit Function
        End If
		GetRemoteData = .ResponseBody
	End With
	Set Retrieval = Nothing
	'If LenB(GetRemoteData) < 100 Then Exit Function
	'If MaxFileSize > 0 Then
			'If LenB(GetRemoteData) > 5000 Then Exit Function
			Response.Write(Round(LenB(GetRemoteData)/1024)) & "KB"
	'End If
	Set Ads = Server.CreateObject("Adodb.Stream")
	With Ads
		.Type = 1
		.Open
		.Write GetRemoteData
		.SaveToFile server.MapPath(LocalFileName),2
		.Cancel()
		.Close()
	End With
	If Err.number<>0 then
	  SaveRemoteFile=False
      Exit Function
   	  Err.Clear
   	End If
	Set Ads=nothing
	
end Function
'==================================================
'函数名：FpHtmlEnCode
'作  用：标题过滤
'参  数：fString ------字符串
'==================================================
Function FpHtmlEnCode(fString)
   If IsNull(fString)=False or fString<>"" or fString<>"$False$" Then
       fString=nohtml(fString)
       fString=FilterJS(fString)
       fString = Replace(fString,"&nbsp;"," ")
       fString = Replace(fString,"&quot;","")
       fString = Replace(fString,"&#39;","")
       fString = replace(fString, ">", "")
       fString = replace(fString, "<", "")
       fString = Replace(fString, CHR(9), " ")'&nbsp;
       fString = Replace(fString, CHR(10), "")
       fString = Replace(fString, CHR(13), "")
       fString = Replace(fString, CHR(34), "")
       fString = Replace(fString, CHR(32), " ")'space
       fString = Replace(fString, CHR(39), "")
       fString = Replace(fString, CHR(10) & CHR(10),"")
       fString = Replace(fString, CHR(10)&CHR(13), "")
       fString=Trim(fString)
       FpHtmlEnCode=fString
   Else
       FpHtmlEnCode="$False$"
   End If
End Function

'==================================================
'函数名：GetPaing
'作  用：获取分页
'==================================================
Function GetPaing(Byval ConStr,StartStr,OverStr,IncluL,IncluR)
If ConStr="$False$" or ConStr="" Or StartStr="" Or OverStr="" or IsNull(ConStr)=True or IsNull(StartStr)=True Or IsNull(OverStr)=True Then
   GetPaing="$False$"
   Exit Function
End If

Dim Start,Over,ConTemp,TempStr
TempStr=LCase(ConStr)
StartStr=LCase(StartStr)
OverStr=LCase(OverStr)
Over=Instr(1,TempStr,OverStr)
If Over<=0 Then
   GetPaing="$False$"
   Exit Function
Else
   If IncluR=True Then
      Over=Over+Len(OverStr)
   End If
End If
TempStr=Mid(TempStr,1,Over)
Start=InstrRev(TempStr,StartStr)
If IncluL=False Then
   Start=Start+Len(StartStr)
End If

If Start<=0 Or Start>=Over Then
   GetPaing="$False$"
   Exit Function
End If
ConTemp=Mid(ConStr,Start,Over-Start)

ConTemp=Trim(ConTemp)
ConTemp=Replace(ConTemp," ","")
ConTemp=Replace(ConTemp,",","")
ConTemp=Replace(ConTemp,"'","")
ConTemp=Replace(ConTemp,"""","")
ConTemp=Replace(ConTemp,">","")
ConTemp=Replace(ConTemp,"<","")
ConTemp=Replace(ConTemp,"&nbsp;","")
GetPaing=ConTemp
End Function

'==================================================
'函数名：ScriptHtml
'作  用：过滤html标记
'参  数：ConStr ------ 要过滤的字符串
'==================================================
Function ScriptHtml(Byval ConStr,TagName,FType)
    Dim Re
    Set Re=new RegExp
    Re.IgnoreCase =true
    Re.Global=True
    Select Case FType
    Case 1
       Re.Pattern="<" & TagName & "([^>])*>"
       ConStr=Re.Replace(ConStr,"")
    Case 2
       Re.Pattern="<" & TagName & "([^>])*>.*?</" & TagName & "([^>])*>"
       ConStr=Re.Replace(ConStr,"") 
	Case 3
       Re.Pattern="<" & TagName & "([^>])*>"
       ConStr=Re.Replace(ConStr,"")
       Re.Pattern="</" & TagName & "([^>])*>"
       ConStr=Re.Replace(ConStr,"")
    End Select
    ScriptHtml=ConStr
    Set Re=Nothing
End Function

'-------------------------
'--检查目录是否存在
'----------------------
Function CheckDir(byval FolderPath)
	dim fso
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	If fso.FolderExists(Server.MapPath(folderpath)) then
	'存在
		CheckDir = True
	Else
	'不存在
		CheckDir = False
	End if
	Set fso = nothing
End Function

'-------------------------
'--建立目录
'----------------------
Function MakeNewsDir(byval foldername)
	dim fso
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
        fso.CreateFolder(Server.MapPath(foldername))
        If fso.FolderExists(Server.MapPath(foldername)) Then
           MakeNewsDir = True
        Else
           MakeNewsDir = False
        End If
	Set fso = nothing
End Function

'**************************************************
'函数名：IsObjInstalled
'作  用：检查组件是否已经安装
'参  数：strClassString ----组件名
'返回值：True  ----已经安装
'       False ----没有安装
'**************************************************
Function IsObjInstalled(strClassString)
	IsObjInstalled = False
	Err = 0
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	If 0 = Err Then IsObjInstalled = True
	Set xTestObj = Nothing
	Err = 0
End Function

'**************************************************
'过程名：WriteErrMsg
'作  用：显示错误提示信息
'参  数：无
'**************************************************
sub WriteErrMsg(ErrMsg)
	dim strErr
	strErr=strErr & "<html><head><title>错误信息</title><meta http-equiv='Content-Type' content='text/html; charset=utf-8'>" & vbcrlf
	strErr=strErr & "<link href='../skin/default/style.css' rel='stylesheet' type='text/css'></head><body><br><br>" & vbcrlf
	strErr=strErr & "<table cellpadding=2 cellspacing=1 border=0 width=400 class='border' align=center>" & vbcrlf
	strErr=strErr & "  <tr align='center' class='title'><td height='22'><strong>错误信息</strong></td></tr>" & vbcrlf
	strErr=strErr & "  <tr class='tdbg'><td height='100' valign='top'><b>产生错误的可能原因：</b>" & ErrMsg &"</td></tr>" & vbcrlf
	strErr=strErr & "  <tr align='center' class='tdbg'><td><a href='javascript:history.go(-1)'>&lt;&lt; 返回上一页</a></td></tr>" & vbcrlf
	strErr=strErr & "</table>" & vbcrlf
	strErr=strErr & "</body></html>" & vbcrlf
	response.write strErr
	response.end
end sub

'**************************************************
'过程名：WriteSucced
'作  用：显示成功提示信息
'参  数：无
'**************************************************
sub WriteSucced(ErrMsg)
	dim strErr
	strErr=strErr & "<html><head><title>成功信息</title><meta http-equiv='Content-Type' content='text/html; charset=utf-8'>" & vbcrlf
	strErr=strErr & "<link href='../admin/Admin_STYLE.CSS' rel='stylesheet' type='text/css'></head><body><br><br>" & vbcrlf
	strErr=strErr & "<table cellpadding=2 cellspacing=1 border=0 width=400 class='border' align=center>" & vbcrlf
	strErr=strErr & "  <tr align='center' class='title'><td height='22'><strong>恭喜你！</strong></td></tr>" & vbcrlf
	strErr=strErr & "  <tr class='tdbg'><td height='100' valign='top' align='center'>" & ErrMsg &"</td></tr>" & vbcrlf
	'strErr=strErr & "  <tr align='center' class='tdbg'><td><a href='javascript:history.go(-1)'>&lt;&lt; 返回上一页</a></td></tr>" & vbcrlf
	strErr=strErr & "</table>" & vbcrlf
	strErr=strErr & "</body></html>" & vbcrlf
	response.write strErr
end sub

'**************************************************
'函数名：ShowPage
'作  用：显示“上一页 下一页”等信息
'参  数：sFileName  ----链接地址
'       TotalNumber ----总数量
'       MaxPerPage  ----每页数量
'       ShowTotal   ----是否显示总数量
'       ShowAllPages ---是否用下拉列表显示所有页面以供跳转。有某些页面不能使用，否则会出现JS错误。
'       strUnit     ----计数单位
'返回值：“上一页 下一页”等信息的HTML代码
'**************************************************
function ShowPage(sFileName,TotalNumber,MaxPerPage,ShowTotal,ShowAllPages,strUnit)
	dim TotalPage,strTemp,strUrl,i

	if TotalNumber=0 or MaxPerPage=0 or isNull(MaxPerPage) then
		ShowPage=""
		exit function
	end if
	if totalnumber mod maxperpage=0 then
    	TotalPage= totalnumber \ maxperpage
  	Else
    	TotalPage= totalnumber \ maxperpage+1
  	end if
	if CurrentPage>TotalPage then CurrentPage=TotalPage
		
  	strTemp= "<table align='center'><tr><td>"
	if ShowTotal=true then 
		strTemp=strTemp & "共 <b>" & totalnumber & "</b> " & strUnit & ""
	end if
	strUrl=JoinChar(sfilename)
  	if CurrentPage<2 then
    	strTemp=strTemp & "<span class=""textbox"">首页</span> <span class=""textbox"">上一页</span>"
  	Else
    	strTemp=strTemp & "<span class=""textbox""><a href='" & strUrl & "page=1'>首页 </a></span>"
    	strTemp=strTemp & "<span class=""textbox""><a href='" & strUrl & "page=" & (CurrentPage-1) & "'>上一页</a></span>"
  	end if

  	if CurrentPage>=TotalPage then
    	strTemp=strTemp & "<span class=""textbox"">下一页</span> <span class=""textbox"">尾页</span>"
  	Else
    	strTemp=strTemp & "<span class=""textbox""><a href='" & strUrl & "page=" & (CurrentPage+1) & "'>下一页</a></span>"
    	strTemp=strTemp & "<span class=""textbox""><a href='" & strUrl & "page=" & TotalPage & "'>尾页</a></span>"
  	end if
   	strTemp=strTemp & "&nbsp;页次：<strong><font color=red>" & CurrentPage & "</font>/" & TotalPage & "</strong>页 "
        strTemp=strTemp & "&nbsp;<b>" & maxperpage & "</b>" & strUnit & "/页"
	if ShowAllPages=True then
		strTemp=strTemp & "&nbsp;&nbsp;转到第<input type='text' name='page' size='3' maxlength='5' value='" & CurrentPage & "' onKeyPress=""if (event.keyCode==13) window.location='" & strUrl & "page=" & "'+this.value;""'>页"
		 'strTemp  = strTemp &"&nbsp;<Input type=""button""  onClick=""window.location.href='" & strUrl & "page='+document.all.page.value;""  name=button1  value=GO >"
	end if
	strTemp=strTemp & "</td></tr></table>"
	ShowPage=strTemp
end function

'**************************************************
'函数名：JoinChar
'作  用：向地址中加入 ? 或 &
'参  数：strUrl  ----网址
'返回值：加了 ? 或 & 的网址
'**************************************************
function JoinChar(strUrl)
	if strUrl="" then
		JoinChar=""
		exit function
	end if
	if InStr(strUrl,"?")<len(strUrl) then 
		if InStr(strUrl,"?")>1 then
			if InStr(strUrl,"&")<len(strUrl) then 
				JoinChar=strUrl & "&"
			Else
				JoinChar=strUrl
			end if
		Else
			JoinChar=strUrl & "?"
		end if
	Else
		JoinChar=strUrl
	end if
end function

'**************************************************
'函数名：CreateKeyWord
'作  用：由给定的字符串生成关键字
'参  数：Constr---要生成关键字的原字符串
'返回值：生成的关键字
'**************************************************
Function CreateKeyWord(byval Constr,Num)
   If Constr="" or IsNull(Constr)=True or Constr="$False$" Then
      CreateKeyWord="$False$"
      Exit Function
   End If
   If Num="" or IsNumeric(Num)=False Then
      Num=2
   End If
   Constr=Replace(Constr,CHR(32),"")
   Constr=Replace(Constr,CHR(9),"")
   Constr=Replace(Constr,"&nbsp;","")
   Constr=Replace(Constr," ","")
   Constr=Replace(Constr,"(","")
   Constr=Replace(Constr,")","")
   Constr=Replace(Constr,"<","")
   Constr=Replace(Constr,">","")
   Constr=Replace(Constr,"""","")
   Constr=Replace(Constr,"?","")
   Constr=Replace(Constr,"*","")
   Constr=Replace(Constr,"|","")
   Constr=Replace(Constr,",","")
   Constr=Replace(Constr,".","")
   Constr=Replace(Constr,"/","")
   Constr=Replace(Constr,"\","")
   Constr=Replace(Constr,"-","")
   Constr=Replace(Constr,"@","")
   Constr=Replace(Constr,"#","")
   Constr=Replace(Constr,"$","")
   Constr=Replace(Constr,"%","")
   Constr=Replace(Constr,"&","")
   Constr=Replace(Constr,"+","")
   Constr=Replace(Constr,":","")
   Constr=Replace(Constr,"：","")   
   Constr=Replace(Constr,"‘","")
   Constr=Replace(Constr,"“","")
   Constr=Replace(Constr,"”","")
   Constr=Replace(Constr,"&","")  
   Constr=Replace(Constr,"gt;","")      
   Dim i,ConstrTemp
   For i=1 To Len(Constr)
      ConstrTemp=ConstrTemp & "|" & Mid(Constr,i,Num)
   Next
   If Len(ConstrTemp)<254 Then
      ConstrTemp=ConstrTemp & "|"
   Else
      ConstrTemp=Left(ConstrTemp,254) & "|"
   End If
   ConstrTemp=left(ConstrTemp,len(ConstrTemp)-1)
   ConstrTemp= Right(ConstrTemp,len(ConstrTemp)-1)
   CreateKeyWord=ConstrTemp
End Function

Function CheckUrl(strUrl)
   Dim Re
   Set Re=new RegExp
   Re.IgnoreCase =true
   Re.Global=True

  Re.Pattern="([http://|https://])([\w-]+\.)+[\w-]+(/[\w-./?%&=]*)?"
    
   If Re.test(strUrl)=True Then
      CheckUrl=strUrl
   Else
      CheckUrl="$False$"
   End If
   Set Rs=Nothing
End Function
Function GetItemId(ItemID,Itemlx) '选择采集
	Dim ItemIdstr,ItemID_id : ItemID_id=Split(ItemID,",")  
	for i=1 to Ubound(ItemID_id)
			If i=1 then 
				ItemIdstr=ItemID_id(i)
			Else
				ItemIdstr=ItemIdstr & "," & ItemID_id(i)
			End if
	NExt
	if Itemlx=0 then GetItemId=ItemID_id(0)	Else GetItemId=ItemIdstr end if
End Function 
%>
