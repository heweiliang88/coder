<%
Dim SKThumb
Set SKThumb=New Thumb
Dim sk
Set sk = New SK_cj
'-----------整合版调用----------------
'Private SKThumb : Set SKThumb =New Thumb
'Private KSCMS : Set KSCMS=New CommonCls
'--------------------------------------
Class SK_cj
Public sub CollPhoto_Show()
If ItemEnd<>True Then
   If (ItemNum-1)>Ubound(Arr_Item,2) then
      ItemEnd=True
   Else
      Call SetItems()
   End If
End If
If ItemEnd<>True Then
   If ListPaingType=0 Then
      If ListNum=1 Then
        ListUrl=ListStr
      Else
         ListEnd=True
      End If
   ElseIf ListPaingType=1 Then
    if listnum=1 and ListPaingID1<>1 and  ListPaingID2<>1 then 
		ListUrl=ListStr
	   else
		  If ListPaingID1>ListPaingID2 then
			 If (ListPaingID1-ListNum+1)<ListPaingID2 or (ListPaingID1-ListNum+1)<0 Then
				Listend=True
			 Else
				ListUrl=Replace(ListPaingStr2,"{$ID}",Cstr(ListpaingID1-ListNum+1))
			 End if
		  Else
			 If (ListPaingID1+ListNum-1)>ListPaingID2 Then
				ListEnd=True
			 Else
				ListUrl=Replace(ListPaingStr2,"{$ID}",CStr(ListPaingID1+ListNum-1))
			 End If
		  End If   
	 end if   
   ElseIf ListPaingType=2 Then
      ListArray=Split(ListPaingStr3,"|")
      If (ListNum-1)>Ubound(ListArray) Then
         ListEnd=True
      Else
         ListUrl=ListArray(ListNum-1)
      End If
   ElseIf ListPaingType=3 Then
	  If ListNum=1 Then
	  	ListUrl=ListStr
		ListCode=ListStr
		application.Lock()
		application("IPQ_SS")=ListUrl
		application.UnLock()
	  Else
	    ListUrl=application("IPQ_SS")
	  	ListCode=SKcj.GetHttpPage(ListUrl,"utf-8")
	    ListCode=SKcj.GetBody(ListCode,LPsString,LPoString,False,False)
	  End if
      If ListCode<>"$False$" or ListCode<>"" Then
	     ListCode=Trim(Skcj.FormatRemoteUrl(ListCode,ListStr))
         ListUrl=ListCode
		 application.Lock()
		 application("IPQ_SS")=ListUrl
		 application.UnLock()
	  Else
	  	 ListEnd=True
	  End if    
   End If
   
   If ListNum>CollecListNum And CollecListNum<>0 Then
      ListEnd=True
   End if
End If

If ItemEnd=True Then
   ErrMsg= "<br>采集任务全部完成"
   ErrMsg=ErrMsg & "<br>成功采集： "  &  NewsSuccesNum  &  "  条,失败： "    &  NewsFalseNum  &  "  条,图片：" & ImagesNumAll & "  张"
   Call DelCache()
   ErrMsg=ErrMsg & "<meta http-equiv=""refresh"" content=""1;url="& Skcj.GetItemConfig("FileName",Colleclx) &""">"
Else
   If ListEnd=True Then
      ItemNum=ItemNum+1
      ListNum=1
      ErrMsg="<br>" & ItemName & "  项目所有列表采集完成，正在整理数据请稍后..."
	  ErrMsg=ErrMsg & "<meta http-equiv=""refresh"" content=""1;url="& Skcj.GetItemConfig("FileName",Colleclx) &""">"
   End If
End If

Call TopItem()
  If ItemEnd=True Or ListEnd=True Then
    If ItemEnd<>True Then
	  Call SetCache_His()'读取历史记录
    End If
    Call WriteSucced(ErrMsg)
Else
   FoundErr=False
   ErrMsg=""
   Call StartCollection()
   Call FootItem2()
End  If
End sub

Public Sub SaveEdit
TsString=Request.Form("TsString")
ToString=Request.Form("ToString")
CsString=Request.Form("CsString")
CoString=Request.Form("CoString")
photourls=Request.Form("photourls")
photourlo=Request.Form("photourlo")

NewsPaingType=Trim(Request.Form("NewsPaingType"))
NPsString=Trim(Request.Form("NPsString"))
NPoString=Trim(Request.Form("NPoString"))
NewsUrlPaing_s=Trim(Request.Form("NewsUrlPaing_s"))
NewsUrlPaing_o=Trim(Request.Form("NewsUrlPaing_o"))

PhotoPaingType=Trim(Request.Form("PhotoPaingType"))
PhotoType_s=Request.Form("PhotoType_s")
PhotoType_o=Request.Form("PhotoType_o")
PhotoLurl_s=Request.Form("PhotoLurl_s")
PhotoLurl_o=Request.Form("PhotoLurl_o")
Phototypefy_s=Request.Form("Phototypefy_s")
Phototypefy_o=Request.Form("Phototypefy_o")
Phototypefyurl_s=Request.Form("Phototypefyurl_s")
Phototypefyurl_o=Request.Form("Phototypefyurl_o")
Phototypeurl_s=Request.Form("Phototypeurl_s")
Phototypeurl_o=Request.Form("Phototypeurl_o")
IF Colleclx=2 Then
   Downlist_s=Request.Form("Downlist_s")
   Downlist_o=Request.Form("Downlist_o")
   DownUrl_s=Request.Form("DownUrl_s")
   DownUrl_o=Request.Form("DownUrl_o")
End If   
IF Colleclx=3 then
   Downlist_s=Request.Form("Downlist_s")'下载地址设置
   Downlist_o=Request.Form("Downlist_o")
   DownUrl_s=Request.Form("DownUrl_s")
   DownUrl_o=Request.Form("DownUrl_o")

   DownNewType=Request.Form("DownNewType")'新窗口打开下载连接
   DownNewlist_s=Request.Form("DownNewlist_s")
   DownNewlist_o=Request.Form("DownNewlist_o")
   DownNewUrl_s=Request.Form("DownNewUrl_s")
   DownNewUrl_o=Request.Form("DownNewUrl_o")
   
   ZdType_001=Request.Form("ZdType_001")'软件大小设置
   Zds_001=Request.Form("Zds_001")
   Zdo_001=Request.Form("Zdo_001")
  
   ZdType_002=Request.Form("ZdType_002")'软件语言设置
   Zds_002=Request.Form("Zds_002")
   Zdo_002=Request.Form("Zdo_002")
   
   ZdType_003=Request.Form("ZdType_003")'授权方式设置
   Zds_003=Request.Form("Zds_003")
   Zdo_003=Request.Form("Zdo_003")
   
   ZdType_004=Request.Form("ZdType_004")'运行环境设置
   Zds_004=Request.Form("Zds_004")
   Zdo_004=Request.Form("Zdo_004")
   
   ZdType_005=Request.Form("ZdType_005")'演示地址设置
   Zds_005=Request.Form("Zds_005")
   Zdo_005=Request.Form("Zdo_005")
   
   ZdType_006=Request.Form("ZdType_006")'注册地址设置
   Zds_006=Request.Form("Zds_006")
   Zdo_006=Request.Form("Zdo_006")
   
   ZdType_007=Request.Form("ZdType_007")'软件图片设置
   Zds_007=Request.Form("Zds_007")
   Zdo_007=Request.Form("Zdo_007")
   
   LinkUrlYn=Request.Form("LinkUrlYn")
   if LinkUrlYn="" then LinkUrlYn=0
End if

IF Colleclx=4 Then
   Downlist_s=Request.Form("Downlist_s")'下载地址设置
   Downlist_o=Request.Form("Downlist_o")
   DownUrl_s=Request.Form("DownUrl_s")
   DownUrl_o=Request.Form("DownUrl_o")
   Zds_007=Request.Form("Zds_007")'重定地址
   
   DownNewType=Request.Form("DownNewType")'新窗口打开下载连接
   DownNewlist_s=Request.Form("DownNewlist_s")
   DownNewlist_o=Request.Form("DownNewlist_o")
   DownNewUrl_s=Request.Form("DownNewUrl_s")
   DownNewUrl_o=Request.Form("DownNewUrl_o")
End if

DateType=Trim(Request.Form("DateType"))
DsString=Request.Form("DsString")
DoString=Request.Form("DoString")

AuthorType=Trim(Request.Form("AuthorType"))
AsString=Request.Form("AsString")
AoString=Request.Form("AoString")
AuthorStr=Request.Form("AuthorStr")

CopyFromType=Trim(Request.Form("CopyFromType"))
FsString=Request.Form("FsString")
FoString=Request.Form("FoString")
CopyFromStr=Request.Form("CopyFromStr")

KeyType=Trim(Request.Form("KeyType"))
KsString=Request.Form("KsString")
KoString=Request.Form("KoString")
KeyStr=Request.Form("KeyStr")

UrlTest=Trim(Request.Form("UrlTest"))

If TsString="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>标题开始标记不能为空</li>"
End If
If ToString="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>标题结束标记不能为空</li>" 
End If
If NewsPaingType="" Then

Else
   NewsPaingType=Clng(NewsPaingType)
   If NewsPaingType=0 Then
   ElseIf NewsPaingType=1 Then
		If NPsString="" or NPoString="" Then
				 FoundErr=True
				 ErrMsg=ErrMsg & "<br><li>正文分页列表开始/结束标不能为空！</li>" 
		End If
		If NewsUrlPaing_s="" or NewsUrlPaing_o="" Then
				 FoundErr=True
				 ErrMsg=ErrMsg & "<br><li>正文分页链接开始/结束标不能为空！！</li>" 
		End If
   ElseIf NewsPaingType=2 Then
   		
   Else	  	
   
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>参数错误，正文分类设置</li>" 
   End If 
End If

If DateType="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>请设置时间类型</li>" 
Else
   DateType=Clng(DateType)
   If DateType=0 Then
   ElseIf DateType=1 Then
      If DsString="" or DoString="" Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>请将时间的开始/结束标记填写完整</li>" 
      End If
   Else
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>参数错误，请从有效链接进入</li>" 
   End If
End If

If AuthorType="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>请设置作者类型</li>" 
Else
   AuthorType=Clng(AuthorType)
   If AuthorType=0 Then
   ElseIf AuthorType=1 Then
      If AsString="" or AoString="" Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>请将作者开始/结束标记填写完整！</li>" 
      End If
   ElseIf AuthorType=2 Then
      If AuthorStr="" Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>请指定作者</li>" 
      End If
   Else
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>参数错误，请从有效链接进入</li>" 
   End If 
End If

if Colleclx=3 then'下载地址设置
	If Downlist_s="" or  Downlist_o="" or DownUrl_s="" or DownUrl_o="" then
    	FoundErr=True
    	ErrMsg=ErrMsg & "<br><li>下载地址设置不能为空</li>" 
	End if
	IF DownNewType=1 then
		if DownNewlist_s="" or DownNewlist_o="" or DownNewUrl_s="" or DownNewUrl_o="" then
			FoundErr=True
    		ErrMsg=ErrMsg & "<br><li>新窗口打开下载连接不能为空</li>"
		end if 
	End if   
	IF ZdType_001=1  then
		if Zds_001="" or Zdo_001="" then
			FoundErr=True
    		ErrMsg=ErrMsg & "<br><li>软件大小设置不能为空</li>"
		end if 
	End if
	IF ZdType_002=1  then
		if Zds_002="" or Zdo_002="" then
			FoundErr=True
    		ErrMsg=ErrMsg & "<br><li>软件语言设置不能为空</li>" 
		end if
	End if
	IF ZdType_003=1 then
		if Zds_003="" or Zdo_003="" then
			FoundErr=True
    		ErrMsg=ErrMsg & "<br><li>授权方式设置不能为空</li>" 
		end if	
	End if
	IF ZdType_004=1 then
		if Zds_004="" or Zdo_004="" then
			FoundErr=True
    		ErrMsg=ErrMsg & "<br><li>运行环境设置不能为空</li>" 
		End if
	End if
	IF ZdType_005=1 then
		if Zds_005="" or Zdo_005="" then
			FoundErr=True
    		ErrMsg=ErrMsg & "<br><li>演示地址设置不能为空</li>" 
		end if
	End if
	IF ZdType_006=1 then
		if Zds_006="" or Zdo_006="" then
			FoundErr=True
    		ErrMsg=ErrMsg & "<br><li>注册地址设置不能为空</li>" 
		end if	
	End if
	IF ZdType_007=1 then
		if Zds_007="" or Zdo_007="" then
			FoundErr=True
    		ErrMsg=ErrMsg & "<br><li>软件图片设置不能为空</li>" 
		end if	
	End if
end if

If CopyFromType="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>请设置来源类型</li>" 
Else
   CopyFromType=Clng(CopyFromType)
   If CopyFromType=0 Then
   ElseIf CopyFromType=1 Then
      If FsString="" or FoString="" Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>请将来源开始/结束标记填写完整！</li>" 
      End If
   ElseIf CopyFromType=2 Then
      If CopyFromStr="" Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>请指定来源</li>" 
      End If
   Else
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>参数错误，请从有效链接进入</li>" 
   End If 
End If

If KeyType="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>请设置关键字类型</li>" 
Else
   KeyType=Clng(KeyType)
   If KeyType=0 Then
   ElseIf KeyType=1 Then
      If KsString="" or KoString="" Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>关键字开始/结束标记不能为空</li>" 
      End If
   ElseIf KeyType=2 Then
      If KeyStr="" Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>请指定关键字</li>" 
      End If
   Else
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>参数错误，请从有效链接进入</li>" 
   End If
End If

If FoundErr<>True Then
   SqlItem="Select * from Item Where ItemID=" & ItemID
   Set RsItem=server.CreateObject("adodb.recordset")
   RsItem.Open SqlItem,ConnItem,2,3
   RsItem("TsString")=TsString
   RsItem("ToString")=ToString
   RsItem("CsString")=CsString
   RsItem("CoString")=CoString
   
   RsItem("photourls")=photourls
   RsItem("photourlo")=photourlo
   if NewsPaingType<>"" then 
   		RsItem("NewsPaingType")=NewsPaingType
   		if NewsPaingType=1 then
   		RsItem("NPsString")=NPsString
  		RsItem("NPoString")=NPoString
   		RsItem("NewsUrlPaing_s")=NewsUrlPaing_s
   		RsItem("NewsUrlPaing_o")=NewsUrlPaing_o
   		end if
   end if
   
   IF Colleclx=2 Then
   		RsItem("PhotoType_s")=PhotoType_s
        RsItem("PhotoType_o")=PhotoType_o
		  
		RsItem("PhotoLurl_s")=PhotoLurl_s
		RsItem("PhotoLurl_o")=PhotoLurl_o
		 
		RsItem("Phototypefy_s")=Phototypefy_s
		RsItem("Phototypefy_o")=Phototypefy_o
		  
		RsItem("Phototypefyurl_s")=Phototypefyurl_s
		RsItem("Phototypefyurl_o")=Phototypefyurl_o
		  
		RsItem("Phototypeurl_s")=Phototypeurl_s
		RsItem("Phototypeurl_o")=Phototypeurl_o
		RsItem("Downlist_s")=Downlist_s'下载地址设置
  		RsItem("Downlist_o")=Downlist_o
   		RsItem("DownUrl_s")=DownUrl_s
   		RsItem("DownUrl_o")=DownUrl_o
   End If
   RsItem("DateType")=DateType
   if Colleclx=3 then  
   RsItem("Downlist_s")=Downlist_s'下载地址设置
   RsItem("Downlist_o")=Downlist_o
   RsItem("DownUrl_s")=DownUrl_s
   RsItem("DownUrl_o")=DownUrl_o
   
   RsItem("DownNewType")=DownNewType'新窗口打开下载连接
   RsItem("DownNewlist_s")=DownNewlist_s
   RsItem("DownNewlist_o")=DownNewlist_o
   RsItem("DownNewUrl_s")=DownNewUrl_s
   RsItem("DownNewUrl_o")=DownNewUrl_o
   
   RsItem("ZdType_001")=ZdType_001'软件大小设置
   RsItem("Zds_001")=Zds_001
   RsItem("Zdo_001")=Zdo_001
   
   RsItem("ZdType_002")=ZdType_002'软件语言设置
   RsItem("Zds_002")=Zds_002
   RsItem("Zdo_002")=Zdo_002
   
   RsItem("ZdType_003")=ZdType_003'授权方式设置
   RsItem("Zds_003")=Zds_003
   RsItem("Zdo_003")=Zdo_003
   
   RsItem("ZdType_004")=ZdType_004'运行环境设置
   RsItem("Zds_004")=Zds_004
   RsItem("Zdo_004")=Zdo_004
   
   RsItem("ZdType_005")=ZdType_005'演示地址设置
   RsItem("Zds_005")=Zds_005
   RsItem("Zdo_005")=Zdo_005
   
   RsItem("ZdType_006")=ZdType_006'注册地址设置
   RsItem("Zds_006")=Zds_006
   RsItem("Zdo_006")=Zdo_006
   
   RsItem("ZdType_007")=ZdType_007'注册地址设置
   RsItem("Zds_007")=Zds_007
   RsItem("Zdo_007")=Zdo_007
   RsItem("LinkUrlYn")= LinkUrlYn
   end if
   
   IF Colleclx=4 Then
   	  RsItem("Downlist_s")=Downlist_s'动漫地址设置
      RsItem("Downlist_o")=Downlist_o
      RsItem("DownUrl_s")=DownUrl_s
      RsItem("DownUrl_o")=DownUrl_o
	  RsItem("Zds_007")=Zds_007
	  
	  RsItem("DownNewType")=DownNewType'新窗口打开下载连接
   	  RsItem("DownNewlist_s")=DownNewlist_s
   	  RsItem("DownNewlist_o")=DownNewlist_o
      RsItem("DownNewUrl_s")=DownNewUrl_s
      RsItem("DownNewUrl_o")=DownNewUrl_o
   End if
   
   If DateType=1 Then
      RsItem("DsString")=DsString
      RsItem("DoString")=DoString
   End If

   RsItem("AuthorType")=AuthorType
   If AuthorType=1 Then
      RsItem("AsString")=AsString
      RsItem("AoString")=AoString
   ElseIf AuthorType=2 Then
      RsItem("AuthorStr")=AuthorStr
   End If

   RsItem("CopyFromType")=CopyFromType
   If CopyFromType=1 Then
      RsItem("FsString")=FsString
      RsItem("FoString")=FoString
   ElseIf CopyFromType=2 Then
      RsItem("CopyFromStr")=CopyFromStr
   End If

   RsItem("KeyType")=KeyType
   If KeyType=1 Then
      RsItem("KsString")=KsString
      RsItem("KoString")=KoString
   ElseIf KeyType=2 Then
      RsItem("KeyStr")=KeyStr
   End If
   RsItem.UpDate
   RsItem.Close
   Set RsItem=Nothing
End If
End Sub
'=============采集测试============================
Public Sub GetTest()
   SqlItem="Select * from Item Where ItemID=" & ItemID
   Set RsItem=server.CreateObject("adodb.recordset")
   RsItem.Open SqlItem,ConnItem,1,1
   If  RsItem.Eof  And  RsItem.Bof  Then
         FoundErr=True
      ErrMsg=ErrMsg  &  "<br><li>参数错误，找不到该项目</li>"
   Else
      ListStr=RsItem("ListStr")
      LsString=RsItem("LsString")
      LoString=RsItem("LoString")
      ListPaingType=RsItem("ListPaingType")
      LPsString=RsItem("LPsString")
      LPoString=RsItem("LPoString")
      ListPaingStr2=RsItem("ListPaingStr2")
      ListPaingID1=RsItem("ListPaingID1")
      ListPaingID2=RsItem("ListPaingID2")
      ListPaingStr3=RsItem("ListPaingStr3")
	  selEncoding=RsItem("selEncoding")
      strReplace=RsItem("strReplace")
      HsString=RsItem("HsString")
      HoString=RsItem("HoString")
	  x_tp=RsItem("x_tp")
	  imhstr=RsItem("imhstr")
      imostr=RsItem("imostr")
	  photourls=RsItem("photourls")
	  photourlo=RsItem("photourlo")
	  PhotoPaingType=RsItem("PhotoPaingType")
      PhotoType_s=RsItem("PhotoType_s")
      PhotoType_o=RsItem("PhotoType_o")
	  PhotoLurl_s=RsItem("PhotoLurl_s")
	  PhotoLurl_o=RsItem("PhotoLurl_o")
	  Phototypefy_s=RsItem("Phototypefy_s")
	  Phototypefy_o=RsItem("Phototypefy_o")
	  Phototypefyurl_s=RsItem("Phototypefyurl_s")
	  Phototypefyurl_o=RsItem("Phototypefyurl_o")
	  Phototypeurl_s=RsItem("Phototypeurl_s")
	  Phototypeurl_o=RsItem("Phototypeurl_o")
	  HttpUrlType=RsItem("HttpUrlType")
	  HttpUrlStr=RsItem("HttpUrlStr")
	  Colleclx=RsItem("Colleclx")
	  IF Colleclx=2 then
	   Downlist_s=RsItem("Downlist_s")'下载地址设置
	   Downlist_o=RsItem("Downlist_o")
	   DownUrl_s=RsItem("DownUrl_s")
	   DownUrl_o=RsItem("DownUrl_o")
	  End If

   	  If  Colleclx=3 then
	   Downlist_s=RsItem("Downlist_s")'下载地址设置
	   Downlist_o=RsItem("Downlist_o")
	   DownUrl_s=RsItem("DownUrl_s")
	   DownUrl_o=RsItem("DownUrl_o")
	   
	   DownNewType=RsItem("DownNewType")'新窗口打开下载连接
   	   DownNewlist_s=RsItem("DownNewlist_s")
   	   DownNewlist_o=RsItem("DownNewlist_o")
   	   DownNewUrl_s=RsItem("DownNewUrl_s")
   	   DownNewUrl_o=RsItem("DownNewUrl_o")
	   
	   ZdType_001=RsItem("ZdType_001")'软件大小设置
	   Zds_001=RsItem("Zds_001")
	   Zdo_001=RsItem("Zdo_001")
	   
	   ZdType_002=RsItem("ZdType_002")'软件语言设置
	   Zds_002=RsItem("Zds_002")
	   Zdo_002=RsItem("Zdo_002")
	   
	   ZdType_003=RsItem("ZdType_003")'授权方式设置
	   Zds_003=RsItem("Zds_003")
	   Zdo_003=RsItem("Zdo_003")
	   
	   ZdType_004=RsItem("ZdType_004")'运行环境设置
	   Zds_004=RsItem("Zds_004")
	   Zdo_004=RsItem("Zdo_004")
	   
	   ZdType_005=RsItem("ZdType_005")'演示地址设置
	   Zds_005=RsItem("Zds_005")
	   Zdo_005=RsItem("Zdo_005")
	   
	   ZdType_006=RsItem("ZdType_006")'注册地址设置
	   Zds_006=RsItem("Zds_006")
	   Zdo_006=RsItem("Zdo_006")
	   
	   ZdType_007=RsItem("ZdType_007")'软件图片设置
	   Zds_007=RsItem("Zds_007")
	   Zdo_007=RsItem("Zdo_007") 
	  end if
	  If  Colleclx=4 then
	   Downlist_s=RsItem("Downlist_s")'动漫地址设置
	   Downlist_o=RsItem("Downlist_o")
	   DownUrl_s=RsItem("DownUrl_s")
	   DownUrl_o=RsItem("DownUrl_o")
	   Zds_007=RsItem("Zds_007")
	   
	   DownNewType=RsItem("DownNewType")'新窗口打开下载连接
   	   DownNewlist_s=RsItem("DownNewlist_s")
   	   DownNewlist_o=RsItem("DownNewlist_o")
   	   DownNewUrl_s=RsItem("DownNewUrl_s")
   	   DownNewUrl_o=RsItem("DownNewUrl_o")
	  End if
      TsString=RsItem("TsString")
      ToString=RsItem("ToString")
      CsString=RsItem("CsString")
      CoString=RsItem("CoString")
      
      DateType=RsItem("DateType")
      DsString=RsItem("DsString")
      DoString=RsItem("DoString")

      AuthorType=RsItem("AuthorType")
      AsString=RsItem("AsString")
      AoString=RsItem("AoString")
      AuthorStr=RsItem("AuthorStr")

      CopyFromType=RsItem("CopyFromType")
      FsString=RsItem("FsString")
      FoString=RsItem("FoString")
      CopyFromStr=RsItem("CopyFromStr")

      KeyType=RsItem("KeyType")
      KsString=RsItem("KsString")
      KoString=RsItem("KoString")
      KeyStr=RsItem("KeyStr")

      NewsPaingType=RsItem("NewsPaingType")
      NPsString=RsItem("NPsString")
      NPoString=RsItem("NPoString")
      NewsUrlPaing_s=RsItem("NewsUrlPaing_s")
      NewsUrlPaing_o=RsItem("NewsUrlPaing_o")
	  
	  Script_Iframe=RsItem("Script_Iframe")
	  Script_Object=RsItem("Script_Object")
	  Script_Script=RsItem("Script_Script")
	  Script_Div=RsItem("Script_Div")
	  Script_Class=RsItem("Script_Class")
	  Script_Span=RsItem("Script_Span")
	  Script_Img=RsItem("Script_Img")
	  Script_Font=RsItem("Script_Font")
	  Script_A=RsItem("Script_A")
	  Script_Html=RsItem("Script_Html")	  
      
      UpDateType=RsItem("UpDateType")
	  Colleclx=RsItem("Colleclx")
   End  If   
   RsItem.Close
   Set RsItem=Nothing

   
   If LsString="" Then
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>列表开始标记不能为空！</li>"
   End If
   If LoString="" Then
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>列表结束标记不能为空！</li>"
   End If
   ListPaingType=Clng(ListPaingType)
   Select Case ListPaingType
   Case 0
   Case 1
      If ListPaingStr2="" Then
         FoundErr=True
		 
         ErrMsg=ErrMsg & "<br><li>批量生成字符不能为空</li>"
      End If
      If isNumeric(ListPaingID1)=False or isNumeric(ListPaingID2)=False Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>批量生成的范围只能是数字</li>"
      Else
         ListPaingID1=Clng(ListPaingID1)
         ListPaingID2=Clng(ListPaingID2)
         If ListPaingID1=0 And ListPaingID2=0 Then
            FoundErr=True
            ErrMsg=ErrMsg & "<br><li>批量生成范围设置不正确</li>"
         End If
      End If
   Case 2
      If ListPaingStr3="" Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>列表索引分页不能为空，请手动添加</li>"
      Else
         ListPaingStr3=Replace(ListPaingStr3,CHR(13),"|") 
      End If
   Case 3
	  IF LPsString="" or LPoString="" then
	  	 FoundErr=True
         ErrMsg=ErrMsg & "<br><li>索引分页开始/结束标记不能为空!</li>"
	  End if 	  
   Case Else
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>请选择列表索引分页类型</li>" 
   End Select
  
 If  HsString=""  or  HoString=""  Then
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>链接开始/结束标记不能为空！</li>"
      End  If

   If FoundErr<>True  Then
      Select Case ListPaingType
      Case 0
            ListUrl=ListStr
      Case 1
         ListUrl=Replace(ListPaingStr2,"{$ID}",CStr(ListPaingID1))
      Case 2
         If Instr(ListPaingStr3,"|")> 0 Then
            ListUrl=Left(ListPaingStr3,Instr(ListPaingStr3,"|")-1)
         Else
            ListUrl=ListPaingStr3
         End If
	  Case 3
	  	 ListUrl=ListStr		 
      End Select
   End If

   If  FoundErr<>True   Then
         ListCode=Skcj.ReplaceTrim(Skcj.GetHttpPage(ListUrl,selEncoding))
         If ListCode<>"$False$" Then
            ListCode=Skcj.GetBody(ListCode,LsString,LoString,False,False)
            If ListCode<>"$False$" Then
               NewsArrayCode=Skcj.GetArray(ListCode,HsString,HoString,False,False)
               If NewsArrayCode<>"$False$" Then
                  If Instr(NewsArrayCode,"$Array$")>0 Then
                     NewsArray=Split(NewsArrayCode,"$Array$")
                     If HttpUrlType=1 Then
                        NewsUrl=Trim(Replace(HttpUrlStr,"{$ID}",NewsArray(0)))
                     Else
                        NewsUrl=Trim(Skcj.FormatRemoteUrl(NewsArray(0),ListUrl))
                     End If
                  Else
                     FoundErr=True
                     ErrMsg=ErrMsg & "<br><li>只发现一个有效链接？：" & NewsArrayCode & "</li>"
                 End If
              Else
                 FoundErr=True
                 ErrMsg=ErrMsg & "<br><li>在获取链接列表时出错。</li>"
              End If   
           Else
               FoundErr=True
              ErrMsg=ErrMsg & "<br><li>在截取列表时发生错误。</li>"
           End If
        Else
            FoundErr=True
           ErrMsg=ErrMsg & "<br><li>在获取:" & ListUrl & "网页源码时发生错误。</li>"
        End If
     End If
'------------列表小图片	 
If x_tp=1 then
	If FoundErr<>True Then
	   dim NewsimageCode
	   NewsimageCode=Skcj.GetArray(ListCode,imhstr,imostr,False,False)
	   If NewsimageCode="$False$" Then
		  FoundErr=True
		  ErrMsg=ErrMsg & "<br><li>在分析：" & ListUrl & "小图片列表时发生错误！</li>"
	   Else
		  Newsimage=Split(NewsimageCode,"$Array$")
		  dim Arr_i
		  For Arr_i=0 to Ubound(Newsimage)
			 If HttpUrlType=1 Then
				Newsimage(Arr_i)=Trim(Replace(HttpUrlStr,"{$ID}",Newsimage(Arr_i)))
			 Else
				Newsimage(Arr_i)=Trim(Skcj.FormatRemoteUrl(Newsimage(Arr_i),ListUrl))           
			 End If
			 'if x_tpUrl<>"" then Newsimage(Arr_i)= x_tpUrl & Newsimage(Arr_i) 
			 Newsimage(Arr_i)=CheckUrl(Newsimage(Arr_i))
			 picpath=Newsimage(arr_i)
			 iF SaveFiles=1 then picpath=Skcj.Sk_SaveFile(Colleclx,picpath)
		  Next
	   End If
	End If
End if	      
If FoundErr<>True Then
   NewsCode=Skcj.ReplaceTrim(Skcj.GetHttpPage(NewsUrl,selEncoding))
   If NewsCode<>"$False$" Then
	  Title=Skcj.GetBody(NewsCode,TsString,ToString,False,False)
	  if CsString<>"" or CoString<>"" then
		  if CsString<>"0" or  CoString<>"0" then
			  Content=SKcj.GetBody(NewsCode,CsString,CoString,False,False)
		  Else
			  Content=""
		  end if
	  end if
      If Title="$False$"  Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>在截取"& ErrMsg_lx &"标题/正文的时候发生错误：" & NewsUrl & "</li>"
      Else
         Title=FpHtmlEnCode(Title)
         Title=dvhtmlencode(Title)
		 Call FilterScript()
         Content=Ubbcode(Content)
If Colleclx=2 then '图片		 
		If NewsPaingType=2 Then '2级分类链接源码时发生错误
			ListTypeCode=Skcj.GetBody(NewsCode,PhotoType_s,PhotoType_o,False,False)
			If ListTypeCode<>"$False$" Then 
				ListTypeUrlCode=Skcj.GetArray(ListTypeCode,PhotoLurl_s,PhotoLurl_o,False,False)
				If ListTypeUrlCode<>"$False$" and Instr(NewsArrayCode,"$Array$")>0  Then 
					TypeUrlArray=Split(ListTypeUrlCode,"$Array$")
					TypeNewsUrl=Trim(Skcj.FormatRemoteUrl(TypeUrlArray(0),NewsUrl))
					if TypeNewsUrl<>"$False$" then
						if Phototypeurl_s<>"0" or Phototypeurl_o<>"0" then
							NewsTypeCode=Skcj.ReplaceTrim(Skcj.GetHttpPage(TypeNewsUrl,selEncoding))
							PicUrls=Skcj.GetBody(NewsTypeCode,Phototypeurl_s,Phototypeurl_o,False,False)
							if HttpUrlStr<>"" then
								PicUrls=Trim(Skcj.FormatRemoteUrl(PicUrls,HttpUrlStr))
							else
								PicUrls=Trim(Skcj.FormatRemoteUrl(PicUrls,TypeNewsUrl))
							end if
						else
							PicUrls=TypeNewsUrl
						end if
					end if
				Else
					FoundErr=True
					ErrMsg=ErrMsg & "<br><li>在获取:" & NewsUrl & "2级分类链接源码时发生错误。</li>"	
				End if
			Else
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>在获取:" & NewsUrl & "2级分类列表源码时发生错误。</li>"
			End if
		Else
			If Downlist_s<>"0" And  Downlist_o<>"0" And DownUrl_s<>"0" And DownUrl_o<>"0" Then
				If Downlist_s="" or  Downlist_o="" or DownUrl_s="" or DownUrl_o="" then'图片下载
					FoundErr=True
					ErrMsg=ErrMsg & "<br><li>图片地址设置不能为空</li>" 
				Else	
					DownUrls=Skcj.GetBody(NewsCode,Downlist_s,Downlist_o,False,False)
					If DownUrls<>"$False$" then
						DownUrls=Skcj.GetBody(DownUrls,DownUrl_s,DownUrl_o,False,False)
						IF DownUrls<>"$False$" then	
							DownUrls=Trim(Skcj.FormatRemoteUrl(DownUrls,NewsUrl))
							PicUrls=DownUrls
						Else
							FoundErr=True
							ErrMsg=ErrMsg & "<br><li>在获取:" & NewsUrl & "图片链接源码时发生错误。</li>"
						End if
					Else
						FoundErr=True
						ErrMsg=ErrMsg & "<br><li>在获取:" & NewsUrl & "图片列表源码时发生错误。</li>"
					End if
				End if
			End if
        End If	
End If	

If Colleclx=3 then '下载
  	If Downlist_s="" or  Downlist_o="" or DownUrl_s="" or DownUrl_o="" then'下载地址设置
    	FoundErr=True
    	ErrMsg=ErrMsg & "<br><li>下载地址设置不能为空</li>" 
	Else	
		DownUrls=Skcj.GetBody(NewsCode,Downlist_s,Downlist_o,False,False)
		If DownUrls<>"$False$" then
			DownUrls=Skcj.GetBody(DownUrls,DownUrl_s,DownUrl_o,False,False)
			IF DownUrls<>"$False$" then
				 	DownUrls=Trim(Skcj.FormatRemoteUrl(DownUrls,NewsUrl))
			Else
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>在获取:" & NewsUrl & "下载地址链接源码时发生错误。</li>"
			End if
		Else
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>在获取:" & NewsUrl & "下载地址列表源码时发生错误。</li>"
		End if
	End if

	If DownNewType=1 then'新窗口打开下载连接
		If DownNewlist_s<>"" or  DownNewlist_o<>"" or DownNewUrl_s<>"" or DownNewUrl_o<>"" then
			DownUrls=Skcj.ReplaceTrim(Skcj.GetHttpPage(DownUrls,selEncoding))
			DownUrls=Skcj.GetBody(DownUrls,DownNewlist_s,DownNewlist_o,False,False)
			If DownUrls<>"$False$" then 
				DownUrls=Skcj.GetBody(DownUrls,DownNewUrl_s,DownNewUrl_o,False,False)
				IF DownUrls<>"$False$" then
					DownUrls=Trim(Skcj.FormatRemoteUrl(DownUrls,NewsUrl)) 
				Else
					FoundErr=True
    				ErrMsg=ErrMsg & "<br><li>在获取:" & NewsUrl & "(新窗口)下载地址列表源码时发生错误</li>" 
				End if
			Else
				FoundErr=True
    			ErrMsg=ErrMsg & "<br><li>在获取:" & NewsUrl & "(新窗口)下载地址列表源码时发生错误</li>" 
			End if
		Else
			FoundErr=True
    		ErrMsg=ErrMsg & "<br><li>下载地址设置不能为空</li>" 
		End if
	End if 

	If ZdType_001=0 then'软件大小设置
   		DownSize=""
	Else
		If  Zds_001="0" and Zdo_001<>"" then
			DownSize=Zdo_001
		Else
			DownSize=Skcj.GetBody(NewsCode,Zds_001,Zdo_001,False,False)
		End If
	End If
	
	If ZdType_002=0 then'软件语言设置
   		DownYY=""
	Else
		If  Zds_002="0" and Zdo_002<>"" then
			DownYY=Zdo_002
		Else
			DownYY=Skcj.GetBody(NewsCode,Zds_002,Zdo_002,False,False)
		End If
	End If
 	If ZdType_003=0 then'授权方式设置
   		DownSQ=""
	Else
		If  Zds_003="0" and Zdo_003<>"" then
			DownSQ=Zdo_003
		Else
			DownSQ=Skcj.GetBody(NewsCode,Zds_003,Zdo_003,False,False)
		End If
	End If
 	If ZdType_004=0 then'运行环境设置
   		DownPT=""
	Else
		If  Zds_004="0" and Zdo_004<>"" then
			DownPT=Zdo_004
		Else
			DownPT=Skcj.GetBody(NewsCode,Zds_004,Zdo_004,False,False)
		End If
	End If
 	If ZdType_005=0 then'演示地址设置
   		YSDZ=""
	Else
		If  Zds_005="0" and Zdo_005<>"" then
			YSDZ=Zdo_005
		Else
			YSDZ=Skcj.GetBody(NewsCode,Zds_005,Zdo_005,False,False)
		End If
	End If
 	If ZdType_006=0 then'注册地址设置
   		ZCDZ=""
	Else
		If  Zds_006="0" and Zdo_006<>"" then
			ZCDZ=Zdo_006
		Else
			ZCDZ=Skcj.GetBody(NewsCode,Zds_006,Zdo_006,False,False)
		End If
	End If	
 	If ZdType_007=0 then'软件图片设置
   		PhotoUrl=""
	Else
		If  Zds_007="0" and Zdo_007<>"" then
			PhotoUrl=Zdo_007
		Else
			PhotoUrl=Skcj.GetBody(NewsCode,Zds_007,Zdo_007,False,False)
			PhotoUrl=Trim(Skcj.FormatRemoteUrl(PhotoUrl,NewsUrl))
		End If
	End If  
End if

IF Colleclx=4 then '动漫
	If Downlist_s="" or  Downlist_o="" or DownUrl_s="" or DownUrl_o="" then'下载地址设置
    	FoundErr=True
    	ErrMsg=ErrMsg & "<br><li>动漫下载地址设置不能为空</li>" 
	Else	
		DownUrls=Skcj.GetBody(NewsCode,Downlist_s,Downlist_o,False,False)
		If DownUrls<>"$False$" then
			DownUrls=Skcj.GetBody(DownUrls,DownUrl_s,DownUrl_o,False,False)
			IF DownUrls<>"$False$" then
				DownUrls=Trim(Skcj.FormatRemoteUrl(DownUrls,NewsUrl))
			Else
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>在获取:" & NewsUrl & "动漫下载地址链接源码时发生错误。</li>"
			End if
		Else
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>在获取:" & NewsUrl & "动漫下载地址列表源码时发生错误。</li>"
		End if
	End if
	
	If DownNewType=1 then'新窗口打开下载连接
		If DownNewlist_s<>"" or  DownNewlist_o<>"" or DownNewUrl_s<>"" or DownNewUrl_o<>"" then
			DownUrls=Skcj.ReplaceTrim(Skcj.GetHttpPage(DownUrls,selEncoding))
			DownUrls=Skcj.GetBody(DownUrls,DownNewlist_s,DownNewlist_o,False,False)
			If DownUrls<>"$False$" then 
				DownUrls=Skcj.GetBody(DownUrls,DownNewUrl_s,DownNewUrl_o,False,False)
				IF DownUrls<>"$False$" then
				 	DownUrls=Trim(Skcj.FormatRemoteUrl(DownUrls,NewsUrl)) 
				Else
					FoundErr=True
    				ErrMsg=ErrMsg & "<br><li>在获取:" & NewsUrl & "(新窗口)动漫地址列表源码时发生错误</li>" 
				End if
			Else
				FoundErr=True
    			ErrMsg=ErrMsg & "<br><li>在获取:" & NewsUrl & "(新窗口)动漫地址列表源码时发生错误</li>" 
			End if
		Else
			FoundErr=True
    		ErrMsg=ErrMsg & "<br><li>动漫地址设置不能为空</li>" 
		End if
	End if 
End if

         If DateType=0 Then
            UpDateTime=Now()
         Else
		 	If DateType=1 Then
               UpDateTime=Skcj.GetBody(NewsCode,DsString,DoString,False,False)
               UpDateTime=FpHtmlEncode(UpDateTime)
               If IsDate(UpDateTime)=True Then
                  UpDateTime=CDate(UpDateTime)
               Else
                  UpDateTime=Now()
               End If
            End If
         End If
         If AuthorType=1 Then
            Author=Skcj.GetBody(NewsCode,AsString,AoString,False,False)
         ElseIf AuthorType=2 Then
            Author=AuthorStr
         End If
         If Author="$False$" Or Trim(Author)="" Then
            Author="佚名"
         Else
            Author=FpHtmlEnCode(Author)
         End If
         If CopyFromType=1 Then
            CopyFrom=Skcj.GetBody(NewsCode,FsString,FoString,False,False)
         ElseIf CopyFromType=2 Then
            CopyFrom=CopyFromStr
         End If
         If CopyFrom="$False$" Or Trim(CopyFrom)="" Then
            CopyFrom="不详"
         Else
            CopyFrom=FpHtmlEnCode(CopyFrom)
         End If

         If KeyType=0 Then
            Key=Title
            Key=CreateKeyWord(Key,2)
         ElseIf KeyType=1 Then
            Key=Skcj.GetBody(NewsCode,KsString,KoString,False,False)
            Key=Replace(Key,",","|")
			Key=Replace(Key," ","|")
			Key=FpHtmlEnCode(Key)
            'Key=CreateKeyWord(Key,2)
         ElseIf KeyType=2 Then
            Key=FpHtmlEnCode(Key)
         End If
         If Key="$False$" Or Trim(Key)="" Then
            Key=KeyStr
         End If	 
     End  If    
   Else
     FoundErr=True
     ErrMsg=ErrMsg & "<br><li>在获取源码时发生错误："& NewsUrl &"</li>"
   End If 
End If
if PicUrls="$False$" then 
	PicUrls=""
end if

If FoundErr<>True Then
   Call GetFilters'过滤
   Call Filters'过滤
   Content=Skcj.ItemReplaceStr(Content,strReplace)'内容替换
   if  Colleclx =1 then 
  	set Rs = ConnItem.execute("select top 1 Dir from SK_cj where ID=1")
			Content=Skcj.ReplaceSaveRemoteFile(Content,"/",rs("Dir"),False,NewsUrl)'远程图片
			Content=Skcj.ReSaveRemoteFile(Content,NewsUrl,rs("Dir"),False)'远程文件
	Rs.close : set Rs=nothing
   End if
   'Content=SKcj.ReplaceSaveRemoteFile(Content,strInstallDir,strChannelDir,False,NewsUrl)'远程图片
   'Content=Skcj.ReSaveRemoteFile(Content,NewsUrl,strChannelDir,False)'远程文件
End If
if FoundErr<>True then
	set RsItem=ConnItem.execute("update Item set Flag=True where ItemID="& ItemID)
end if
End Sub

Sub SetItems()
      Dim ItemNumTemp
      ItemNumTemp=ItemNum-1
      ItemID=Arr_Item(0,ItemNumTemp)
      ItemName=Arr_Item(1,ItemNumTemp)
      ChannelID=Arr_Item(2,ItemNumTemp)'频道ID
      strChannelDir=Arr_Item(3,ItemNumTemp)'频道目录
      ClassID=Arr_Item(4,ItemNumTemp)            '栏目
      SpecialID=Arr_Item(5,ItemNumTemp)        '专题
      LoginType=Arr_Item(9,ItemNumTemp)
      LoginUrl=Arr_Item(10,ItemNumTemp)          '登录
      LoginPostUrl=Arr_Item(11,ItemNumTemp)
      LoginUser=Arr_Item(12,ItemNumTemp)
      LoginPass=Arr_Item(13,ItemNumTemp)
      LoginFalse=Arr_Item(14,ItemNumTemp)
      ListStr=Arr_Item(15,ItemNumTemp)            '列表地址
      LsString=Arr_Item(16,ItemNumTemp)          '列表
      LoString=Arr_Item(17,ItemNumTemp)
      ListPaingType=Arr_Item(18,ItemNumTemp)
      LPsString=Arr_Item(19,ItemNumTemp)          
      LPoString=Arr_Item(20,ItemNumTemp)
      ListPaingStr1=Arr_Item(21,ItemNumTemp)
      ListPaingStr2=Arr_Item(22,ItemNumTemp)
      ListPaingID1=Arr_Item(23,ItemNumTemp)
      ListPaingID2=Arr_Item(24,ItemNumTemp)
      ListPaingStr3=Arr_Item(25,ItemNumTemp)
      HsString=Arr_Item(26,ItemNumTemp)  
      HoString=Arr_Item(27,ItemNumTemp)
      HttpUrlType=Arr_Item(28,ItemNumTemp)
      HttpUrlStr=Arr_Item(29,ItemNumTemp)
      TsString=Arr_Item(30,ItemNumTemp)          '标题
      ToString=Arr_Item(31,ItemNumTemp)
      CsString=Arr_Item(32,ItemNumTemp)          '正文
      CoString=Arr_Item(33,ItemNumTemp)
      DateType=Arr_Item(34,ItemNumTemp)      '作者
      DsString=Arr_Item(35,ItemNumTemp)          
      DoString=Arr_Item(36,ItemNumTemp)
      AuthorType=Arr_Item(37,ItemNumTemp)      '作者
      AsString=Arr_Item(38,ItemNumTemp)          
      AoString=Arr_Item(39,ItemNumTemp)
      AuthorStr=Arr_Item(40,ItemNumTemp)
      CopyFromType=Arr_Item(41,ItemNumTemp)  '来源
      FsString=Arr_Item(42,ItemNumTemp)          
      FoString=Arr_Item(43,ItemNumTemp)
      CopyFromStr=Arr_Item(44,ItemNumTemp)
      KeyType=Arr_Item(45,ItemNumTemp)            '关键词
      KsString=Arr_Item(46,ItemNumTemp)          
      KoString=Arr_Item(47,ItemNumTemp)
      KeyStr=Arr_Item(48,ItemNumTemp)
      NewsPaingType=Arr_Item(49,ItemNumTemp)            '关键词
      NPsString=Arr_Item(50,ItemNumTemp)          
      NPoString=Arr_Item(51,ItemNumTemp)
      NewsUrlPaing_s=Arr_Item(52,ItemNumTemp)
      NewsUrlPaing_o=Arr_Item(53,ItemNumTemp)
      PaginationType=Arr_Item(55,ItemNumTemp)
      MaxCharPerPage=Arr_Item(56,ItemNumTemp)
      ReadLevel=Arr_Item(57,ItemNumTemp)
      Stars=Arr_Item(58,ItemNumTemp)
      ReadPoint=Arr_Item(59,ItemNumTemp)
      Hits=Arr_Item(60,ItemNumTemp)
      UpDateType=Arr_Item(61,ItemNumTemp)
      UpDateTime=Arr_Item(62,ItemNumTemp)
      IncludePicYn=Arr_Item(63,ItemNumTemp)
      DefaultPicYn=Arr_Item(64,ItemNumTemp)
      OnTop=Arr_Item(65,ItemNumTemp)
      Elite=Arr_Item(66,ItemNumTemp)
      Hot=Arr_Item(67,ItemNumTemp)
      SkinID=Arr_Item(68,ItemNumTemp)
      TemplateID=Arr_Item(69,ItemNumTemp)
      Script_Iframe=Arr_Item(70,ItemNumTemp)
      Script_Object=Arr_Item(71,ItemNumTemp)
      Script_Script=Arr_Item(72,ItemNumTemp)
      Script_Div=Arr_Item(73,ItemNumTemp)
      Script_Class=Arr_Item(74,ItemNumTemp)
      Script_Span=Arr_Item(75,ItemNumTemp)
      Script_Img=Arr_Item(76,ItemNumTemp)
      Script_Font=Arr_Item(77,ItemNumTemp)
      Script_A=Arr_Item(78,ItemNumTemp)
      Script_Html=Arr_Item(79,ItemNumTemp)
      CollecListNum=Arr_Item(80,ItemNumTemp)
      CollecNewsNum=Arr_Item(81,ItemNumTemp)
      Passed=Arr_Item(82,ItemNumTemp)
      SaveFiles=Arr_Item(83,ItemNumTemp)
      CollecOrder=Arr_Item(84,ItemNumTemp)
      LinkUrlYn=Arr_Item(85,ItemNumTemp)
      InputerType=Arr_Item(86,ItemNumTemp)
      Inputer=Arr_Item(87,ItemNumTemp)
      EditorType=Arr_Item(88,ItemNumTemp)
      Editor=Arr_Item(89,ItemNumTemp)
      ShowCommentLink=Arr_Item(90,ItemNumTemp)
      Script_Table=Arr_Item(91,ItemNumTemp)
      Script_Tr=Arr_Item(92,ItemNumTemp)
      Script_Td=Arr_Item(93,ItemNumTemp)
	  imhstr=Arr_Item(94,ItemNumTemp)
      imostr=Arr_Item(95,ItemNumTemp)
	  photourls=Arr_Item(96,ItemNumTemp)
      photourlo=Arr_Item(97,ItemNumTemp)
	  PhotoPaingType=Arr_Item(98,ItemNumTemp)
	  PhotoType_s=Arr_Item(99,ItemNumTemp)
	  PhotoType_o=Arr_Item(100,ItemNumTemp)
	  PhotoLurl_s=Arr_Item(101,ItemNumTemp)
	  PhotoLurl_o=Arr_Item(102,ItemNumTemp)
	  Phototypefy_s=Arr_Item(103,ItemNumTemp)
	  Phototypefy_o=Arr_Item(104,ItemNumTemp)
	  Phototypefyurl_s=Arr_Item(105,ItemNumTemp)
	  Phototypefyurl_o=Arr_Item(106,ItemNumTemp)
	  Phototypeurl_s=Arr_Item(107,ItemNumTemp)
	  Phototypeurl_o=Arr_Item(108,ItemNumTemp)
	  Colleclx=Arr_Item(109,ItemNumTemp)
	  x_tp=Arr_Item(110,ItemNumTemp)
	  selEncoding=Arr_Item(111,ItemNumTemp)
	  SaveFileUrl=Arr_Item(112,ItemNumTemp)
	  x_tpUrl=Arr_Item(113,ItemNumTemp)
	  If Colleclx=2 Then
	  	  Downlist_s=Arr_Item(114,ItemNumTemp)
		  Downlist_o=Arr_Item(115,ItemNumTemp)
		  DownUrl_s=Arr_Item(116,ItemNumTemp)
		  DownUrl_o=Arr_Item(117,ItemNumTemp)
	  End If
	  IF Colleclx=4 then
	  	  Downlist_s=Arr_Item(114,ItemNumTemp)
		  Downlist_o=Arr_Item(115,ItemNumTemp)
		  DownUrl_s=Arr_Item(116,ItemNumTemp)
		  DownUrl_o=Arr_Item(117,ItemNumTemp) 
		  DownNewlist_s=Arr_Item(119,ItemNumTemp)
	  End if
	  if Colleclx=3 then
		  Downlist_s=Arr_Item(114,ItemNumTemp)
		  Downlist_o=Arr_Item(115,ItemNumTemp)
		  DownUrl_s=Arr_Item(116,ItemNumTemp)
		  DownUrl_o=Arr_Item(117,ItemNumTemp)
		  DownNewType=Arr_Item(118,ItemNumTemp)
		  DownNewlist_s=Arr_Item(119,ItemNumTemp)
		  DownNewlist_o=Arr_Item(120,ItemNumTemp)
		  DownNewUrl_s=Arr_Item(121,ItemNumTemp)
		  DownNewUrl_o=Arr_Item(122,ItemNumTemp)
		  ZdType_001=Arr_Item(123,ItemNumTemp)
		  Zds_001=Arr_Item(124,ItemNumTemp)
		  Zdo_001=Arr_Item(125,ItemNumTemp)
		  ZdType_002=Arr_Item(126,ItemNumTemp)
		  Zds_002=Arr_Item(127,ItemNumTemp)
		  Zdo_002=Arr_Item(128,ItemNumTemp)
		  ZdType_003=Arr_Item(129,ItemNumTemp)
		  Zds_003=Arr_Item(130,ItemNumTemp)
		  Zdo_003=Arr_Item(131,ItemNumTemp)
		  ZdType_004=Arr_Item(132,ItemNumTemp)
		  Zds_004=Arr_Item(133,ItemNumTemp)
		  Zdo_004=Arr_Item(134,ItemNumTemp)
		  ZdType_005=Arr_Item(135,ItemNumTemp)
		  Zds_005=Arr_Item(136,ItemNumTemp)
		  Zdo_005=Arr_Item(137,ItemNumTemp)
		  ZdType_006=Arr_Item(138,ItemNumTemp)
		  Zds_006=Arr_Item(139,ItemNumTemp)
		  Zdo_006=Arr_Item(140,ItemNumTemp)
		  ZdType_007=Arr_Item(141,ItemNumTemp)
		  Zds_007=Arr_Item(142,ItemNumTemp)
		  Zdo_007=Arr_Item(143,ItemNumTemp)	
		  ZdType_008=Arr_Item(144,ItemNumTemp)
		  Zds_008=Arr_Item(145,ItemNumTemp)
		  Zdo_008=Arr_Item(146,ItemNumTemp)	 
  	  End if
	  Thumb_WaterMark=Arr_Item(147,ItemNumTemp)
	  Thumbs_Create=Arr_Item(148,ItemNumTemp)
	  strReplace=Arr_Item(150,ItemNumTemp)	
      If InputerType=1 Then
         Inputer=FpHtmlEnCode(Inputer)
      Else
         Inputer=session("AdminName")
      End If
      If EditorType=1 Then
         Editor=FpHtmlEnCode(Editor)
      Else
         Editor=session("AdminName")
      End If
      If IsObjInstalled("Scripting.FileSystemObject")=False or strChannelDir="" Then
         SaveFiles=False
      End if

End Sub

Sub StartCollection'开始采集
IF Colleclx <> 0 then 
	Set Rs = ConnItem.execute("Select top 1 Dir,MaxFileSize,FileExtName,Timeout from SK_Cj where ID="& Colleclx )
Else
	Set Rs = ConnItem.execute("Select top 1 Dir,MaxFileSize,FileExtName,Timeout from SK_Cj where ID=1" )
End if
Skcj.CjTimeout=Rs("Timeout") 
Skcj.DownExtName=Rs("FileExtName")
Skcj.MaxSize=Rs("MaxFileSize")
Rs.close : Set Rs=Nothing
If NewsSuccesNum >= CollecNewsNum And CollecNewsNum<>0 then 
	 If Itemon="" then
	  	  if Collecdate<>"" then
			 response.write("<script>location.href='sk_Timing.asp?action=GoTiming&Collecdate="&  Day(now()) &"';</script>")
		  Else
			  Response.Write "<br> &nbsp;&nbsp;&nbsp;&nbsp;采集完成，正在整理数据请稍后..." 
			  Response.Write  "<meta http-equiv=""refresh"" content=""1;url="& Skcj.GetItemConfig("FileName",Colleclx) &""">"
	  	  End if
	  Else
		 response.write  "<script>location.href='Sk_Collection.asp?ItemID="& ItemID &"&ItemNum=1&ListNum=1&NewsSuccesNum=0&NewsFalseNum=0&ImagesNumAll=0&NewsNum_i=0&Itemon="& Itemon &"&Itemok=yes&Collecdate="& Collecdate &"';</script>"
	  End if
	  Response.end
End if
If FoundErr<>True then
   ListCode=Skcj.ReplaceTrim(Skcj.GetHttpPage(ListUrl,selEncoding))
   Call GetListPaing()
   If ListCode="$False$" Then
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>在获取列表：" & ListUrl & "网页源码时发生错误！</li>"
   Else
      ListCode=Skcj.GetBody(ListCode,LsString,LoString,False,False)
      If ListCode="$False$" Or ListCode="" Then
         FoundErr=True
		 FoundErr_1=True
         ErrMsg=ErrMsg & "<br><li>在截取：" & ListUrl & "的"& ErrMsg_lx &"列表时发生错误！</li>"
      End If
   End If
End If

If FoundErr<>True Then
   NewsArrayCode=Skcj.GetArray(ListCode,HsString,HoString,False,False)
   If NewsArrayCode="$False$" Then
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>在分析：" & ListUrl & ""& Skcj.GetItemConfig("CjName",Colleclx) &"列表时发生错误！</li>"
   Else
      NewsArray=Split(NewsArrayCode,"$Array$")
      For Arr_i=0 to Ubound(NewsArray)
         If HttpUrlType=1 Then
            NewsArray(Arr_i)=Trim(Replace(HttpUrlStr,"{$ID}",NewsArray(Arr_i)))
         Else
            NewsArray(Arr_i)=Trim(FormatRemoteUrl(NewsArray(Arr_i),ListUrl))    
         End If
         NewsArray(Arr_i)=CheckUrl(NewsArray(Arr_i))
      Next 
      If CollecOrder=1 Then
         For Arr_i=0 to Fix(Ubound(NewsArray)/2)
            OrderTemp=NewsArray(Arr_i)
            NewsArray(Arr_i)=NewsArray(Ubound(NewsArray)-Arr_i)
            NewsArray(Ubound(NewsArray)-Arr_i)=OrderTemp
         Next
      End If
   End If
End If


If FoundErr<>True Then
	If x_tp=1 then
	   NewsimageCode=Skcj.GetArray(ListCode,imhstr,imostr,False,False)
	   If NewsimageCode="$False$" Then
		  FoundErr=True
		  ErrMsg=ErrMsg & "<br><li>在分析：" & ListUrl & "小图片列表时发生错误！</li>"
	   Else
		  Newsimage=Split(NewsimageCode,"$Array$")
		  For Arr_i=0 to Ubound(Newsimage)
			 If HttpUrlType=1 Then
				Newsimage(Arr_i)=Trim(Replace(HttpUrlStr,"{$ID}",Newsimage(Arr_i)))
			 Else
				Newsimage(Arr_i)=Trim(Skcj.FormatRemoteUrl(Newsimage(Arr_i),ListUrl))           
			 End If
			 if x_tpUrl<>"" then Newsimage(Arr_i)= x_tpUrl & Newsimage(Arr_i) 
			 Newsimage(Arr_i)=CheckUrl(Newsimage(Arr_i))
			 
		  Next
		  If CollecOrder=True Then
			 For Arr_i=0 to Fix(Ubound(Newsimage)/2)
				OrderTemp=Newsimage(Arr_i)
				Newsimage(Arr_i)=Newsimage(Ubound(Newsimage)-Arr_i)
				Newsimage(Ubound(Newsimage)-Arr_i)=OrderTemp
			 Next
		  End If
	   End If
	End If
End if
If FoundErr<>True Then
dim PicUrls_i
   Call TopItem2()
   CollecNewsAll=0
	  Arr_i=NewsNum_i
	  If CollecNewsAll>=CollecNewsNum And CollecNewsNum<>0 then Call Foot_Item
      CollecNewsAll=CollecNewsAll+1
      UploadFiles=""
      DefaultPicUrl=""
      IncludePic=0
      ImagesNum=0
      NewsCode=""
      FoundErr=False
      ErrMsg=""
      His_Repeat=False
	  NewsUrl=NewsArray(Arr_i)
      Title=""
      PaingNum=1
      If Response.IsClientConnected Then 
         Response.Flush 
      Else 
         Response.End 
      End If

      If CollecTest=False Then
         His_Repeat=CheckRepeat(NewsUrl)
      Else
         His_Repeat=False
      End If
      If His_Repeat=True Then
         FoundErr=True
      End If
	  If FoundErr<>True then
		  If x_tp=1 then
		  	'On Error Resume Next
			picpath=Newsimage(arr_i)
			iF SaveFiles=1 then picpath=Skcj.Sk_SaveFile(Colleclx,picpath)
		  End if
	  End if
      If FoundErr<>True Then
         NewsCode=Skcj.ReplaceTrim(skcj.GetHttpPage(NewsUrl,selEncoding))
         If NewsCode="$False$" Then
            FoundErr=True
			ErrMsg=ErrMsg & "<br>在获取：" & NewsUrl & "网页源码时发生错误！"
			Title="获取网页源码失败"
         End If
      End If

      If FoundErr<>True Then
         Title=Skcj.GetBody(NewsCode,TsString,ToString,False,False)
         If Title="$False$" or Title="" then
            FoundErr=True
            ErrMsg=ErrMsg & "<br>在分析：" & NewsUrl & "的"& Skcj.GetItemConfig("CjName",Colleclx) &"标题时发生错误"
			Title="<br>标题分析错误" 
         End If
         If FoundErr<>True Then
		 	if CsString<>"0" or CoString<>"0" then
            Content=Skcj.GetBody(NewsCode,CsString,CoString,False,False)
            else
			Content=""
			end if
			If Content="$False$" Then
               'FoundErr=True
               'ErrMsg=ErrMsg & "<br>在分析：" & NewsUrl & "的"& ErrMsg_lx &"正文时发生错误"
               Title=Title & "<br>正文分析错误" 
            End If
         End If
         If FoundErr<>True Then
         If NewsPaingType=1 Then
            NewsPaingNext=Skcj.GetBody(NewsCode,NPsString,NPoString,False,False)
			   If NewsPaingNext<>"$False$" Then
			   		 NewsPaingNext_Code=Skcj.GetArray(NewsPaingNext,NewsUrlPaing_s,NewsUrlPaing_o,False,False)
					 TypeArray_Url=Split( NewsPaingNext_Code,"$Array$")
					 if Ubound(TypeArray_Url)<>0 Then						
						 For i=0 to Ubound(TypeArray_Url)
							Call ShowMsg_1("分页采集中...  当前采集第"&I+1&"页<br>")
							Response.Flush()
							 TypeNews_Url=Trim(Skcj.FormatRemoteUrl(TypeArray_Url(i),NewsUrl))
							 NewsPaingNextCode=Skcj.ReplaceTrim(skcj.GetHttpPage(TypeNews_Url,selEncoding)) 
							 '---------------------------图片分页--------------------------------------------
							 IF Colleclx=2 Then 
								 	PicUrls=Skcj.GetBody(NewsPaingNextCode,photourls,photourlo,False,False)
									PicUrls=Trim(Skcj.FormatRemoteUrl(PicUrls,NewsUrl))
							 		IF SaveFiles=1 then 
										PicUrls=Skcj.Sk_SaveFile(Colleclx,PicUrls)
										If PicUrls=False then
										Response.Write "&nbsp;----" & PicUrls & " 保存失败<br>"
										Else
										Response.Write "&nbsp;" & Skcj.GetItemConfig("CjName",Colleclx) & I &"--" & PicUrls & " 保存成功<br>"
										End if
										Response.Flush()
									End IF
									if PicUrls<>False then
										If i=0 then
												PicUrls_i="图片1|" & PicUrls 
										Else
												PicUrls_i= PicUrls_i & "|||图片" & i  & "|" & PicUrls 
										End if
									End if	
									PicUrls=PicUrls_i
							 End if
							 '---------------------------图片分页------------------------------------------------
							 ContentTemp=Skcj.GetBody(NewsPaingNextCode,CsString,CoString,False,False)
							 If ContentTemp<>"$False$" then
							 	If InStr(Content,ContentTemp)<=0 Then
									Content=Content & "[NextPage]" & ContentTemp
								End If
							End If
						 Next
					 End if
               End If
         End If
		 '过滤
		 Call FilterScript()
		 Call GetFilters
         Call Filters
         Title=FpHtmlEnCode(Title)
         Content=Ubbcode(Content)
		 Content=Skcj.ItemReplaceStr(Content,strReplace)'内容替换
         End If
      End If
	  
If Colleclx=2 And FoundErr<>True then '图片下载
	'--------------------------------深度3采集-------------------------------------	
	IF NewsPaingType=2 Then
			i=1
			ListTypeCode=Skcj.GetBody(NewsCode,PhotoType_s,PhotoType_o,False,False)
			If ListTypeCode<>"$False$" Then 
				ListTypeUrlCode=Skcj.GetArray(ListTypeCode,PhotoLurl_s,PhotoLurl_o,False,False)
				If Phototypefy_s<>"0" AND Phototypefy_o<>"0" AND Phototypefyurl_s<>"0" AND Phototypefyurl_o<>"0" Then
				ListTypeCode_2=Skcj.GetBody(NewsCode,Phototypefy_s,Phototypefy_o,False,False)
					If ListTypeCode_2<>"$False$"   Then 
						ListTypeUrlCode_2=Skcj.GetArray(ListTypeCode_2,Phototypefyurl_s,Phototypefyurl_o,False,False)
						TypeUrlArray_2=Split(ListTypeUrlCode_2,"$Array$")
						For Arr_ii_2=0 to Ubound(TypeUrlArray_2)
						TypeNewsUrl=Trim(Skcj.FormatRemoteUrl(TypeUrlArray_2(Arr_ii_2),NewsUrl))
							NewsTypeCode=Skcj.ReplaceTrim(skcj.GetHttpPage(TypeNewsUrl,selEncoding))
							ListTypeCode=Skcj.GetBody(NewsTypeCode,PhotoType_s,PhotoType_o,False,False)
							If ListTypeCode<>"$False$" Then 
								ListTypeUrlCode=Skcj.GetArray(ListTypeCode,PhotoLurl_s,PhotoLurl_o,False,False)
								TypeUrlArray=Split(ListTypeUrlCode,"$Array$")
								For Arr_ii=0 to Ubound(TypeUrlArray)
									TypeNewsUrl=Trim(Skcj.FormatRemoteUrl(TypeUrlArray(Arr_ii),NewsUrl))
									If TypeNewsUrl<>"$False$" Then 
										
										if Phototypeurl_s<>"0" or Phototypeurl_o<>"0" then
											NewsTypeCode=Skcj.ReplaceTrim(skcj.GetHttpPage(TypeNewsUrl,selEncoding))
											PicUrls=Skcj.GetBody(NewsTypeCode,Phototypeurl_s,Phototypeurl_o,False,False)
											PicUrls=Trim(Skcj.FormatRemoteUrl(PicUrls,TypeNewsUrl))
											if HttpUrlStr<>"" then PicUrls=HttpUrlStr & PicUrls'重定地址
										Else
											PicUrls=TypeNewsUrl
										end if
										IF SaveFiles=1 then 
											PicUrls=Skcj.Sk_SaveFile(Colleclx,PicUrls)
											if PicUrls=False then
											Response.Write "&nbsp;----" & PicUrls & " 保存失败<br>"
											Else
											Response.Write "&nbsp;" & Skcj.GetItemConfig("CjName",Colleclx) & I &"--" & PicUrls & " 保存成功<br>"
											End if
											Response.Flush()
										End IF
										if PicUrls<>False then
											If arr_ii=0 and Arr_ii_2=0 then
													PicUrls_i="图片1|" & PicUrls 
													i=i+1
											Else
													PicUrls_i= PicUrls_i & "|||图片" & i  & "|" & PicUrls 
													i=i+1
											End if
											PicUrls=PicUrls_i
										End if
									End If
								Next
							End If
						Next
						PicUrls=PicUrls_i
						Call SaveArticle
					Else
						Call Coll_ListType_2	
					End if
				Else
					Call Coll_ListType_2	
				End If
			Else
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>在获取:" & NewsUrl & "2级分类列表源码时发生错误。</li>"	
			End If
	End if
	'-----------------------------深度3采集----------------------------------------
	If NewsPaingType=0 Then
		If Downlist_s="" or  Downlist_o="" or DownUrl_s="" or DownUrl_o="" then'图片下载
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>图片地址设置不能为空</li>" 
		Else	
			DownUrls=Skcj.GetBody(NewsCode,Downlist_s,Downlist_o,False,False)
			If DownUrls<>"$False$" then
				DownUrls=Skcj.GetBody(DownUrls,DownUrl_s,DownUrl_o,False,False)
				IF DownUrls<>"$False$" then	
					DownUrls=Trim(Skcj.FormatRemoteUrl(DownUrls,NewsUrl))
					IF SaveFiles=1 then
						DownUrls=Skcj.Sk_SaveFile(Colleclx,DownUrls)
						if DownUrls=False then
							Response.Write "&nbsp;----" & DownUrls & " 保存失败<br>"
						Else
							Response.Write "&nbsp;图片" & DownUrls & " 保存成功<br>"
						End if
						Response.Flush()
					End IF
					PicUrls=DownUrls
					PicUrls= "图片1|" & PicUrls
				Else
					FoundErr=True
					ErrMsg=ErrMsg & "<br><li>在获取:" & NewsUrl & "图片链接源码时发生错误。</li>"
				End if
			Else
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>在获取:" & NewsUrl & "图片列表源码时发生错误。</li>"
			End if
		End if
	End if
End If
	  	  
If Colleclx=3 And FoundErr<>True then '下载
dim DownUrls_i
  	If Downlist_s="" or  Downlist_o="" or DownUrl_s="" or DownUrl_o="" then'下载地址设置
    	FoundErr=True
    	ErrMsg=ErrMsg & "<br><li>下载地址设置不能为空</li>" 
	Else	
		DownUrls=Skcj.GetBody(NewsCode,Downlist_s,Downlist_o,False,False)
		If DownUrls<>"$False$" then
			IF LinkUrlYn=1 then
				DownUrls=Skcj.GetArray(DownUrls,DownUrl_s,DownUrl_o,False,False)
			else
				DownUrls=Skcj.GetBody(DownUrls,DownUrl_s,DownUrl_o,False,False)
			end if
			IF DownUrls<>"$False$" then
					if LinkUrlYn=1 then
						i=1	
						TypeUrlArray=Split(DownUrls,"$Array$")
						For Arr_ii=0 to Ubound(TypeUrlArray)
							DownUrls=Trim(Skcj.FormatRemoteUrl(TypeUrlArray(Arr_ii),NewsUrl))
							If arr_ii=0  then
								DownUrls_i="下载地址1|" & DownUrls
								i=i+1
							Else
								DownUrls_i= DownUrls_i & "|||下载地址" & i  & "|" & DownUrls
								i=i+1
							End if
						Next
						DownUrls=DownUrls_i
					Else
						DownUrls=Trim(Skcj.FormatRemoteUrl(DownUrls,NewsUrl))
						DownUrls="下载地址1|" & DownUrls
					End if	 
			Else
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>在获取:" & NewsUrl & "下载地址链接源码时发生错误。</li>"
			End if
		Else
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>在获取:" & NewsUrl & "下载地址列表源码时发生错误。</li>"
		End if
	End if
	
	If DownNewType=1 then'新窗口打开下载连接
		If DownNewlist_s<>"" or  DownNewlist_o<>"" or DownNewUrl_s<>"" or DownNewUrl_o<>"" then
			DownUrls=Replace(DownUrls,"下载地址1|","")
			DownUrls=Skcj.ReplaceTrim(skcj.GetHttpPage(DownUrls,selEncoding))
			DownUrls=Skcj.GetBody(DownUrls,DownNewlist_s,DownNewlist_o,False,False)
			If DownUrls<>"$False$" then 
				DownUrls=Skcj.GetArray(DownUrls,DownNewUrl_s,DownNewUrl_o,False,False)
				IF DownUrls<>"$False$" then
					i=1	
					TypeUrlArray=Split(DownUrls,"$Array$")
					For Arr_ii=0 to Ubound(TypeUrlArray)
						DownUrls=Trim(Skcj.FormatRemoteUrl(TypeUrlArray(Arr_ii),NewsUrl))
						If arr_ii=0  then
							DownUrls_i="下载地址1|" & DownUrls
							i=i+1
						Else
							DownUrls_i= DownUrls_i & "|||下载地址" & i  & "|" & DownUrls
							i=i+1
						End if
					Next
					DownUrls=DownUrls_i
				Else
					'FoundErr=True
    				ErrMsg=ErrMsg & "<br><li>在获取:" & NewsUrl & "(新窗口)下载地址列表源码时发生错误</li>" 
				End if
			Else
				'FoundErr=True
    			ErrMsg=ErrMsg & "<br><li>在获取:" & NewsUrl & "(新窗口)下载地址列表源码时发生错误</li>" 
			End if
		Else
			FoundErr=True
    		ErrMsg=ErrMsg & "<br><li>下载地址设置不能为空</li>" 
		End if
	End if 
	
	If ZdType_001=0 then'软件大小设置
   		DownSize=""
	Else
		If  Zds_001="0" and Zdo_001<>"" then
			DownSize=Zdo_001
		Else
			DownSize=FpHtmlEncode(Skcj.GetBody(NewsCode,Zds_001,Zdo_001,False,False))
		End If
	End If
	If ZdType_002=0 then'软件语言设置
   		DownYY=""
	Else
		If  Zds_002="0" and Zdo_002<>"" then
			DownYY=Zdo_002
		Else
			DownYY=FpHtmlEncode(Skcj.GetBody(NewsCode,Zds_002,Zdo_002,False,False))
		End If
	End If
 	If ZdType_003=0 then'授权方式设置
   		DownSQ=""
	Else
		If  Zds_003="0" and Zdo_003<>"" then
			DownSQ=Zdo_003
		Else
			DownSQ=FpHtmlEncode(Skcj.GetBody(NewsCode,Zds_003,Zdo_003,False,False))
		End If
	End If
 	If ZdType_004=0 then'运行环境设置
   		DownPT=""
	Else
		If  Zds_004="0" and Zdo_004<>"" then
			DownPT=Zdo_004
		Else
			DownPT=FpHtmlEncode(Skcj.GetBody(NewsCode,Zds_004,Zdo_004,False,False))
		End If
	End If
 	If ZdType_005=0 then'演示地址设置
   		YSDZ=""
	Else
		If  Zds_005="0" and Zdo_005<>"" then
			YSDZ=Zdo_005
		Else
			YSDZ=FpHtmlEncode(Skcj.GetBody(NewsCode,Zds_005,Zdo_005,False,False))
		End If
	End If
 	If ZdType_006=0 then'注册地址设置
   		ZCDZ=""
	Else
		If  Zds_006="0" and Zdo_006<>"" then
			ZCDZ=Zdo_006
		Else
			ZCDZ=FpHtmlEncode(Skcj.GetBody(NewsCode,Zds_006,Zdo_006,False,False))
		End If
	End If	
	
 	If ZdType_007=0 then'软件图片设置
   		PhotoUrl=""
	Else
		If  Zds_007="0" and Zdo_007<>"" then
			PhotoUrl=Zdo_007
		Else
			PhotoUrl=Skcj.GetBody(NewsCode,Zds_007,Zdo_007,False,False)
			PhotoUrl=Trim(Skcj.FormatRemoteUrl(PhotoUrl,NewsUrl))
		End If
	End If  
End if

IF Colleclx=4 And FoundErr<>True then '动漫
	If Downlist_s="" or  Downlist_o="" or DownUrl_s="" or DownUrl_o="" then'下载地址设置
    	FoundErr=True
    	ErrMsg=ErrMsg & "<br><li>动漫下载地址设置不能为空</li>" 
	Else	
		DownUrls=Skcj.GetBody(NewsCode,Downlist_s,Downlist_o,False,False)
		If DownUrls<>"$False$" then
			DownUrls=Skcj.GetBody(DownUrls,DownUrl_s,DownUrl_o,False,False)
			IF DownUrls<>"$False$" then
						DownUrls=Trim(Skcj.FormatRemoteUrl(DownUrls,NewsUrl))
						IF SaveFiles=1 then 
							DownUrls=Skcj.Sk_SaveFile(Colleclx,DownUrls)
						End IF 
			Else
				'FoundErr=True
				ErrMsg=ErrMsg & "<br><li>在获取:" & NewsUrl & "下载地址链接源码时发生错误。</li>"
			End if
		Else
			'FoundErr=True
			ErrMsg=ErrMsg & "<br><li>在获取:" & NewsUrl & "下载地址列表源码时发生错误。</li>"
		End if
	End if
	
	If DownNewType=1 then'新窗口打开下载连接
		If DownNewlist_s<>"" or  DownNewlist_o<>"" or DownNewUrl_s<>"" or DownNewUrl_o<>"" then
			DownUrls=Skcj.ReplaceTrim(skcj.GetHttpPage(DownUrls,selEncoding))
			DownUrls=Skcj.GetBody(DownUrls,DownNewlist_s,DownNewlist_o,False,False)
			If DownUrls<>"$False$" then 
				DownUrls=Skcj.GetBody(DownUrls,DownNewUrl_s,DownNewUrl_o,False,False)
				IF DownUrls<>"$False$" then
					DownUrls=Trim(Skcj.FormatRemoteUrl(DownUrls,NewsUrl))
				Else
					'FoundErr=True
    				ErrMsg=ErrMsg & "<br><li>在获取:" & NewsUrl & "(新窗口)下载地址列表源码时发生错误</li>" 
				End if
			Else
				'FoundErr=True
    			ErrMsg=ErrMsg & "<br><li>在获取:" & NewsUrl & "(新窗口)下载地址列表源码时发生错误</li>" 
			End if
		Else
			FoundErr=True
    		ErrMsg=ErrMsg & "<br><li>下载地址设置不能为空</li>" 
		End if
	End if 
End if 	  

      If FoundErr<>True Then
         If DateType=0 Then
            UpDateTime=Now()
         Else
		 	If DateType=1 Then
               UpDateTime=Skcj.GetBody(NewsCode,DsString,DoString,False,False)
               UpDateTime=FpHtmlEncode(UpDateTime)
               UpDateTime=Trim(Replace(UpDateTime,"&nbsp;"," "))
               If IsDate(UpDateTime)=True Then
                  UpDateTime=CDate(UpDateTime)
               Else
                  UpDateTime=Now()
               End If
            End If
         End If
                  
         If AuthorType=1 Then
            Author=Skcj.GetBody(NewsCode,AsString,AoString,False,False)
         ElseIf AuthorType=2 Then
            Author=AuthorStr
         Else
            Author="佚名"
         End If
         Author=FpHtmlEncode(Author)
         If Author="" or Author="$False$" then
            Author="佚名"
         Else
            If Len(Author)>255 then
               Author=Left(Author,255)
            End If
         End If
           
         If CopyFromType=1 Then
            CopyFrom=Skcj.GetBody(NewsCode,FsString,FoString,False,False)
         ElseIf CopyFromType=2 Then
            CopyFrom=CopyFromStr
         Else
            CopyFrom="不详"
         End If
         CopyFrom=FpHtmlEncode(CopyFrom)
         If CopyFrom="" or CopyFrom="$False$" Then
	            CopyFrom="不详"
         Else
            If Len(CopyFrom)>255 Then 
               CopyFrom=Left(CopyFrom,255)
            End If
         End If
         If KeyType=0 Then
            Key=Title
            Key=CreateKeyWord(Key,2)
         ElseIf KeyType=1 Then
            Key=Skcj.GetBody(NewsCode,KsString,KoString,False,False)
			Key=Replace(Key,",","|")
			Key=Replace(Key," ","|")
            Key=FpHtmlEncode(Key)
            'Key=CreateKeyWord(Key,2)
         ElseIf KeyType=2 Then
            Key=KeyStr
            Key=FpHtmlEncode(Key)
            If Len(Key)>253 Then
               Key= Left(Key,253)
            Else
               Key=Key
            End If
         End If
         If Key="" or Key="$False$" Then
            Key=KeyStr
         End If
		 
         If UploadFiles<>"" Then
            If Instr(UploadFiles,"|")>0 Then
               Arr_Images=Split(UploadFiles,"|") 
               ImagesNum=Ubound(Arr_Images)+1
               DefaultPicUrl=Arr_Images(0)
            Else
               ImagesNum=1
               DefaultPicUrl=UploadFiles
            End If

            If DefaultPicYn=False then
               DefaultPicUrl=""
            End If
            If IncludePicYn=True Then
               IncludePic=-1
            Else
               IncludePic=0
            End If
            If SaveFiles<>True Then
               UploadFiles=""
            End If
         Else
            ImagesNum=0
            DefaultPicUrl=""
            IncludePic=0         
         End If
         ImagesNumAll=ImagesNumAll+ImagesNum
      End If

	  if  Colleclx =1 then 
	  	set rs = ConnItem.execute("select top 1 Dir from SK_cj where ID=" & Colleclx)
			IF SaveFiles=1 then
				Content=Skcj.ReplaceSaveRemoteFile(Content,"/",rs("Dir") & SaveFileUrl,True,NewsUrl)'远程图片
				Content=SKcj.ReSaveRemoteFile(Content,NewsUrl,rs("Dir") & SaveFileUrl,True)'远程文件
			Else
				Content=Skcj.ReplaceSaveRemoteFile(Content,"/",rs("Dir"),False,NewsUrl)'远程图片
				Content=Skcj.ReSaveRemoteFile(Content,NewsUrl,rs("Dir"),False)'远程文件
			End if 
		rs.close
		set rs=nothing
	  End if
	  '--
      If FoundErr<>True Then
			If His_Repeat<>True Then
				Call SaveArticle
			End if		
	  	If CollecTest=False Then
            SqlItem="INSERT INTO Histroly(ItemID,ChannelID,ClassID,SpecialID,ArticleID,Title,CollecDate,NewsUrl,Result) VALUES ('" & ItemID & "','" & ChannelID & "','" & ClassID & "','" & SpecialID & "','" & ArticleID & "','" & Title & "','" & Now() & "','" & NewsUrl & "',True)"
            ConnItem.Execute(SqlItem)
            Content=Replace(Content,"[InstallDir_ChannelDir]",strInstallDir & strChannelDir & "/")
         End If
         NewsSuccesNum=NewsSuccesNum+1
         ErrMsg=ErrMsg & "No:<font color=red>" & NewsSuccesNum+NewsFalseNum & "</font><br>"
		 ErrMsg=ErrMsg & Skcj.GetItemConfig("CjName",Colleclx) &"标题："
         ErrMsg=ErrMsg & "<font color=red>" & Title & "</font><br>"
         ErrMsg=ErrMsg & "更新时间：" & UpDateTime & "<br>"
		 If Colleclx=1 then ErrMsg=ErrMsg & "新闻标题：" : ErrMsg=ErrMsg & "新闻作者：" & Author & "<br>" : ErrMsg=ErrMsg & "新闻来源：" & CopyFrom & "<br>"
         ErrMsg=ErrMsg & "采集页面：<a href=" & NewsUrl & " target=_blank>" & NewsUrl & "</a><br>"
		 if x_tp =1 then ErrMsg=ErrMsg & "采集小图片：<a href=" & picpath & " target=_blank>" & picpath & "</a><br>"
		 
		 ErrMsg=ErrMsg & "正文预览："
         If Content_View=True Then
            ErrMsg=ErrMsg & "<br>" & Content
         Else
            ErrMsg=ErrMsg & "您没有启用正文预览功能"
         End If
         ErrMsg=ErrMsg & "<br><br>关 键 字：" & Key & ""
      Else
         NewsFalseNum=NewsFalseNum+1
         If His_Repeat=True Then
            ErrMsg=ErrMsg & "No:<font color=red>" & NewsSuccesNum+NewsFalseNum & "</font><br>"
			ErrMsg=ErrMsg & "目标"& Skcj.GetItemConfig("CjName",Colleclx) &"：<font color=red>"
            If His_Result=True Then
               ErrMsg=ErrMsg & His_Title
            Else
               ErrMsg=ErrMsg & NewsUrl
            End If
            ErrMsg=ErrMsg & "</font> 的记录已存在，不给予采集。<br>"
            ErrMsg=ErrMsg & "采集时间：" & His_CollecDate & "<br>"
			ErrMsg=ErrMsg & ""& Skcj.GetItemConfig("CjName",Colleclx) &"来源：<a href='" & NewsUrl & "' target=_blank>"&NewsUrl&"</a><br>"
            ErrMsg=ErrMsg & "采集结果："
            If His_Result=False Then
               ErrMsg=ErrMsg & "失败"
               ErrMsg=ErrMsg & "<br>失败原因：" & Title
            Else
               ErrMsg=ErrMsg & "成功"
            End If            
            ErrMsg=ErrMsg & "<br>提示信息：如想再次采集，请先将该新闻的历史记录<font color=red>删除</font><br>"
         End If
         If CollecTest=False And His_Repeat=False Then
            SqlItem="INSERT INTO Histroly(ItemID,ChannelID,ClassID,SpecialID,Title,CollecDate,NewsUrl,Result) VALUES ('" & ItemID & "','" & ChannelID & "','" & ClassID & "','" & SpecialID & "','" & Title & "','" & Now() & "','" & NewsUrl & "',False)"
            ConnItem.Execute(SqlItem)
         End If
	  	 
      End If
	   Response.Write "<table width=""100%"" border=""0"" align=""center"" cellpadding=""2"" class=""tableBorder""  cellspacing=""1"">"
	   Response.Write "   <tr class='tdbg'>"          
	   Response.Write "      <td height=""22"" colspan=""2"" align=""left"">"
	   Response.Write ErrMsg
	   Response.Write "      </td>"
	   Response.Write "   </tr><br>"
	   Response.Write "</table>"
	    
	   if NewsNum_i = Ubound(NewsArray) then
	   		if Itemon<>"" then Itemok= "yes"'全选采集
			if Collecdate<>"" Then 
				Collecdate=Day(now())
				response.write("<script>location.href='sk_Timing.asp?action=GoTiming&Collecdate="&  Day(now()) &"';</script>")'到页面
	    	Else
				response.write("<script>location.href='Sk_Collection.asp?ItemID="& ItemID &"&ItemNum=" & ItemNum & "&ListNum=" & ListNum +1 & "&NewsSuccesNum=" & NewsSuccesNum & "&NewsFalseNum=" & NewsFalseNum & "&ImagesNumAll=" & ImagesNumAll & "&ListPaingNext=" & ListPaingNext &"&NewsNum_i="& 0 &"&Itemok="& Itemok &"&Itemon="& Itemon &"&Collecdate="& Collecdate &"';</script>")'完成
			End if
	   Else
	   response.write "<meta http-equiv=""refresh"" content=""0;url=Sk_Collection.asp?ItemID="& ItemID &"&ItemNum=" & ItemNum & "&ListNum=" & ListNum & "&NewsSuccesNum=" & NewsSuccesNum & "&NewsFalseNum=" & NewsFalseNum & "&ImagesNumAll=" & ImagesNumAll & "&ListPaingNext=" & ListPaingNext &"&NewsNum_i="& NewsNum_i + 1 &"&Itemon="& Itemon &"&Collecdate="& Collecdate &""">"
	   End if
Else
   If FoundErr_1=True Then
		response.write("<script>location.href='Sk_Collection.asp?ItemID="& ItemID &"&ItemNum=" & ItemNum & "&ListNum=" & ListNum +1 & "&NewsSuccesNum=" & NewsSuccesNum & "&NewsFalseNum=" & NewsFalseNum & "&ImagesNumAll=" & ImagesNumAll & "&ListPaingNext=" & ListPaingNext &"&NewsNum_i="& 0 &"&Itemok="& Itemok &"&Itemon="& Itemon &"&Collecdate="& Collecdate &"';</script>")'完成
	 	FoundErr_1=False
	End If
   Call ShowMsg(ErrMsg)
End If
End Sub
Sub  Coll_ListType_2()
i=1
	If ListTypeUrlCode<>"$False$" and Instr(NewsArrayCode,"$Array$")>0  Then 
			TypeUrlArray=Split(ListTypeUrlCode,"$Array$")
			For Arr_ii=0 to Ubound(TypeUrlArray)
				TypeNewsUrl=Trim(Skcj.FormatRemoteUrl(TypeUrlArray(Arr_ii),NewsUrl))
				If TypeNewsUrl<>"$False$" Then
					if Phototypeurl_s<>"0" or Phototypeurl_o<>"0" then
						NewsTypeCode=Skcj.ReplaceTrim(skcj.GetHttpPage(TypeNewsUrl,selEncoding))
						PicUrls=Skcj.GetBody(NewsTypeCode,Phototypeurl_s,Phototypeurl_o,False,False)
						PicUrls=Trim(Skcj.FormatRemoteUrl(PicUrls,TypeNewsUrl))
						if HttpUrlStr<>"" then PicUrls=HttpUrlStr & PicUrls'重定地址
					else
						PicUrls=TypeNewsUrl
					end if
					IF SaveFiles=1 then 
                    	PicUrls=Skcj.Sk_SaveFile(Colleclx,PicUrls)
						If PicUrls=False then
                    		Response.Write "&nbsp;----" & PicUrls & " 保存失败<br>"
						Else 
							Response.Write "&nbsp;" & Skcj.GetItemConfig("CjName",Colleclx) & i &"--" & PicUrls & " 保存成功<br>"
						End if
						Response.Flush()
                    End IF
					If PicUrls<>False then
						If i=1 then 
							PicUrls_i= "图片1|" & PicUrls
							i=i+1
						Else
							PicUrls_i= PicUrls_i & "|||图片"  & i & "|" & PicUrls
							i=i+1 
						End if
					End if
				End If
			Next
			PicUrls=PicUrls_i
			Call SaveArticle
	Else
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>在获取:" & NewsUrl & "2级分类链接源码时发生错误。</li>"	
	End If
End Sub
Sub ShowMsg(Msg)
   Dim strTemp
   strTemp= "<table width=""100%"" border=""0"" align=""center"" cellpadding=""2"" class=""tableBorder""  cellspacing=""1"">"
   strTemp=strTemp & "   <tr class='tdbg'>"          
   strTemp=strTemp & "      <td height=""22"" colspan=""2"" align=""left"">"
   strTemp=strTemp & Msg
   strTemp=strTemp & "      </td>"
   strTemp=strTemp & "   </tr><br>"
   strTemp=strTemp & "</table>"
   Response.Write StrTemp     
End Sub
Sub ShowMsg_1(Msg)
   Dim strTemp
   strTemp= "<table width=""100%"" border=""0"" align=""center"" cellpadding=""2""   cellspacing=""1"">"
   strTemp=strTemp & "   <tr class='tdbg'>"          
   strTemp=strTemp & "      <td height=""22"" colspan=""2"" align=""left"">"
   strTemp=strTemp & Msg
   strTemp=strTemp & "      </td>"
   strTemp=strTemp & "   </tr>"
   strTemp=strTemp & "</table>"
   Response.Write StrTemp     
End Sub
Sub GetListPaing()
   If ListPaingType=1 Then
      ListPaingNext=GetPaing(ListCode,LPsString,LPoString,False,False)
      ListPaingNext=FpHtmlEnCode(ListPaingNext)
      If ListPaingNext<>"$False$" And ListPaingNext<>"" Then
         If ListPaingStr1<>""  Then  
            ListPaingNext=Replace(ListPaingStr1,"{$ID}",ListPaingNext)
         Else
            ListPaingNext=Skcj.FormatRemoteUrl(ListPaingNext,ListUrl)
         End If
         ListPaingNext=Replace(ListPaingNext,"&","{$ID}")
      End If
   Else
      ListPaingNext="$False$"
   End If
End Sub
Sub Foot_Item()
Response.Write "<table width=""100%"" border=""0"" align=""center"" cellpadding=""2"" class=""tableBorder""  cellspacing=""1"">"
Response.Write "<tr>"
Response.write "<td height=""22"" colspan=""2"" align=""left"" class=""tdbg"">"
If CollecTest=False Then
   Response.Write "数据整理中，5秒后继续......5秒后如果还没反应请点击 <a href='Sk_Collection.asp?ItemID="& ItemID &"&ItemNum=" & ItemNum & "&ListNum=" & ListNum+1 & "&NewsSuccesNum=" & NewsSuccesNum & "&NewsFalseNum=" & NewsFalseNum & "&ImagesNumAll=" & ImagesNumAll & "&ListPaingNext=" & ListPaingNext & "'><font color=red>这里</font></a> 继续<br>"
   Response.Write "<meta http-equiv=""refresh"" content=""5;url=Sk_Collection.asp?ItemID="& ItemID &"&ItemNum=" & ItemNum & "&ListNum=" & ListNum+1 & "&NewsSuccesNum=" & NewsSuccesNum & "&NewsFalseNum=" & NewsFalseNum & "&ImagesNumAll=" & ImagesNumAll & "&ListPaingNext=" & ListPaingNext  & """>"
Else
   Response.Write "<a href='Sk_Collection.asp?ItemID="& ItemID &"&ItemNum=" & ItemNum & "&ListNum=" & ListNum+1 & "&NewsSuccesNum=" & NewsSuccesNum & "&NewsFalseNum=" & NewsFalseNum & "&ImagesNumAll=" & ImagesNumAll & "&ListPaingNext=" & ListPaingNext & "'><font color=red>请 继 续</font></a>"
End If
Response.Write "</td></tr>"
Response.Write "</table>"
end Sub
Sub FootItem2()
   Dim strTemp
   OverTime=Timer()
   strTemp= "<table width=""100%"" border=""0"" align=""center"" cellpadding=""2"" cellspacing=""1"">"       
   strTemp=strTemp & "<tr>"          
   strTemp=strTemp & "<td height=""22"" colspan=""2"" align=""left"" class=""tdbg"">"
   strTemp=strTemp & "执行时间：" & CStr(FormatNumber((OverTime-StartTime)*1000,2)) & " 毫秒"
   strTemp=strTemp & "</td></tr><br>"
   strTemp=strTemp & "</table>"
   Response.write strTemp
End Sub
Sub SetCache_His()
   SqlItem ="select NewsUrl,Title,CollecDate,Result from Histroly"
   Set RsItem=Server.CreateObject("adodb.recordset")
   RsItem.Open SqlItem,ConnItem,1,1
   If Not RsItem.Eof Then
      Arr_Histrolys=RsItem.GetRows()
   End If
   RsItem.Close
   Set RsItem=Nothing

   Dim myCache
   Set myCache=new SK_clsCache
   myCache.name=CacheTemp & "histrolys"
   Call myCache.clean()
   If IsArray(Arr_Histrolys)=True Then   
      myCache.add Arr_Histrolys,Dateadd("n",1000,now)
   End If
End Sub
Public Sub GetFilters()
dim sqlf,RSF
   SqlF ="Select * from Filters Where Flag=1 And (PublicTf=True Or ItemID=" & ItemID & ") order by FilterID ASC"
   Set RSF=connItem.Execute(SqlF)
   If RsF.Eof And RsF.Bof Then
      Arr_Filters=""
   Else
      Arr_Filters=RsF.GetRows()
   End If
   RsF.Close
   Set RsF=Nothing
End Sub
Sub DelCache()
   Dim myCache

   Set myCache=new SK_clsCache
   myCache.name=CacheTemp & "items"
   Call myCache.clean()
   myCache.name=CacheTemp & "filters"
   Call myCache.clean()
   myCache.name=CacheTemp & "histrolys"
   Call myCache.clean()
   myCache.name=CacheTemp & "collectest"
   Call myCache.clean()
   myCache.name=CacheTemp & "contentview"
   Call myCache.clean()
   Set myCache=Nothing
End Sub

Sub SaveArticle
    If Passed=1 then
		Call SaveDate
		'Call SaveDate_Passed
	Else
		Call SaveDate
	End if
   CollecTest=False
End Sub

Sub SaveDate_Passed'保存数据
	dim rslb,DownID,CMSRS,UploadFiles_1
	Select Case Colleclx
	Case 1'--文章数据库
		Set rslb=Conn.Execute("select top 1 * from ks_Class Where ID='" & ClassID &"'") 
		if not rslb.eof then
		   Set CMSRS=Server.CreateObject("ADODB.RECORDSET")
		   CMSRS.Open "Select top 1 * From KS_Article Where Title='" & Title & "' And Tid='" & ClassID & "'",Conn,1,3
		   If CMSRS.EOF Then
			   CMSRS.addnew
			   CMSRS("NewsID")=KSCMS.GetInfoID(1)
			   CMSRS("Tid")=ClassID'目标ID	  
			   CMSRS("Keywords")=Key
			   CMSRS("Title")=Title
			   if IncludePic=1 then 
					CMSRS("PicNews")=1
					If x_tp=1 then
						CMSRS("picurl")=picpath
					Else
						UploadFiles_1=Split(UploadFiles,"|")
						If Ubound(UploadFiles_1) >= 0 then 
							CMSRS("picurl")=UploadFiles_1(0)
						Else
							CMSRS("picurl")=UploadFiles
						End if
					End if
			   Else
					CMSRS("PicNews")=0
			   End if
			   CMSRS("ShowComment")=0
			   CMSRS("ArticleContent")=Content 
			   CMSRS("Author")=Author
			   CMSRS("Origin")=CopyFrom
			   CMSRS("Rank")=GetStars(Stars)
			   CMSRS("Hits")=Hits '点击数
			   CMSRS("AddDate")=UpDateTime
			   CMSRS("TemplateID")=rslb("ArticleTemplateID")
			   CMSRS("ArticleFsoType")=rslb("ArticleFsoType")
			   CMSRS("Fname")=KSCMS.GetFileName(rslb("ArticleFsoType"),now,rslb("ArticleFnameType"))
			   CMSRS("RefreshTF")=0
			   CMSRS("ArticleInput") = Inputer
			   CMSRS("Changes") = 0
			   CMSRS("Recommend") = 0
			   CMSRS("Rolls") = 0
			   CMSRS("strip") = 0
			   CMSRS("Popular") = 0
			   CMSRS("Verific") = 1      '审核与否
			   CMSRS("Slide") = 0
			   CMSRS("BeyondSavePic") = 0
			   CMSRS("Comment") = 0
			   CMSRS("OrderID") = 1
			   CMSRS.update
		   Else
			   FoundErr=True
               ErrMsg=ErrMsg & "<font color=red>相同标题不能入库!</font><br>"
		   End if
		   CMSRS.close
		   Set CMSRS = Nothing
		End if
		rslb.close
		set rslb=nothing
	Case 3
		set rslb=Conn.Execute("select top 1 * from ks_Class Where ID='" & ClassID &"'")
		if not rslb.eof then 
			Set CMSRS=Server.CreateObject("ADODB.RECORDSET")
			CMSRS.Open "Select top 1 * From KS_Photo Where Title='" & Title & "' And Tid='" & ClassID & "'",Conn,1,3
	   		If CMSRS.EOF Then
	   	    	CMSRS.addnew
				CMSRS("PicID")=KSCMS.GetInfoID(2)
				CMSRS("Tid")=ClassID
				CMSRS("KeyWords")=Key
				CMSRS("Title")=Title
				if x_tp =1 then 
					CMSRS("PhotoUrl")=picpath
				else
					TypeUrlArray_2=Split(PicUrls,"|||")
					picpath=Replace(TypeUrlArray_2(0),"图片1|","")
					if x_tp =2 then '自动缩略图
						dim picpath_1,picpath_2,picpath_3 : picpath_1=Split(picpath,"/")
						picpath_2=Sk_GetSaveDir(Colleclx) & "small" & picpath_1(Ubound(picpath_1)) : picpath_3=picpath_2
						IF SaveFiles=1 then call SKThumb.CreateThumbs(picpath,picpath_2) : picpath=picpath_3 
					end if	
				CMSRS("PhotoUrl")=picpath
				end if
				CMSRS("PicUrls")=PicUrls
				CMSRS("PictureContent")=Content
				CMSRS("Author")=Author
				CMSRS("Origin")=CopyFrom
				CMSRS("Rank")=GetStars(Stars)
				CMSRS("Hits")=Hits
				CMSRS("HitsByDay")=MakeRandom(1)
				CMSRS("HitsByWeek")=MakeRandom(2)
				CMSRS("HitsByMonth")=MakeRandom(3)
				CMSRS("AddDate")=UpdateTime
				CMSRS("TemplateID")=rslb("ArticleTemplateID")
				CMSRS("PictureFsoType")=rslb("ArticleFsoType")
				CMSRS("Fname")=KSCMS.GetFileName(rslb("ArticleFsoType"),now,rslb("ArticleFnameType"))
				CMSRS("PictureInput")=Inputer
				CMSRS("RefreshTF")=0
				CMSRS("Recommend")=0
				CMSRS("Rolls")=0
				CMSRS("Strip")=0
				CMSRS("Popular")=0
				CMSRS("Verific")=1
				CMSRS("Comment")=1
				CMSRS("Score")=MakeRandom(3)
				CMSRS("Slide")=0
				CMSRS("BeyondSavePic")=0
				CMSRS("DelTF")=0
				CMSRS("OrderID")=1
				CMSRS.update
	   		Else
				FoundErr=True
                ErrMsg=ErrMsg & "<font color=red>相同标题不能入库!</font><br>"
			End if
			CMSRS.close
	    	Set CMSRS = Nothing
		End if
		rslb.close
		set rslb=nothing
	Case 5
		set rslb=Conn.Execute("select top 1 * from ks_Class Where ID='" & ClassID &"'") 
	    if not rslb.eof then  
			Set CMSRS=Server.CreateObject("ADODB.RECORDSET")
			CMSRS.Open "Select top 1 * From KS_download Where Title='" & Title & "' And Tid='" & ClassID & "'",Conn,1,3
	   	   If CMSRS.EOF Then
				CMSRS.addnew
				CMSRS("DownID")=KSCMS.GetInfoID(3)
				CMSRS("Tid")=ClassID
				CMSRS("KeyWords")=Key
				CMSRS("Title")=Title
				CMSRS("DownVersion")="" '版本信息
				CMSRS("DownLB")="国产软件"  '下载类别
				CMSRS("DownYY")=DownYY     '下载语言
				CMSRS("DownSQ")= DownSQ    '下载授权
				CMSRS("DownPT")= DownPT     '下载平台
				CMSRS("DownSize")=DownSize   '大小
				CMSRS("YSDZ")=YSDZ       '演示地址
				CMSRS("ZCDZ")=ZCDZ     '注册地址
				CMSRS("JYMM")=""       '解压密码
				CMSRS("PhotoUrl")=PhotoUrl
				CMSRS("BigPhoto")=PhotoUrl  '软件大图
				CMSRS("FlagUrl")=0            '0、默认方式 1、镜像服务器方式
				CMSRS("DownUrls")=DownUrls    '下载地址，若是默认方式，则以|||隔开
				CMSRS("DownContent")=Content '下载简介
				CMSRS("Author")=Author     '下载作者
				CMSRS("Origin")=CopyFrom     '下载来源
				CMSRS("Rank")=GetStars(Stars)      '阅读等级
				CMSRS("Hits")=Hits          '总浏览数
				CMSRS("HitsByDay")=MakeRandom(1)  '日浏览数
				CMSRS("HitsByWeek")=MakeRandom(2)  '周浏览数
				CMSRS("HitsByMonth")=MakeRandom(3)  '月浏览数
				CMSRS("AddDate")=UpdateTime       '更新时间
				CMSRS("TemplateID")=rslb("ArticleTemplateID")           '模板ID
				CMSRS("DownFsoType")=rslb("ArticleFsoType")     '
				CMSRS("Fname")=KSCMS.GetFileName(rslb("ArticleFsoType"),now,rslb("ArticleFnameType"))
				CMSRS("DownInput")=Inputer
				CMSRS("RefreshTF")=0
				CMSRS("Recommend")=0
				CMSRS("Popular")=0
				CMSRS("Verific")=1
				CMSRS("Comment")=0
				CMSRS("DelTF")=0
				CMSRS("OrderID")=1
				CMSRS("InfoPurview")=0'查看权限0继承栏目权限,1所有会员,2指定会员组
				CMSRS("ArrGroupID")="" '指定会员组的查看权限
				CMSRS("ReadPoint")=0  '阅读点数
				CMSRS("ChargeType")=0  '重复收费方式
				CMSRS("PitchTime")=24   '重复收费小时数
				CMSRS("ReadTimes")=10    '重复收费查看次数
				CMSRS("DividePercent")=0   
				CMSRS.update
	  	   Else
		  	    FoundErr=True
                ErrMsg=ErrMsg & "<font color=red>相同标题不能入库!</font><br>"
	  	   End if
	   	   CMSRS.close : Set CMSRS = Nothing	
	   end if
	   rslb.close : set rslb=nothing  
	End select  
End Sub
Sub SaveDate
	dim rslb,DownID
	if  Colleclx =1 then'--文章数据库
	   set rs=server.createobject("adodb.recordset")
	   sql="select top 1 * from SK_Article" 
	   rs.open sql,ConnItem,1,3
	   rs.addnew
	   rs("ChannelID")=1
	   rs("ClassID")=ClassID
	   rs("SpecialID")=SpecialID
	   rs("Title")=Title
       rs("Content")=Content
	   rs("Keyword")=Key
	   rs("Hits")=Hits
	   rs("Author")=Author
	   rs("CopyFrom")=CopyFrom
	   rs("IncludePic")=IncludePic
	   rs("Passed")=0
	   rs("OnTop")=OnTop
	   rs("Hot")=Hot
	   rs("Elite")=Elite
	   rs("Stars")=Stars
	   rs("UpdateTime")=UpDateTime
	   rs("PaginationType")=PaginationType
	   rs("MaxCharPerPage")=MaxCharPerPage
	   rs("ReadLevel")=ReadLevel
	   rs("ReadPoint")=ReadPoint
	   rs("SkinID")=SkinID
	   rs("TemplateID")=TemplateID
	   rs("DefaultPicUrl")=DefaultPicUrl
	   rs("UploadFiles")=UploadFiles
	   rs("ShowCommentLink")=ShowCommentLink
	   rs("Inputer")=Inputer
	   if Editor="" then Editor="Admin"
	   rs("Editor")=Editor
	   rs("picpath")=picpath 

	   rs.update
	   rs.close
	   set rs=nothing
   end if
   if  Colleclx =4 then 
	   	set rs=server.createobject("adodb.recordset")
	   	sql="select top 1 * from SK_flash" 
	   	rs.open sql,ConnItem,1,3
	   	rs.addnew
	 	rs("flashID")=GetInfoID(1) : rs("Tid")=ClassID : rs("KeyWords")=Key : rs("Title")=Title
		if x_tp =1 then
			rs("PhotoUrl")=picpath
		end if 
		rs("flashUrl")=DownUrls : rs("flashContent")=Content : rs("Author")=Author : rs("Origin")=CopyFrom 
		rs("Rank")=Stars  
		rs("LastHitsTime")=UpdateTime :rs("Hits")=Hits :rs("HitsByDay")=makeRandom(3) 
		rs("HitsByWeek")=makeRandom(3) : rs("HitsByMonth")=makeRandom(3) : rs("AddDate")=UpdateTime
		'rs("TemplateID")=rslb("ArticleTemplateID")
		'rs("FlashFsoType")=rslb("ArticleFsoType")
		'rs("Fname")=KSCMS.GetFileName(rslb("ArticleFsoType"),now,rslb("ArticleFnameType"))
		rs("flashInput")=Inputer
		rs("RefreshTF")=0
		rs("Recommend")=0
		rs("Rolls")=0
		rs("Strip")=0
		rs("Popular")=0
		rs("Verific")=0
		rs("Comment")=1
		rs("Score")=MakeRandom(3)
		RS("Verific") = 1      '审核与否
		rs("Slide")=0
		rs("BeyondSavePic")=0
		rs("DelTF")=0
		rs("OrderID")=1
	   	rs.update
	   	rs.close
	   	set rs=nothing
   end if
    if  Colleclx =3 then'--软件
	    set rs=server.createobject("adodb.recordset")
	    sql="select top 1 * from Sk_download" 
	    rs.open sql,ConnItem,1,3
	    rs.addnew
		DownID=GetInfoID_CMS(3)
		rs("DownID")=DownID
		rs("Tid")=ClassID
		rs("KeyWords")=Key
		rs("Title")=Title
		'DownSize,DownYY,DownSQ,DownPT,YSDZ,ZCDZ,PhotoUrl,DownUrls
		'rs("DownVersion")= '版本信息
		'rs("DownLB")=      '下载类别
		rs("DownYY")=DownYY     '下载语言
		rs("DownSQ")=DownSQ     '下载授权
		rs("DownPT")=DownPT     '下载平台
		rs("DownSize")=DownSize    '大小
		rs("YSDZ")=YSDZ       '演示地址
		rs("ZCDZ")=ZCDZ      '注册地址
		'rs("JYMM")=       '解压密码
		rs("PhotoUrl")=PhotoUrl
		'rs("BigPhoto")=    '软件大图
		'rs("FlagUrl")=            '0、默认方式 1、镜像服务器方式
		rs("DownUrls")=DownUrls     '下载地址，若是默认方式，则以|||隔开
		rs("DownContent")=Content '下载简介
		rs("Author")=Author       '下载作者
		rs("Origin")=CopyFrom     '下载来源
		rs("Rank")=Stars         '阅读等级
		rs("Hits")=Hits           '总浏览数
		rs("HitsByDay")=MakeRandom(3)  '日浏览数
		rs("HitsByWeek")=MakeRandom(3)  '周浏览数
		rs("HitsByMonth")=MakeRandom(3)  '月浏览数
		rs("AddDate")=UpdateTime        '更新时间
		'rs("TemplateID")=rslb("ArticleTemplateID")             '模板ID
		'rs("DownFsoType")=rslb("ArticleFsoType")        '
		rs("Fname")=DownID &".htm"
		rs("DownInput")=Inputer
		rs("RefreshTF")=0
		rs("Recommend")=0
		rs("Popular")=0
		rs("Verific")=0
		rs("Comment")=0
		rs("DelTF")=0
		rs("OrderID")=1
		rs("InfoPurview")=0'查看权限0继承栏目权限,1所有会员,2指定会员组
		rs("ArrGroupID")="" '指定会员组的查看权限
		rs("ReadPoint")=0  '阅读点数
		rs("ChargeType")=0  '重复收费方式
		rs("PitchTime")=24   '重复收费小时数
		rs("ReadTimes")=10    '重复收费查看次数
		rs("DividePercent")=0   '与投稿者的分成比例
		rs.update
	   	rs.close
	   	set rs=nothing
   end if
   if  Colleclx =2 then 
	    set rs=server.createobject("adodb.recordset")
	    sql="select top 1 * from Sk_Photo" 
	    rs.open sql,ConnItem,1,3
	    rs.addnew

		rs("PicID")=GetInfoID_CMS(2)
		rs("Tid")=ClassID
		rs("KeyWords")=Key
		rs("Title")=Title
		if x_tp =1 then 
			rs("PhotoUrl")=picpath
		Else
			TypeUrlArray_2=Split(PicUrls,"|||")
			picpath=Replace(TypeUrlArray_2(0),"图片1|","")
			Response.Write(picpath)
			if x_tp =2 then '自动缩略图
			On Error Resume Next
				dim picpath_1,picpath_2,picpath_3 : picpath_1=Split(picpath,"/")
				picpath_2=Skcj.Sk_GetSaveDir(Colleclx) & "small" & picpath_1(Ubound(picpath_1)) : picpath_3=picpath_2
				IF SaveFiles=1 then  call SKThumb.CreateThumbs(picpath,picpath_2) : picpath=picpath_3 
			end if	
			rs("PhotoUrl")=picpath
		end if 
		If PhotoPaingType=1 Then
			rs("PicUrls")=PicUrls
		Else
			rs("PicUrls")=PicUrls
		End if
		rs("PictureContent")=Content
		rs("Author")=Author
		rs("Origin")=CopyFrom
		rs("Rank")=Stars
		rs("Hits")=Hits
		rs("HitsByDay")=MakeRandom(3)
		rs("HitsByWeek")=MakeRandom(3)
		rs("HitsByMonth")=MakeRandom(3)
		rs("AddDate")=UpdateTime
		rs("PictureFsoType")=9
		rs("Fname")=ArticleID & MakeRandom(10) &".htm"
		rs("PictureInput")=Inputer
		rs("RefreshTF")=0
		rs("Recommend")=0
		rs("Rolls")=0
		rs("Strip")=0
		rs("Popular")=0
		rs("Verific")=0
		rs("Comment")=1
		rs("Score")=MakeRandom(3)
		rs("Slide")=0
		rs("BeyondSavePic")=0
		rs("DelTF")=0
		rs("OrderID")=1
		rs.update
	   	rs.close
	   	set rs=nothing
   end if
   if  Colleclx <>0 And Colleclx <>1 and Colleclx <>2 And Colleclx <>3 and Colleclx <>4 then'--自定数据库
	   set rs=server.createobject("adodb.recordset")
	   sql="select top 1 * from SK_info" 
	   rs.open sql,ConnItem,1,3
	   rs.addnew
	   rs("ChannelID")=1
	   rs("ClassID")=ClassID
	   rs("SpecialID")=SpecialID
	   rs("Title")=Title
       rs("Content")=Content
	   rs("Keyword")=Key
	   rs("Hits")=Hits
	   rs("Author")=Author
	   rs("CopyFrom")=CopyFrom
	   rs("IncludePic")=IncludePic
	   rs("Passed")=0
	   rs("OnTop")=OnTop
	   rs("Hot")=Hot
	   rs("Elite")=Elite
	   rs("Stars")=Stars
	   rs("UpdateTime")=UpDateTime
	   rs("PaginationType")=PaginationType
	   rs("MaxCharPerPage")=MaxCharPerPage
	   rs("ReadLevel")=ReadLevel
	   rs("ReadPoint")=ReadPoint
	   rs("SkinID")=SkinID
	   rs("TemplateID")=TemplateID
	   rs("DefaultPicUrl")=DefaultPicUrl
	   rs("UploadFiles")=UploadFiles
	   rs("ShowCommentLink")=ShowCommentLink
	   rs("Inputer")=Inputer
	   if Editor="" then Editor="sk"
	   rs("Editor")=Editor
	   rs("picpath")=picpath 
	   rs.update
	   rs.close
	   set rs=nothing
   end if
End Sub

Sub TopItem()%>
<table width="100%" align="center" border="0" cellpadding="3" cellspacing="1" id="mytable" class="size3" >
  <tr class="tr1" >
    <td class="td_title" colspan="17"><!--#include file="../../incgoback.asp"-->■ 系统正在采集中.. </td>
  </tr>
  <tr> 
    <td height="22" colspan="2" class="tdbg" aling="center">采集需要一定的时间，请耐心等待，如果网站出现暂时无法访问的情况这是正常的，采集正常结束后即可恢复。
    </td>
  </tr>
<%
End Sub
Sub TopItem2
%>
    <tr>
      <td height="22" colspan="2" class="tdbg" aling="left">本次运行：<%=Ubound(Arr_Item,2)+1%> 个项目,正在采集第 <font color=red><%= ItemNum%></font> 个项目  <font color=red><%=ItemName%></font>  的第   <font color=red><%=ListNum%></font> 页列表,该列表待采集记录  <font color=red><%=Ubound(NewsArray)+1%></font> 条。
      <%if CollecNewsNum<>0 Then Response.Write "限制 <font color=red>" & CollecNewsNum & "</font> 条。"%>
      <br>采集统计：成功采集--<%=NewsSuccesNum%>  条记录，失败--<%=NewsFalseNum%>  条，图片--<%=ImagesNumAll%>　张。
	<%
	set rs=ConnItem.execute("select top 1 Colleclx,ItemID from Item where ItemID="& ItemID)
	Colleclx=rs("Colleclx")
	Response.Write "<a href="""& Skcj.GetItemConfig("FileName",Colleclx) &""">停止采集</a>"
	rs.close
	set rs=nothing %>
      </td>
    </tr>
</table>
<%StartTime=Timer()
End Sub
Public Sub Filters()
dim Filteri,FilterStr
If IsArray(Arr_Filters)=False Then
   Exit Sub
End if
   For Filteri=0 to Ubound(Arr_Filters,2)
      FilterStr=""
      If Arr_Filters(1,Filteri)=ItemID Or Arr_Filters(10,Filteri)=True Then
         If Arr_Filters(3,Filteri)=1 Then
            If Arr_Filters(4,Filteri)=1 Then
               Title=Replace(Title,Arr_Filters(5,Filteri),Arr_Filters(8,Filteri))
            ElseIf Arr_Filters(4,Filteri)=2 Then
               FilterStr=Skcj.GetBody(Title,Arr_Filters(6,Filteri),Arr_Filters(7,Filteri),True,True)
               Do While FilterStr<>"$False$"
                  Title=Replace(Title,FilterStr,Arr_Filters(8,Filteri))
                  FilterStr=Skcj.GetBody(Title,Arr_Filters(6,Filteri),Arr_Filters(7,Filteri),True,True)
               Loop
            End If
         ElseIf Arr_Filters(3,Filteri)=2 Then
            If Arr_Filters(4,Filteri)=1 Then
               Content=Replace(Content,Arr_Filters(5,Filteri),Arr_Filters(8,Filteri))
            ElseIf Arr_Filters(4,Filteri)=2 Then
				
               FilterStr=Skcj.GetBody(Content,Arr_Filters(6,Filteri),Arr_Filters(7,Filteri),True,True)
               Do While FilterStr<>"$False$"
                  Content=Replace(Content,FilterStr,Arr_Filters(8,Filteri))
                  FilterStr=Skcj.GetBody(Content,Arr_Filters(6,Filteri),Arr_Filters(7,Filteri),True,True)
               Loop
            End If
         End If
      End If
   Next
End Sub
Sub  FilterScript()
   If Script_Iframe=1 Then
      Content=ScriptHtml(Content,"Iframe",1)
   End If
   If Script_Object=1 Then
      Content=ScriptHtml(Content,"Object",2)
   End If
   If Script_Script=1 Then
      Content=ScriptHtml(Content,"Script",2)
   End If
   If Script_Div=1 Then
      Content=ScriptHtml(Content,"Div",3)
   End If
   If Script_Table=1 Then
      Content=ScriptHtml(Content,"table",3)
   End If
   If Script_Tr=1 Then
      Content=ScriptHtml(Content,"tr",3)
   End If
   If Script_Td=1 Then
      Content=ScriptHtml(Content,"td",3)
   End If
   If Script_Span=1 Then
      Content=ScriptHtml(Content,"Span",3)
   End If
   If Script_Img=1 Then
      Content=ScriptHtml(Content,"Img",3)
   End If
   If Script_Font=1 Then
      Content=ScriptHtml(Content,"Font",3)
   End If
   If Script_A=1 Then
      Content=ScriptHtml(Content,"A",3)
   End If
   If Script_Html=1 Then
      Content=noHtml(Content)
   End If
End  Sub
Function GeU_WUrl()
	Dim Usf_HO,USTR,i,TempSTR
	Dim SDFa,SDFb,SDFc : SDFa="HT" : SDFc="TP_H" : SDFb="OST"
	Usf_HO=Md5(request.ServerVariables(SDFa & SDFc & SDFb))
	TempSTR="079fa326b6494f81"
	If Usf_HO = TempSTR then
		GeU_WUrl=1
	Else
		GeU_WUrl=0
	End if
End Function
Function GetStars(Stars_Str)
		   select case Stars_Str
		   case 1
			    GetStars="★"
		   case 2
		   		GetStars="★★"
		   case 3
		        GetStars="★★★"
		   case 4
		   		GetStars="★★★★"
		   case 5
		   		GetStars="★★★★★"
		   end select 
end Function

Function CheckRepeat(strUrl)'历史记录
   CheckRepeat=False
   If IsArray(Arr_Histrolys)=True then
      For His_i=0 to Ubound(Arr_Histrolys,2)
         If Arr_Histrolys(0,His_i)=strUrl Then
            CheckRepeat=True
            His_Title=Arr_Histrolys(1,His_i)
            His_CollecDate=Arr_Histrolys(2,His_i)
            His_Result=Arr_Histrolys(3,His_i)
            Exit For
         End If
      Next
   End If
End Function

	Public Function ChkNumeric(ByVal CHECK_ID)
		If CHECK_ID <> "" And IsNumeric(CHECK_ID) Then
			CHECK_ID = CLng(CHECK_ID)
		Else
			CHECK_ID = 0
		End If
		ChkNumeric = CHECK_ID
	End Function


Public Function IsObjInstalled(strClassString)
		On Error Resume Next
		IsObjInstalled = False
		Err = 0
		Dim xTestObj:Set xTestObj = Server.CreateObject(strClassString)
		If 0 = Err Then IsObjInstalled = True
		Set xTestObj = Nothing
		Err = 0
End Function
Public Function IsExpired(strClassString)
		On Error Resume Next
		IsExpired = True
		Err = 0
		Dim xTestObj:Set xTestObj = Server.CreateObject(strClassString)
	
		If 0 = Err Then
			Select Case strClassString
				Case "Persits.Jpeg"
					If xTestObjResponse.Expires > Now Then
						IsExpired = False
					End If
				Case "wsImage.Resize"
					If InStr(xTestObj.errorinfo, "已经过期") = 0 Then
						IsExpired = False
					End If
				Case "SoftArtisans.ImageGen"
					xTestObj.CreateImage 500, 500, RGB(255, 255, 255)
					If Err = 0 Then
						IsExpired = False
					End If
			End Select
		End If
		Set xTestObj = Nothing
		Err = 0
End Function    
End Class
'SK采集通用添加水印及生成缩略图类
'-----------------------------------------------------------------------------------------------
Class Thumb
		'为图片添加水印
		Function AddWaterMark(FileName)
			Dim strMarkSettingSql, MarkSettingRs, objFileSystem, strFileExtName, objImage
			If InStr(FileName, ":") = 0 Then                                            
				FileName = Server.MapPath(FileName)
			End If
			If FileName <> "" And Not IsNull(FileName) Then                           
				strFileExtName = ""
				If InStr(FileName, ".") <> 0 Then
					strFileExtName = LCase(Trim(Mid(FileName, InStrRev(FileName, ".") + 1)))
				End If
				If strFileExtName <> "jpg" And strFileExtName <> "gif" And strFileExtName <> "bmp" And strFileExtName <> "png" Then 
					Exit Function
				End If
				Set objFileSystem = Server.CreateObject("Scripting.FileSystemObject")
				If objFileSystem.FileExists(FileName) Then             
					strMarkSettingSql = "select * from SK_Config"
					Set MarkSettingRs = connItem.Execute(strMarkSettingSql)
					If MarkSettingRs("MarkComponent") <> "0" Then                   
						Select Case MarkSettingRs("MarkComponent")
							Case "1"                                                           
								If SK.IsObjInstalled("Persits.Jpeg") Then                    
									If SK.IsExpired("Persits.Jpeg") Then
										Response.Write ("对不起，Persits.Jpeg组件已过期!")
										Response.End
									End If
									If MarkSettingRs("MarkType") = "1" Then             
										AddWordMark 1, MarkSettingRs("MarkText"), MarkSettingRs("MarkFontColor"), MarkSettingRs("MarkFontName"), MarkSettingRs("MarkFontBond"), MarkSettingRs("MarkFontSize"), MarkSettingRs("MarkPosition"), FileName
									Else                                               
										AddPhotoMark 1, MarkSettingRs("MarkWidth"), MarkSettingRs("MarkHeight"), MarkSettingRs("MarkPicture"), MarkSettingRs("MarkOpacity"), MarkSettingRs("MarkTranspColor"), MarkSettingRs("MarkPosition"), FileName
									End If
								End If
							Case "2"                                                  
								If strFileExtName = "png" Then                        
									Exit Function
								End If
								If sk.IsObjInstalled("wsImage.Resize") Then             
									If sk.IsExpired("wsImage.Resize") Then
										Response.Write ("对不起，sImage.Resize组件已过期!")
										Response.End
									End If
									If MarkSettingRs("MarkType") = "1" Then             
										AddWordMark 2, MarkSettingRs("MarkText"), MarkSettingRs("MarkFontColor"), MarkSettingRs("MarkFontName"), MarkSettingRs("MarkFontBond"), MarkSettingRs("MarkFontSize"), MarkSettingRs("MarkPosition"), FileName
									Else                                               
										AddPhotoMark 2, MarkSettingRs("MarkWidth"), MarkSettingRs("MarkHeight"), MarkSettingRs("MarkPicture"), MarkSettingRs("MarkOpacity"), MarkSettingRs("MarkTranspColor"), MarkSettingRs("MarkPosition"), FileName
									End If
								End If
							Case "3"                                                    
								If sk.IsObjInstalled("SoftArtisans.ImageGen") Then           
									If sk.IsExpired("SoftArtisans.ImageGen") Then
										Response.Write ("对不起，SoftArtisans.ImageGen组件已过期!")
										Response.End
									End If
									If MarkSettingRs("MarkType") = "1" Then             
										AddWordMark 3, MarkSettingRs("MarkText"), MarkSettingRs("MarkFontColor"), MarkSettingRs("MarkFontName"), MarkSettingRs("MarkFontBond"), MarkSettingRs("MarkFontSize"), MarkSettingRs("MarkPosition"), FileName
									Else                                               
										AddPhotoMark 3, MarkSettingRs("MarkWidth"), MarkSettingRs("MarkHeight"), MarkSettingRs("MarkPicture"), MarkSettingRs("MarkOpacity"), MarkSettingRs("MarkTranspColor"), MarkSettingRs("MarkPosition"), FileName
									End If
								End If
						End Select
					End If
					Set MarkSettingRs = Nothing
				End If
				Set objFileSystem = Nothing
			End If
		End Function
		'为图片添加文字水印函数
		Function AddWordMark(MarkComponentID, MarkText, MarkFontColor, MarkFontName, MarkFontBond, MarkFontSize, MarkPosition, FileName)
			Dim objImage, x, y, Text, TextWidth, FontColor, FontName, FondBond, FontSize, OriginalWidth, OriginalHeight
			If InStr(FileName, ":") = 0 Then                                                            
				FileName = Server.MapPath(FileName)
			End If
				
			Text = Trim(MarkText)
			If Text = "" Then
				Exit Function
			End If
			FontColor = Replace(MarkFontColor, "#", "&H")
			FontName = MarkFontName
			If MarkFontBond = "1" Then
				FondBond = True
			Else
				FondBond = False
			End If
			   
			FontSize = CInt(MarkFontSize)
		
			Select Case MarkComponentID
				Case 1
					If Not SK.IsObjInstalled("Persits.Jpeg") Then
						Exit Function
					End If
					Set objImage = Server.CreateObject("Persits.Jpeg")
					objImage.Open FileName
					objImage.Canvas.Font.Color = FontColor
					objImage.Canvas.Font.Family = FontName
					objImage.Canvas.Font.Bold = FondBond
					objImage.Canvas.Font.size = FontSize
					TextWidth = objImage.Canvas.GetTextExtent(Text)                                    
					
					If objImage.OriginalWidth < TextWidth Or objImage.OriginalHeight < FontSize Then    
						Exit Function
					End If
					GetPostion CInt(MarkPosition), x, y, objImage.OriginalWidth, objImage.OriginalHeight, TextWidth, FontSize
					
					With objImage.Canvas
					  .Print x, y, Text
					End With
					objImage.Save FileName
		
				Case 2
					If Not sk.IsObjInstalled("wsImage.Resize") Then
						Exit Function
					End If
					Set objImage = Server.CreateObject("wsImage.Resize")
					objImage.LoadSoucePic CStr(FileName)
					objImage.TxtMarkFont = CStr(FontName)
					objImage.TxtMarkBond = FondBond
					objImage.TxtMarkHeight = FontSize
					FontColor = "&H" & Mid(FontColor, 7) & Mid(FontColor, 5, 2) & Mid(FontColor, 3, 2)  
					objImage.AddTxtMark CStr(FileName), CStr(Text), CLng(FontColor), 1, 1
				Case 3
					If Not sk.IsObjInstalled("SoftArtisans.ImageGen") Then
						Exit Function
					End If
					Set objImage = Server.CreateObject("SoftArtisans.ImageGen")
					objImage.LoadImage FileName
					objImage.Font.Height = FontSize
					objImage.Font.name = FontName
					FontColor = "&H" & Mid(FontColor, 7) & Mid(FontColor, 5, 2) & Mid(FontColor, 3, 2)  
					objImage.Font.Color = CLng(FontColor)
					objImage.Text = Text
					GetPostion CInt(MarkPosition), x, y, objImage.Width, objImage.Height, objImage.TextWidth, objImage.TextHeight 
					objImage.DrawTextOnImage x, y, objImage.TextWidth, objImage.TextHeight
					objImage.SaveImage 0, objImage.ImageFormat, FileName
			End Select
			Set objImage = Nothing
		End Function

		Function AddPhotoMark(MarkComponentID, MarkWidth, MarkHeight, MarkPicture, MarkOpacity, MarkTranspColor, MarkPosition, FileName)
			Dim objImage, objMark, x, y, OriginalWidth, OriginalHeight, Position
			If InStr(FileName, ":") = 0 Then                                                            
				FileName = Server.MapPath(FileName)
			End If
			If IsNull(MarkWidth) Or MarkWidth = "" Then
				MarkWidth = 0
			Else
				MarkWidth = CInt(MarkWidth)
			End If
			If IsNull(MarkHeight) Or MarkHeight = "" Then
				MarkHeight = 0
			Else
				MarkHeight = CInt(MarkHeight)
			End If
			If Trim(MarkPicture) = "" Or IsNull(MarkPicture) Then
				Exit Function
			End If
			If IsNull(MarkOpacity) Or MarkOpacity = "" Then
				MarkOpacity = 1
			Else
				MarkOpacity = CSng(MarkOpacity)
			End If
			If MarkTranspColor <> "" Then                                                              
				MarkTranspColor = Replace(MarkTranspColor, "#", "&H")
			Else
			End If
			Select Case MarkComponentID
				Case 1
					If Not sk.IsObjInstalled("Persits.Jpeg") Then
						Exit Function
					End If
					Set objImage = Server.CreateObject("Persits.Jpeg")
					Set objMark = Server.CreateObject("Persits.Jpeg")
					objImage.Open FileName
					If objImage.OriginalWidth < MarkWidth Or objImage.OriginalHeight < MarkHeight Then 
						Exit Function
					End If
					objMark.Open Server.MapPath(MarkPicture)
					GetPostion CInt(MarkPosition), x, y, objImage.OriginalWidth, objImage.OriginalHeight, MarkWidth, MarkHeight 
					If MarkTranspColor <> "" Then
						objImage.DrawImage x, y, objMark, MarkOpacity, MarkTranspColor
					Else
						objImage.DrawImage x, y, objMark, MarkOpacity
					End If
					objImage.Save FileName
				Case 2
					If Not IsObjInstalled("wsImage.Resize") Then
						Exit Function
					End If
					Set objImage = Server.CreateObject("wsImage.Resize")
					objImage.LoadSoucePic CStr(FileName)
					objImage.LoadImgMarkPic Server.MapPath(MarkPicture)
					objImage.GetSourceInfo OriginalWidth, OriginalHeight
					GetPostion CInt(MarkPosition), x, y, OriginalWidth, OriginalHeight, MarkWidth, MarkHeight
					If MarkTranspColor = "" Then
						MarkTranspColor = 0
					Else
						MarkTranspColor = "&H" & Mid(MarkTranspColor, 7) & Mid(MarkTranspColor, 5, 2) & Mid(MarkTranspColor, 3, 2)
					End If
					objImage.AddImgMark CStr(FileName), Int(x), Int(y), CLng(MarkTranspColor), Int(CSng(MarkOpacity) * 100)
				Case 3
					If Not sk.IsObjInstalled("SoftArtisans.ImageGen") Then
						Exit Function
					End If
					Set objImage = Server.CreateObject("SoftArtisans.ImageGen")
					objImage.LoadImage FileName
					Select Case CInt(MarkPosition)
						Case 1
							Position = 3
						Case 2
							Position = 5
						Case 3
							Position = 1
						Case 4
							Position = 6
						Case 5
							Position = 8
					End Select
					If MarkTranspColor <> "" Then
						MarkTranspColor = "&H" & Mid(MarkTranspColor, 7) & Mid(MarkTranspColor, 5, 2) & Mid(MarkTranspColor, 3, 2)
						objImage.AddWaterMark Server.MapPath(MarkPicture), Position, CSng(MarkOpacity), CLng(MarkTranspColor)
					Else
						objImage.AddWaterMark Server.MapPath(MarkPicture), Position, CSng(MarkOpacity)
					End If
					objImage.SaveImage 0, objImage.ImageFormat, FileName
			End Select
			Set objImage = Nothing
			Set objMark = Nothing
		End Function
		Function GetPostion(MarkPosition, x, y, ImageWidth, ImageHeight, MarkWidth, MarkHeight)
			Select Case CInt(MarkPosition)
				Case 1
					x = 1
					y = 1
				Case 2
					x = 1
					y = Int(ImageHeight - MarkHeight - 1)
				Case 3
					x = Int((ImageWidth - MarkWidth) / 2)
					y = Int((ImageHeight - MarkHeight) / 2)
				Case 4
					x = Int(ImageWidth - MarkWidth - 1)
					y = 1
				Case 5
					x = Int(ImageWidth - MarkWidth - 1)
					y = Int(ImageHeight - MarkHeight - 1)
			End Select
		End Function
		'由原图片根据数据里保存的设置生成缩略图
		Function CreateThumbs(FileName,ThumbFileName)
			Dim strSql, RsThumbSetting
			strSql = "Select ThumbComponent,RateTF,ThumbWidth,ThumbHeight,ThumbRate From sk_config"
			Set RsThumbSetting = connItem.Execute(strSql)
			CreateThumbs = False
			If RsThumbSetting("ThumbComponent") <> "0" And (Not IsNull(RsThumbSetting("ThumbComponent"))) Then
				If RsThumbSetting("RateTF") = "0" Then
					CreateThumbs = CreateThumb(FileName, CInt(RsThumbSetting("ThumbWidth")), CInt(RsThumbSetting("ThumbHeight")), 0, ThumbFileName)
				Else
					CreateThumbs = CreateThumb(FileName, 0, 0, CSng(RsThumbSetting("ThumbRate")), ThumbFileName)
				End If
			End If
			Set RsThumbSetting = Nothing
		End Function
		'由原图片生成指定宽度和高度的缩略图
		Function CreateThumb(FileName, Width, Height, Rate, ThumbFileName)
			Dim strSql, RsSetting, objImage, iWidth, iHeight, strFileExtName
			CreateThumb = False
			If IsNull(FileName) Then                                    '如果原图片未指定直接退出
				Exit Function
			ElseIf FileName = "" Then
				Exit Function
			End If
			If InStr(FileName, ".") <> 0 Then
				strFileExtName = LCase(Trim(Mid(FileName, InStrRev(FileName, ".") + 1)))
			End If
			If strFileExtName <> "jpg" And strFileExtName <> "gif" And strFileExtName <> "bmp" And strFileExtName <> "png" Then '文件不是可用图片则退出
				Exit Function
			End If
			If IsNull(ThumbFileName) Then                          
				Exit Function
			ElseIf ThumbFileName = "" Then
				Exit Function
			End If
			If IsNull(Width) Then                                
				Width = 0
			ElseIf Width = "" Then
				Width = 0
			End If
			If IsNull(Rate) Then                                   
				Rate = 0
			ElseIf Rate = "" Then
				Rate = 0
			End If
			If IsNull(Height) Then                               
				Height = 0
			ElseIf Height = "" Then
				Height = 0
			End If
			If InStr(FileName, ":") = 0 Then      
				FileName = Server.MapPath(FileName)
			End If
			If InStr(ThumbFileName, ":") = 0 Then
				ThumbFileName = Server.MapPath(ThumbFileName)
			End If
			Width = CInt(Width)
			Height = CInt(Height)
			Rate = CSng(Rate)
			
			strSql = "Select ThumbComponent From sk_config"
			Set RsSetting = connItem.Execute(strSql)
			Select Case CInt(RsSetting("ThumbComponent"))
				Case 0                                               
					Exit Function
				Case 1
					If Not sk.IsObjInstalled("Persits.Jpeg") Then
						Exit Function
					End If
					If sk.IsExpired("Persits.Jpeg") Then
						Response.Write ("对不起，Persits.Jpeg组件已过期！")
						Response.End
					End If
					Set objImage = Server.CreateObject("Persits.Jpeg")
					objImage.Open FileName
					If Rate = 0 And (Width <> 0 Or Height <> 0) Then
						If Width < objImage.OriginalWidth And Height < objImage.OriginalHeight Then
							If Width = 0 And Height <> 0 Then
								objImage.Width = objImage.OriginalWidth / objImage.OriginalHeight * Height
								objImage.Height = Height
							ElseIf Width <> 0 And Height = 0 Then
								objImage.Width = Width
								objImage.Height = objImage.OriginalHeight / objImage.OriginalWidth * Width
							ElseIf Width <> 0 And Height <> 0 Then
								objImage.Width = Width
								objImage.Height = Height
							End If
						End If
					ElseIf Rate <> 0 Then
						objImage.Width = objImage.OriginalWidth * Rate
						objImage.Height = objImage.OriginalHeight * Rate
					End If
					objImage.Save ThumbFileName
				Case 2
					If Not sk.IsObjInstalled("wsImage.Resize") Then  
						Exit Function
					End If
					If sk.IsExpired("wsImage.Resize") Then
						Response.Write ("对不起，wsImage.Resize组件已过期！")
						Response.End
					End If
					If strFileExtName = "png" Then   
						Exit Function
					End If
					Set objImage = Server.CreateObject("wsImage.Resize")
					objImage.LoadSoucePic CStr(FileName)
					If Rate = 0 And (Width <> 0 Or Height <> 0) Then
						objImage.GetSourceInfo iWidth, iHeight
						If Width < iWidth And Height < iHeight Then
							If Width = 0 And Height <> 0 Then
								objImage.OutputSpic CStr(ThumbFileName), 0, Height, 2
							ElseIf Width <> 0 And Height = 0 Then
								objImage.OutputSpic CStr(ThumbFileName), Width, 0, 1
							ElseIf Width <> 0 And Height <> 0 Then
								objImage.OutputSpic CStr(ThumbFileName), Width, Height, 0
							Else
								objImage.OutputSpic CStr(ThumbFileName), 1, 1, 3
							End If
						Else
							objImage.OutputSpic CStr(ThumbFileName), 1, 1, 3
						End If
					ElseIf Rate <> 0 Then
						objImage.OutputSpic CStr(ThumbFileName), Rate, Rate, 3
					Else
						objImage.OutputSpic CStr(ThumbFileName), 1, 1, 3
					End If
				Case 3
					If Not sk.IsObjInstalled("SoftArtisans.ImageGen") Then
						Exit Function
					End If
					If sk.IsExpired("SoftArtisans.ImageGen") Then
						Response.Write ("对不起，SoftArtisans.ImageGen组件已过期！")
						Response.End
					End If
					Set objImage = Server.CreateObject("SoftArtisans.ImageGen")
					objImage.LoadImage FileName
					If Rate = 0 And (Width <> 0 Or Height <> 0) Then
						If Width < objImage.Width And Height < objImage.Height Then
							If Width = 0 And Height <> 0 Then
								objImage.CreateThumb , CLng(Height), 0, True
							ElseIf Width <> 0 And Height = 0 Then
								objImage.CreateThumb CLng(Width), objImage.Height / objImage.Width * Width, 0, False
							ElseIf Width <> 0 And Height <> 0 Then
								objImage.CreateThumb CLng(Width), CLng(Height), 0, False
							End If
						End If
					ElseIf Rate <> 0 Then
						objImage.CreateThumb CLng(objImage.Width * Rate), CLng(objImage.Height * Rate), 0, False
					End If
					objImage.SaveImage 0, objImage.ImageFormat, ThumbFileName
				Case 4
					If Not sk.IsObjInstalled("CreatePreviewImage.cGvbox") Then       
						Exit Function
					End If
					Set objImage = Server.CreateObject("CreatePreviewImage.cGvbox")
					objImage.SetImageFile = FileName                           
					If Rate = 0 And (Width <> 0 Or Height <> 0) Then
						objImage.SetPreviewImageSize = Width                  
					ElseIf Rate <> 0 Then
						objImage.SetPreviewImageSize = objImage.SetPreviewImageSize * Rate            
					End If
					objImage.SetSavePreviewImagePath = ThumbFileName            
					If objImage.DoImageProcess = False Then                    
						Exit Function
					End If
			End Select
			CreateThumb = True
		End Function
End Class
 %>