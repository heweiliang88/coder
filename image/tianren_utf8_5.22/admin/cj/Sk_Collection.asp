<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../../inc/charsetasp.asp"-->
<%
'采集内核：SK采集系统
'QQ81962480 天人文章管理系统
%>
<%
'option explicit
Response.Buffer = True 
Server.ScriptTimeOut=999
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache"
%>
<!doctype html>
<html>
<head>
<title>天人文章管理系统采集模块</title>
<meta charset="utf-8">

<!--#include file="inc/conn.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="inc/ubbcode.asp"-->
<!--#include file="inc/SK_FunCls.asp"-->
<!--#include file="inc/clsCache.asp"-->
<!--#include file="inc/cj_cls.asp"-->
<%
If InStr(","&adminlimitidrange&",", ",5,")<1 Then contrl_message "您没有访问权限，请尝试其他版块。", "", "", "", ""
If adminctrllimit = 1 Then Call contrl_message ("您只有查看权限，没有管理权限，请升级权限后继续操作。", "", "", "", "")

dim Skcj
Set Skcj= New FunCls
Dim Action,CollecType
Dim myCache
Dim ItemNum,ListNum,PaingNum,NewsSuccesNum,NewsFalseNum,NewsNum_i,Itemon,ItemIdstr,Itemok
Dim Sql,RsItem,SqlItem,FoundErr,ItemEnd,ListEnd
Dim PicUrls_i,NewsUrlPaing_s,NewsUrlPaing_o,NewsPaingNext_Code,TypeArray_Url,TypeNews_Url

'项目变量
Dim ItemID,ItemName,ChannelID,strChannelDir,ClassID,SpecialID,LoginType,LoginUrl,LoginPostUrl,LoginUser,LoginPass,LoginFalse
Dim ListStr,LsString,LoString,ListPaingType,LPsString,LPoString,ListPaingStr1,ListPaingStr2,ListPaingID1,ListPaingID2,ListPaingStr3,HsString,HoString,HttpUrlType,HttpUrlStr,photourls,photourlo,PhotoPaingType,PhotoType_s,PhotoType_o,PhotoLurl_s,PhotoLurl_o,Phototypefy_s,Phototypefy_o,Phototypefyurl_s,Phototypefyurl_o,Phototypeurl_s,Phototypeurl_o,Colleclx,selEncoding,SaveFileUrl,x_tpUrl,Thumb_WaterMark,Thumbs_Create,Timing,strReplace

'下载变量
dim DownSize,DownYY,DownSQ,DownPT,YSDZ,ZCDZ,PhotoUrl,DownUrls
'下载变量项目字段
dim Downlist_s,Downlist_o,DownUrl_s,DownUrl_o,DownNewType,DownNewlist_s,DownNewlist_o,DownNewUrl_s,DownNewUrl_o,LinkUrlYn
dim ZdType_001,Zds_001,Zdo_001,ZD_001,ZdType_002,Zds_002,Zdo_002,ZD_002,ZdType_003,Zds_003,Zdo_003,ZD_003,ZdType_004,Zds_004,Zdo_004,ZD_004,ZdType_005,Zds_005,Zdo_005,ZD_005,ZdType_006,Zds_006,Zdo_006,ZD_006,ZdType_007,Zds_007,Zdo_007,ZD_007,ZdType_008,Zds_008,Zdo_008,ZD_008

'--图片列表链接
dim imhstr,imostr,NewsimageCode,Newsimage,picpath,Radiobutton,x_tp
'--图片列表链接
Dim TsString,ToString,CsString,CoString,DateType,DsString,DoString,AuthorType,AsString,AoString,AuthorStr,CopyFromType,FsString,FoString
Dim CopyFromStr,KeyType,KsString,KoString,KeyStr,NewsPaingType,NPsString,NpoString,NewsPaingStr,NewsPaingHtml
Dim ItemCollecDate,PaginationType,MaxCharPerPage,ReadLevel,Stars,ReadPoint,Hits,UpDateType,UpDateTime,IncludePicYn,DefaultPicYn,OnTop,Elite,Hot
Dim SkinID,TemplateID,Script_Iframe,Script_Object,Script_Script,Script_Div,Script_Class,Script_Span,Script_Img,Script_Font,Script_A,Script_Html,CollecListNum,CollecNewsNum,Passed,SaveFiles,CollecOrder,InputerType,Inputer,EditorType,Editor,ShowCommentLink,Script_Table,Script_Tr,Script_Td

'过滤变量
Dim Arr_Filters,FilterStr,Filteri

'采集相关的变量
Dim ContentTemp,NewsPaingNext,NewsPaingNextCode,Arr_i,NewsUrl,NewsCode,ListTypeCode,ListTypeUrlCode,TypeUrlArray,TypeNewsUrl,NewsTypeCode,PicUrls,Arr_ii,Arr_ii_2,ListTypeCode_2,ListTypeUrlCode_2,TypeUrlArray_2

'文章保存变量
Dim ArticleID,Title,Content,Author,CopyFrom,Key,IncludePic,UploadFiles,DefaultPicUrl,Coll_DefiniteUrl
'其它变量
Dim LoginData,LoginResult,OrderTemp,i
Dim Arr_Item,CollecTest,Content_View,CollecNewsAll
Dim StepID

'历史记录
Dim Arr_Histrolys,His_Title,His_CollecDate,His_Result,His_Repeat,His_i 

'执行时间变量
Dim StartTime,OverTime

'图片统计
Dim Arr_Images,ImagesNum,ImagesNumAll

'列表
Dim ListUrl,ListCode,NewsArrayCode,NewsArray,ListArray,ListPaingNext

'安装路径
Dim strInstallDir,CacheTemp
Dim DiyFieldSTR_z,DiyFieldSTR_l'自定义
Dim FoundErr_1
'----是否登陆---------------------
'Call Skcj.Admin()
'If Skcj.IsAdmin=False Then
'	ErrMsg="<li> 您没有登陆或不是管理员。请<a href='sk_login.asp' target='_top'>登陆</a>。"
'	response.Redirect("Sk_err.asp?action=AdminErr&ErrMsg="&ErrMsg&"")
'	response.End()
'End If
'-------------------------------------
'On Error Resume Next
strInstallDir=trim(request.ServerVariables("SCRIPT_NAME"))
strInstallDir=left(strInstallDir,instrrev(lcase(strInstallDir),"/")-1)
'strInstallDir=left(strInstallDir,instrrev(lcase(strInstallDir),"/"))
'缓存路径
CacheTemp=Lcase(trim(request.ServerVariables("SCRIPT_NAME")))
CacheTemp=left(CacheTemp,instrrev(CacheTemp,"/"))
CacheTemp=replace(CacheTemp,"\","_")
CacheTemp=replace(CacheTemp,"/","_")
CacheTemp="ansir" & CacheTemp

'数据初始化
CollecListNum=0
CollecNewsNum=0
ArticleID=0
ItemNum=Clng(Trim(Request("ItemNum")))
ListNum=Clng(Trim(Request("ListNum")))
NewsNum_i=Clng(Trim(Request("NewsNum_i")))
NewsSuccesNum=Clng(Trim(Request("NewsSuccesNum")))
NewsFalseNum=Clng(Trim(Request("NewsFalseNum")))
ImagesNumAll=Clng(Trim(Request("ImagesNumAll")))
ListPaingNext=Trim(Request("ListPaingNext"))
Itemon=Trim(Request("Itemon"))
Itemok=Trim(Request("Itemok"))
FoundErr=False
ItemEnd=False
ListEnd=False
ErrMsg=""

Call DelNews()'
Call CheckForm()''检察ItemID值
Dim Collecdate : Collecdate=Trim(Request("Collecdate"))
if Itemok = "yes" then 
	If Instr(Itemon,",")>0 Then
		ItemIdstr=GetItemId(Itemon,1)
		response.write("<script>location.href='Sk_Collection.asp?ItemID="&GetItemId(Itemon,0)&"&ItemNum=1&ListNum=1&NewsSuccesNum=0&NewsFalseNum=0&ImagesNumAll=0&NewsNum_i=0&Itemon="& ItemIdstr &"&Collecdate="& Collecdate &"';</script>")'到页面
	Else
		response.write("<script>location.href='Sk_Collection.asp?ItemID="&Itemon&"&ItemNum=1&ListNum=1&NewsSuccesNum=0&NewsFalseNum=0&ImagesNumAll=0&NewsNum_i=0&Collecdate="& Collecdate &"';</script>")'到页面
	End if
	Response.end
End if
If Instr(ItemID,",")>0 Then 
	ItemIdstr=GetItemId(ItemID,1)
	response.write("<script>location.href='Sk_Collection.asp?ItemID="&GetItemId(ItemID,0)&"&ItemNum=1&ListNum=1&NewsSuccesNum=0&NewsFalseNum=0&ImagesNumAll=0&NewsNum_i=0&Itemon="& ItemIdstr&"&Collecdate="& Collecdate &"';</script>")'到页面
	Response.end
End if

If FoundErr<>True Then
  Call SetCache()'项目信息写入缓存
End If
If FoundErr=True Then
   Call WriteErrMsg(ErrMsg)
Else
   Call GetCache()
   Call Main()
   sk.CollPhoto_Show
End If
'关闭数据库链接
Call CloseConnItem()
%>
<%Sub Main%>
<link rel="stylesheet" type="text/css" href="../skin/default/style.css">
</head>
<body style="background:#fff; " class="w1">
</body>         
</html>
<%End Sub

Sub CheckForm()'提取表单
   ItemID=Trim(Request("ItemID"))
   'CollecType=Trim(Request.Form("CollecType"))
   CollecTest=Trim(Request.Form("CollecTest"))
   Content_View=Trim(Request.Form("Content_View"))
   '检察表单
   If ItemID="" Then
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>请您选择项目!</li>"
   Else
      If Instr(ItemID,",")>0 Then
         ItemID=Replace(ItemID," ","")
      End If
	  Response.Flush()
'	  set rs=connItem.execute("select top 1 * from Item Where ItemID in(" & ItemID &")" )
'   	  if connItem.Execute("select count(ClassID) from SK_Class Where ClassID=" & RS("ClassID"))(0)=0 then
'	  'if conn.Execute("select count(id) from ks_Class Where ID='" & RS("ClassID") &"'")(0)=0 then
'	  FoundErr=True
'	  ErrMsg=ErrMsg & "<br><li>请您设置频道栏目! </li>"
'	  rs.close
'	  set rs=nothing
'	  end if
   End If 
   If CollecTest="yes" Then
      CollecTest=True
   Else
      CollecTest=False
   End If
   If Content_View="yes" Then
      Content_View=True
   Else
      Content_View=False
   End If
End Sub

'==================================================
'过程名：SetCache1
'作  用：存取缓存
'参  数：无
'==================================================
Sub GetCache()
   Dim myCache
   Set myCache=new SK_clsCache

   '项目信息
   myCache.name=CacheTemp & "items"
   If myCache.valid then 
      Arr_Item=myCache.value
   Else
      ItemEnd=True
   End If

   '过滤信息
   myCache.name=CacheTemp & "filters"
   If myCache.valid then 
      Arr_Filters=myCache.value
   End If

   '历史记录
   myCache.name=CacheTemp & "histrolys"
   If myCache.valid then 
      Arr_Histrolys=myCache.value
   End If

   '其它信息
   myCache.name=CacheTemp & "collectest"
   If myCache.valid then 
      CollecTest=myCache.value
   Else
      CollecTest=False
   End If
   myCache.name=CacheTemp & "contentview"
   If myCache.valid then 
      Content_View=myCache.value
   Else
      Content_View=False
   End If

   Set myCache=Nothing
End Sub

Sub SetCache()'项目信息写入缓存
   SqlItem ="select * from Item where ItemID in(" & ItemID & ")"
   Set RsItem=Server.CreateObject("adodb.recordset")
   RsItem.Open SqlItem,ConnItem,1,1
   If Not RsItem.Eof Then
      Arr_Item=RsItem.GetRows()
   End If
   RsItem.Close
   Set RsItem=Nothing

   Set myCache=new SK_clsCache
   myCache.name=CacheTemp & "items"
   Call myCache.clean()
   If IsArray(Arr_Item)=True Then
      myCache.add Arr_Item,Dateadd("n",1000,now)
   Else
      FoundErr=True
      ErrMsg=ErrMsg & "<br>发生意外错误！"
   End If

   '过滤信息
   SqlItem ="select * from Filters where Flag=True"
   Set RsItem=Server.CreateObject("adodb.recordset")
   RsItem.Open SqlItem,ConnItem,1,1
   If Not RsItem.Eof Then
      Arr_Filters=RsItem.GetRows()
   End If
   RsItem.Close
   Set Rsitem=Nothing

   myCache.name=CacheTemp & "filters"
   Call myCache.clean()
   If IsArray(Arr_Filters)=True Then
      myCache.add Arr_Filters,Dateadd("n",1000,now)
   End If

   '历史记录
   SqlItem ="select NewsUrl,Title,CollecDate,Result from Histroly"
   Set RsItem=Server.CreateObject("adodb.recordset")
   RsItem.Open SqlItem,ConnItem,1,1
   If Not RsItem.Eof Then
      Arr_Histrolys=RsItem.GetRows()
   End If
   RsItem.Close
   Set RsItem=Nothing

   myCache.name=CacheTemp & "histrolys"
   Call myCache.clean()
   If IsArray(Arr_Histrolys)=True Then
      myCache.add Arr_Histrolys,Dateadd("n",1000,now)
   End If

   '其它信息
   myCache.name=CacheTemp & "collectest"
   Call myCache.clean()
   myCache.add CollecTest,Dateadd("n",1000,now)

   myCache.name=CacheTemp & "contentview"
   Call myCache.clean()
   myCache.add Content_View,Dateadd("n",1000,now)

   set myCache=nothing
End Sub
Sub DelNews()
   ConnItem.execute("Delete From [NewsList]")
End Sub
%>
