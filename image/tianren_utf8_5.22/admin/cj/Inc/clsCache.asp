<%
class SK_clsCache
'----------------------------
private cache           '缓存内容
private cacheName       '缓存Application名称
private expireTime      '缓存过期时间
private expireTimeName  '缓存过期时间Application名称
private path            '缓存页URL路径
private vaild           'ansir添加
private sub class_initialize()
path=request.servervariables("url")
path=left(path,instrRev(path,"/"))
end sub

private sub class_terminate()
end sub

Public Property Get Version
	Version="缓存类2006"
End Property

public property get valid '读取缓存是否有效/属性
	if isempty(cache) or (not isdate(expireTime)) then
	vaild=false
	else
	valid=true
	end if
	end property
	
public property get value '读取当前缓存内容/属性
	if isempty(cache) or (not isDate(expireTime)) then
	value=null
	elseif CDate(expireTime)<now then
	value=null
	else
	value=cache
	end if
end property

public property let name(str) '设置缓存名称/属性
	cacheName=str&path
	cache=application(cacheName)
	expireTimeName=str&"expire"&path
	expireTime=application(expireTimeName)
end property

public property let expire(tm) '设置缓存过期时间/属性
	expireTime=tm
	application.Lock()
	application(expireTimeName)=expireTime
	application.UnLock()
end property

public sub add(varCache,varExpireTime) '对缓存赋值/方法
	if isempty(varCache) or not isDate(varExpireTime) then
	exit sub
	end if
	cache=varCache
	expireTime=varExpireTime
	application.lock
	application(cacheName)=cache
	application(expireTimeName)=expireTime
	application.unlock
end sub

public sub clean() '释放缓存/方法
	application.lock
	application(cacheName)=empty
	application(expireTimeName)=empty
	application.unlock
	cache=empty
	expireTime=empty
end sub
 
public function verify(varcache2) '比较缓存值是否相同/方法——返回是或否
	if typename(cache)<>typename(varcache2) then
		verify=false
	elseif typename(cache)="Object" then
		if cache is varcache2 then
			verify=true
		else
			verify=false
		end if
	elseif typename(cache)="Variant()" then
		if join(cache,"^")=join(varcache2,"^") then
			verify=true
		else
			verify=false
		end if
	else
		if cache=varcache2 then
			verify=true
		else
			verify=false
		end if
	end if
end function
end class
%>
