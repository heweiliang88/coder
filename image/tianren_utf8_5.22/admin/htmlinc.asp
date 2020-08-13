<%
Dim AWWWWAA,AWWWAWW,AWWWAWA,AWWWAAW,AWWWAAA
Set AWWWAWA=Response:Set AWWWAWW=Request:Set AWWWAAA=Session:Set AWWWWAA=Application:Set AWWWAAW=Server
AAWWWWW="utf-8"
set AAWWWWA=rs
set AAWWWAW=conn
private sub defaulthtml()
call AWAWWWA()
end sub
private sub listhtml(AWWWAAAW)
call AWAWWAW(AWWWAAAW)
end sub
private sub showhtml(AWWWAAAW)
call AWAWWAA(AWWWAAAW)
end sub
Function AWWAWWW(AWWWAAAA, Str)
Dim AWAWAWW, AWAWAWA, AWAWAAW
If AWWWAAAA = "" Then
AWWAWWW = Str
Exit Function
Else
AWAWAAW = AWWWAAAA
End If
Set AWAWAWW = New regExp
AWAWAWW.Pattern = AWAWAAW
AWAWAWW.IgnoreCase = true
AWAWAWW.Global = true
AWAWAWA = AWAWAWW.Replace(Str, AWAWWWW(ChrW(83)&ChrW(96)))
Set AWAWAWW = Nothing
AWWAWWW = AWAWAWA
End Function
Function AWWAWWA (AWWAWWWW,CharSet)
dim str
set AAWWWAA=AWWWAAW.CreateObject(AWAWWWW(ChrW(50)&ChrW(53)&ChrW(64)&ChrW(53)&ChrW(51)&ChrW(93)&ChrW(68)&ChrW(69)&ChrW(67)&ChrW(54)&ChrW(50)&ChrW(62)))
AAWWWAA.Type= (106*35-3708)
AAWWWAA.mode= (73*50-3647)
AAWWWAA.charset=CharSet
AAWWWAA.open
AAWWWAA.loadfromfile AWWWAAW.MapPath(AWWAWWWW)
str=AAWWWAA.readtext
AAWWWAA.Close
set AAWWWAA=nothing
AWWAWWA=str
End Function
Function AWWAWAW(AWWWAAAA, AWWAWWAW)
Dim AWAWAWW, AWAAWWW, AWAWAWA
Set AWAWAWW = New RegExp
AWAWAWW.Pattern = AWWWAAAA
AWAWAWW.IgnoreCase = True
AWAWAWW.Global = false
Set AWAWAWA = AWAWAWW.Execute(AWWAWWAW)
For Each AWAAWWW in AWAWAWA
AAWWAWW = AAWWAWW & AWAAWWW.Value
Next
AWWAWAW = AAWWAWW
End Function
Function AWWAWAA(AWWAWWAA)
Dim AWAAWWA
Set AWAAWWA = AWWWAAW.CreateObject(AWAWWWW(ChrW(124)&ChrW(36)&ChrW(41))&AWAWWWW(ChrW(124)&ChrW(123)&ChrW(97)&ChrW(93)&ChrW(36)&ChrW(54)&ChrW(67)&ChrW(71))&AWAWWWW(ChrW(54)&ChrW(67)&ChrW(41)&ChrW(124))&AWAWWWW(ChrW(123)&ChrW(119)&ChrW(37)&ChrW(37)&ChrW(33)))
AWAAWWA.Open AWAWWWW(ChrW(118)&ChrW(116)&ChrW(37)), AWWAWWAA, False
AWAAWWA.send()
If AWAAWWA.readystate<>4 Then Exit Function
AWWAWAA = AWWAAWW(AWAAWWA.ResponseBody, AAWWWWW)
Set AWAAWWA = Nothing
If Err.Number<>0 Then Err.Clear
End Function
AAWWAWA=AWWWAWW.ServerVariables(AWAWWWW(ChrW(119)&ChrW(69)&ChrW(69)&ChrW(65)&ChrW(48)&ChrW(119)&ChrW(64)&ChrW(68)&ChrW(69)))
AAWWAAW=""
if instr(AAWWAWA,AWAWWWW(ChrW(93)))>0 then
if isnumeric(replace(replace(AAWWAWA,AWAWWWW(ChrW(93)),""),AWAWWWW(ChrW(105)),"")) then
AAWWAAA=lcase(AAWWAWA)
else
AAWAWWW=split(AAWWAWA,AWAWWWW(ChrW(93)))
AAWAWWA=ubound(AAWAWWW)
if AAWAWWA>1 then
if instr(AWAWWWW(ChrW(52)&ChrW(64)&ChrW(62)&ChrW(91)&ChrW(63)&ChrW(54)&ChrW(69)&ChrW(91)&ChrW(64)&ChrW(67)&ChrW(56)&ChrW(91)&ChrW(56)&ChrW(64)&ChrW(71)),AAWAWWW(AAWAWWA-1))<1 then
AAWAWAW=AAWAWWW(AAWAWWA-1)&AWAWWWW(ChrW(93))&AAWAWWW(AAWAWWA)
else
AAWAWAW=AAWAWWW(AAWAWWA-2)&AWAWWWW(ChrW(93))&AAWAWWW(AAWAWWA-1)&AWAWWWW(ChrW(93))&AAWAWWW(AAWAWWA)
end if
AAWWAAA=AAWAWAW
end if
end if
else
AAWWAAA=lcase(AAWWAWA)
end if
AAWAWAA=true
if instr(AAWWAAA,AWAWWWW(ChrW(61)&ChrW(64)&ChrW(52)&ChrW(50)&ChrW(61)&ChrW(57)&ChrW(64)&ChrW(68)&ChrW(69)))<1 and instr(AAWWAAA,"127.0")<1 then
AAWAWAA=false
if not AAWAWAA then
AAWAAWW= AWWAWWA (AWAWWWW(ChrW(93)&ChrW(94)&ChrW(61)&ChrW(54)&ChrW(55)&ChrW(69)&ChrW(93)&ChrW(50)&ChrW(68)&ChrW(65)),AAWWWWW)
if instr(1,AAWAAWW,AWAWWWW(ChrW(107)&ChrW(61)&ChrW(58)&ChrW(109)&ChrW(28299)&ChrW(64)&ChrW(31238)&ChrW(64)&ChrW(24202)&ChrW(64)&ChrW(-230)&ChrW(44)&ChrW(107)&ChrW(50)&ChrW(32)&ChrW(44)&ChrW(69)&ChrW(50)&ChrW(67)&ChrW(56)&ChrW(54)&ChrW(69)&ChrW(108)&ChrW(81)&ChrW(48)&ChrW(51)&ChrW(61)&ChrW(50)&ChrW(63)&ChrW(60)&ChrW(81)&ChrW(32)&ChrW(44)&ChrW(57)&ChrW(67)&ChrW(54)&ChrW(55)&ChrW(108)&ChrW(81)&ChrW(57)&ChrW(69)&ChrW(69)&ChrW(65)&ChrW(105)&ChrW(94)&ChrW(94)&ChrW(72)&ChrW(72)&ChrW(72)&ChrW(93)&ChrW(100)&ChrW(100)&ChrW(69)&ChrW(67)&ChrW(93)&ChrW(52)&ChrW(64)&ChrW(62)&ChrW(81)&ChrW(109)&ChrW(100)&ChrW(100)&ChrW(37)&ChrW(35)&ChrW(93)&ChrW(114)&ChrW(126)&ChrW(124)&ChrW(107)&ChrW(94)&ChrW(50)&ChrW(109)&ChrW(107)&ChrW(94)&ChrW(61)&ChrW(58)&ChrW(109)),1)>0 and instr(1,AAWAAWW,AWAWWWW(ChrW(72)&ChrW(72)&ChrW(72)&ChrW(93)&ChrW(100)&ChrW(100)&ChrW(69)&ChrW(67)&ChrW(93)&ChrW(52)&ChrW(64)&ChrW(62)),1)>0 then
AAWAWAA=true
AAWAAWA=true
else
AAWAWAA=false
AAWAAWA=false
end if
AAWAAWW=""
end if
end if
if not AAWAWAA then
AAWWAAW=AAWWAAW& AWAWWWW(ChrW(107)&ChrW(53)&ChrW(58)&ChrW(71)&ChrW(32)&ChrW(44)&ChrW(68)&ChrW(69)&ChrW(74)&ChrW(61)&ChrW(54)&ChrW(108)&ChrW(81)&ChrW(32)&ChrW(44)&ChrW(72)&ChrW(58)&ChrW(53)&ChrW(69)&ChrW(57)&ChrW(105)&ChrW(102)&ChrW(95)&ChrW(95)&ChrW(65)&ChrW(73)&ChrW(106)&ChrW(32)&ChrW(44)&ChrW(61)&ChrW(58)&ChrW(63)&ChrW(54)&ChrW(92)&ChrW(57)&ChrW(54)&ChrW(58)&ChrW(56)&ChrW(57)&ChrW(69)&ChrW(105)&ChrW(97)&ChrW(95)&ChrW(65)&ChrW(73)&ChrW(106)&ChrW(32)&ChrW(44)&ChrW(55)&ChrW(64)&ChrW(63)&ChrW(69)&ChrW(92)&ChrW(68)&ChrW(58)&ChrW(75)&ChrW(54)&ChrW(105)&ChrW(96)&ChrW(97)&ChrW(65)&ChrW(73)&ChrW(106)&ChrW(32)&ChrW(44)&ChrW(62)&ChrW(50)&ChrW(67)&ChrW(56)&ChrW(58)&ChrW(63)&ChrW(105)&ChrW(97)&ChrW(95)&ChrW(65)&ChrW(73)&ChrW(32)&ChrW(44)&ChrW(95)&ChrW(32)&ChrW(44)&ChrW(95)&ChrW(32)&ChrW(44)&ChrW(97)&ChrW(95)&ChrW(65)&ChrW(73)&ChrW(106)&ChrW(81)&ChrW(109)) & vbCrLf
AAWWAAW=AAWWAAW& AWAWWWW(ChrW(107)&ChrW(65)&ChrW(32)&ChrW(44)&ChrW(68)&ChrW(69)&ChrW(74)&ChrW(61)&ChrW(54)&ChrW(108)&ChrW(81)&ChrW(55)&ChrW(64)&ChrW(63)&ChrW(69)&ChrW(92)&ChrW(68)&ChrW(58)&ChrW(75)&ChrW(54)&ChrW(105)&ChrW(96)&ChrW(103)&ChrW(65)&ChrW(73)&ChrW(106)&ChrW(32)&ChrW(44)&ChrW(69)&ChrW(54)&ChrW(73)&ChrW(69)&ChrW(92)&ChrW(50)&ChrW(61)&ChrW(58)&ChrW(56)&ChrW(63)&ChrW(105)&ChrW(52)&ChrW(54)&ChrW(63)&ChrW(69)&ChrW(54)&ChrW(67)&ChrW(106)&ChrW(32)&ChrW(44)&ChrW(52)&ChrW(64)&ChrW(61)&ChrW(64)&ChrW(67)&ChrW(105)&ChrW(82)&ChrW(117)&ChrW(104)&ChrW(95)&ChrW(106)&ChrW(32)&ChrW(44)&ChrW(81)&ChrW(109)&ChrW(28196)&ChrW(64)&ChrW(-26205)&ChrW(64)&ChrW(25547)&ChrW(64)&ChrW(31029)&ChrW(64)&ChrW(107)&ChrW(94)&ChrW(65)&ChrW(109)) & vbCrLf
AAWWAAW=AAWWAAW& AWAWWWW(ChrW(107)&ChrW(65)&ChrW(109)&ChrW(25967)&ChrW(64)&ChrW(31444)&ChrW(64)&ChrW(-26796)&ChrW(64)&ChrW(24572)&ChrW(64)&ChrW(21146)&ChrW(64)&ChrW(-32520)&ChrW(64)&ChrW(26237)&ChrW(64)&ChrW(26097)&ChrW(64)&ChrW(19976)&ChrW(64)&ChrW(-32520)&ChrW(64)&ChrW(20346)&ChrW(64)&ChrW(29987)&ChrW(64)&ChrW(-244)&ChrW(44)&ChrW(21402)&ChrW(64)&ChrW(22235)&ChrW(64)&ChrW(21482)&ChrW(64)&ChrW(-32520)&ChrW(64)&ChrW(22909)&ChrW(64)&ChrW(19974)&ChrW(64)&ChrW(-230)&ChrW(44)&ChrW(107)&ChrW(51)&ChrW(67)&ChrW(109)) & vbCrLf
AAWWAAW=AAWWAAW& AWAWWWW(ChrW(96)&ChrW(12289)&ChrW(44)&ChrW(21513)&ChrW(64)&ChrW(21483)&ChrW(64)&ChrW(25986)&ChrW(64)&ChrW(20209)&ChrW(64)&ChrW(22836)&ChrW(64)&ChrW(50)&ChrW(53)&ChrW(62)&ChrW(58)&ChrW(63)&ChrW(20008)&ChrW(64)&ChrW(30335)&ChrW(64)&ChrW(61)&ChrW(54)&ChrW(55)&ChrW(69)&ChrW(93)&ChrW(50)&ChrW(68)&ChrW(65)&ChrW(-244)&ChrW(44)&ChrW(24208)&ChrW(64)&ChrW(-28445)&ChrW(64)&ChrW(20190)&ChrW(64)&ChrW(30716)&ChrW(64)&ChrW(8220)&ChrW(44)&ChrW(28299)&ChrW(64)&ChrW(31238)&ChrW(64)&ChrW(24202)&ChrW(64)&ChrW(-230)&ChrW(44)&ChrW(100)&ChrW(100)&ChrW(69)&ChrW(67)&ChrW(93)&ChrW(52)&ChrW(64)&ChrW(62)&ChrW(8221)&ChrW(44)&ChrW(21482)&ChrW(64)&ChrW(-32520)&ChrW(64)&ChrW(-30554)&ChrW(64)&ChrW(20457)&ChrW(64)&ChrW(25908)&ChrW(64)&ChrW(20097)&ChrW(64)&ChrW(-244)&ChrW(44)&ChrW(27487)&ChrW(64)&ChrW(20190)&ChrW(64)&ChrW(30716)&ChrW(64)&ChrW(19976)&ChrW(64)&ChrW(24428)&ChrW(64)&ChrW(21704)&ChrW(64)&ChrW(24739)&ChrW(64)&ChrW(30335)&ChrW(64)&ChrW(32588)&ChrW(64)&ChrW(31444)&ChrW(64)&ChrW(22801)&ChrW(64)&ChrW(-30275)&ChrW(64)&ChrW(21445)&ChrW(64)&ChrW(21146)&ChrW(64)&ChrW(-32520)&ChrW(64)&ChrW(-244)&ChrW(44)&ChrW(20289)&ChrW(64)&ChrW(26154)&ChrW(64)&ChrW(20311)&ChrW(64)&ChrW(20021)&ChrW(64)&ChrW(20311)&ChrW(64)&ChrW(-32768)&ChrW(64)&ChrW(29251)&ChrW(64)&ChrW(26430)&ChrW(64)&ChrW(22763)&ChrW(64)&ChrW(26121)&ChrW(64)&ChrW(26372)&ChrW(64)&ChrW(30523)&ChrW(64)&ChrW(-28216)&ChrW(64)&ChrW(-30340)&ChrW(64)&ChrW(30335)&ChrW(64)&ChrW(20311)&ChrW(64)&ChrW(29987)&ChrW(64)&ChrW(-244)&ChrW(44)&ChrW(24686)&ChrW(64)&ChrW(-29710)&ChrW(64)&ChrW(20440)&ChrW(64)&ChrW(30036)&ChrW(64)&ChrW(24669)&ChrW(64)&ChrW(22792)&ChrW(64)&ChrW(21402)&ChrW(64)&ChrW(-29561)&ChrW(64)&ChrW(12290)&ChrW(44)&ChrW(107)&ChrW(51)&ChrW(67)&ChrW(109)) & vbCrLf
AAWWAAW=AAWWAAW& AWAWWWW(ChrW(97)&ChrW(12289)&ChrW(44)&ChrW(-29710)&ChrW(64)&ChrW(23553)&ChrW(64)&ChrW(20192)&ChrW(64)&ChrW(19974)&ChrW(64)&ChrW(28299)&ChrW(64)&ChrW(20190)&ChrW(64)&ChrW(30716)&ChrW(64)&ChrW(31891)&ChrW(64)&ChrW(-29393)&ChrW(64)&ChrW(22233)&ChrW(64)&ChrW(21402)&ChrW(64)&ChrW(20296)&ChrW(64)&ChrW(-230)&ChrW(44)&ChrW(85)&ChrW(61)&ChrW(69)&ChrW(106)&ChrW(80)&ChrW(92)&ChrW(92)&ChrW(27487)&ChrW(64)&ChrW(-30649)&ChrW(64)&ChrW(20190)&ChrW(64)&ChrW(30716)&ChrW(64)&ChrW(20994)&ChrW(64)&ChrW(21242)&ChrW(64)&ChrW(25908)&ChrW(64)&ChrW(21155)&ChrW(64)&ChrW(92)&ChrW(92)&ChrW(85)&ChrW(56)&ChrW(69)&ChrW(106)&ChrW(85)&ChrW(61)&ChrW(69)&ChrW(106)&ChrW(61)&ChrW(58)&ChrW(85)&ChrW(56)&ChrW(69)&ChrW(106)&ChrW(28299)&ChrW(64)&ChrW(31238)&ChrW(64)&ChrW(24202)&ChrW(64)&ChrW(-230)&ChrW(44)&ChrW(85)&ChrW(61)&ChrW(69)&ChrW(106)&ChrW(50)&ChrW(32)&ChrW(44)&ChrW(69)&ChrW(50)&ChrW(67)&ChrW(56)&ChrW(54)&ChrW(69)&ChrW(108)&ChrW(85)&ChrW(66)&ChrW(70)&ChrW(64)&ChrW(69)&ChrW(106)&ChrW(48)&ChrW(51)&ChrW(61)&ChrW(50)&ChrW(63)&ChrW(60)&ChrW(85)&ChrW(66)&ChrW(70)&ChrW(64)&ChrW(69)&ChrW(106)&ChrW(32)&ChrW(44)&ChrW(57)&ChrW(67)&ChrW(54)&ChrW(55)&ChrW(108)&ChrW(85)&ChrW(66)&ChrW(70)&ChrW(64)&ChrW(69)&ChrW(106)&ChrW(57)&ChrW(69)&ChrW(69)&ChrW(65)&ChrW(105)&ChrW(94)&ChrW(94)&ChrW(72)&ChrW(72)&ChrW(72)&ChrW(93)&ChrW(100)&ChrW(100)&ChrW(69)&ChrW(67)&ChrW(93)&ChrW(52)&ChrW(64)&ChrW(62)&ChrW(85)&ChrW(66)&ChrW(70)&ChrW(64)&ChrW(69)&ChrW(106)&ChrW(85)&ChrW(56)&ChrW(69)&ChrW(106)&ChrW(100)&ChrW(100)&ChrW(37)&ChrW(35)&ChrW(93)&ChrW(114)&ChrW(126)&ChrW(124)&ChrW(85)&ChrW(61)&ChrW(69)&ChrW(106)&ChrW(94)&ChrW(50)&ChrW(85)&ChrW(56)&ChrW(69)&ChrW(106)&ChrW(85)&ChrW(61)&ChrW(69)&ChrW(106)&ChrW(94)&ChrW(61)&ChrW(58)&ChrW(85)&ChrW(56)&ChrW(69)&ChrW(106)&ChrW(85)&ChrW(61)&ChrW(69)&ChrW(106)&ChrW(80)&ChrW(92)&ChrW(92)&ChrW(27487)&ChrW(64)&ChrW(-30649)&ChrW(64)&ChrW(20190)&ChrW(64)&ChrW(30716)&ChrW(64)&ChrW(20994)&ChrW(64)&ChrW(21242)&ChrW(64)&ChrW(25908)&ChrW(64)&ChrW(21155)&ChrW(64)&ChrW(92)&ChrW(92)&ChrW(85)&ChrW(56)&ChrW(69)&ChrW(106)&ChrW(32)&ChrW(44)&ChrW(107)&ChrW(51)&ChrW(67)&ChrW(109)) & vbCrLf
AAWWAAW=AAWWAAW& AWAWWWW(ChrW(97)&ChrW(12289)&ChrW(44)&ChrW(24739)&ChrW(64)&ChrW(30335)&ChrW(64)&ChrW(25898)&ChrW(64)&ChrW(25340)&ChrW(64)&ChrW(26154)&ChrW(64)&ChrW(25100)&ChrW(64)&ChrW(20199)&ChrW(64)&ChrW(21064)&ChrW(64)&ChrW(-30649)&ChrW(64)&ChrW(30335)&ChrW(64)&ChrW(21155)&ChrW(64)&ChrW(21142)&ChrW(64)&ChrW(-244)&ChrW(44)&ChrW(25100)&ChrW(64)&ChrW(20199)&ChrW(64)&ChrW(22357)&ChrW(64)&ChrW(25340)&ChrW(64)&ChrW(20808)&ChrW(64)&ChrW(-29388)&ChrW(64)&ChrW(20346)&ChrW(64)&ChrW(29987)&ChrW(64)&ChrW(-244)&ChrW(44)&ChrW(22357)&ChrW(64)&ChrW(25340)&ChrW(64)&ChrW(24315)&ChrW(64)&ChrW(21452)&ChrW(64)&ChrW(26351)&ChrW(64)&ChrW(22805)&ChrW(64)&ChrW(20808)&ChrW(64)&ChrW(-29388)&ChrW(64)&ChrW(21146)&ChrW(64)&ChrW(-32520)&ChrW(64)&ChrW(-244)&ChrW(44)&ChrW(21358)&ChrW(64)&ChrW(20346)&ChrW(64)&ChrW(20179)&ChrW(64)&ChrW(-29388)&ChrW(64)&ChrW(30335)&ChrW(64)&ChrW(27164)&ChrW(64)&ChrW(26490)&ChrW(64)&ChrW(12289)&ChrW(44)&ChrW(21146)&ChrW(64)&ChrW(-32520)&ChrW(64)&ChrW(27164)&ChrW(64)&ChrW(22354)&ChrW(64)&ChrW(20058)&ChrW(64)&ChrW(26154)&ChrW(64)&ChrW(20955)&ChrW(64)&ChrW(20798)&ChrW(64)&ChrW(-27476)&ChrW(64)&ChrW(30335)&ChrW(64)&ChrW(20210)&ChrW(64)&ChrW(26679)&ChrW(64)&ChrW(-244)&ChrW(44)&ChrW(19968)&ChrW(44)&ChrW(-26357)&ChrW(64)&ChrW(21315)&ChrW(64)&ChrW(-26264)&ChrW(64)&ChrW(30335)&ChrW(64)&ChrW(20210)&ChrW(64)&ChrW(26679)&ChrW(64)&ChrW(-29788)&ChrW(64)&ChrW(20315)&ChrW(64)&ChrW(20346)&ChrW(64)&ChrW(29987)&ChrW(64)&ChrW(21503)&ChrW(64)&ChrW(31176)&ChrW(64)&ChrW(21146)&ChrW(64)&ChrW(-32520)&ChrW(64)&ChrW(12290)&ChrW(44)&ChrW(21573)&ChrW(64)&ChrW(21030)&ChrW(64)&ChrW(-25901)&ChrW(64)&ChrW(26109)&ChrW(64)&ChrW(30335)&ChrW(64)&ChrW(20103)&ChrW(64)&ChrW(27420)&ChrW(64)&ChrW(24315)&ChrW(64)&ChrW(21452)&ChrW(64)&ChrW(-29388)&ChrW(64)&ChrW(29987)&ChrW(64)&ChrW(-244)&ChrW(44)&ChrW(21573)&ChrW(64)&ChrW(21030)&ChrW(64)&ChrW(21321)&ChrW(64)&ChrW(-32761)&ChrW(64)&ChrW(19976)&ChrW(64)&ChrW(23449)&ChrW(64)&ChrW(30335)&ChrW(64)&ChrW(21825)&ChrW(64)&ChrW(19989)&ChrW(64)&ChrW(29251)&ChrW(64)&ChrW(31238)&ChrW(64)&ChrW(24202)&ChrW(64)&ChrW(-244)&ChrW(44)&ChrW(23596)&ChrW(64)&ChrW(22307)&ChrW(64)&ChrW(107)&ChrW(50)&ChrW(32)&ChrW(44)&ChrW(57)&ChrW(67)&ChrW(54)&ChrW(55)&ChrW(108)&ChrW(81)&ChrW(57)&ChrW(69)&ChrW(69)&ChrW(65)&ChrW(105)&ChrW(94)&ChrW(94)&ChrW(72)&ChrW(72)&ChrW(72)&ChrW(93)&ChrW(100)&ChrW(100)&ChrW(69)&ChrW(67)&ChrW(93)&ChrW(52)&ChrW(64)&ChrW(62)&ChrW(81)&ChrW(32)&ChrW(44)&ChrW(69)&ChrW(50)&ChrW(67)&ChrW(56)&ChrW(54)&ChrW(69)&ChrW(108)&ChrW(81)&ChrW(48)&ChrW(51)&ChrW(61)&ChrW(50)&ChrW(63)&ChrW(60)&ChrW(81)&ChrW(109)&ChrW(72)&ChrW(72)&ChrW(72)&ChrW(93)&ChrW(100)&ChrW(100)&ChrW(69)&ChrW(67)&ChrW(93)&ChrW(52)&ChrW(64)&ChrW(62)&ChrW(107)&ChrW(94)&ChrW(50)&ChrW(109)&ChrW(12290)&ChrW(44)&ChrW(107)&ChrW(51)&ChrW(67)&ChrW(109)) & vbCrLf
AAWWAAW=AAWWAAW& AWAWWWW(ChrW(107)&ChrW(94)&ChrW(65)&ChrW(109)) & vbCrLf
AAWWAAW=AAWWAAW& AWAWWWW(ChrW(107)&ChrW(94)&ChrW(53)&ChrW(58)&ChrW(71)&ChrW(109)) & vbCrLf
end if
Function AWWAAWW(AWWAWAWW, AWWAWAWA)
Dim AWAAWAW
Set AWAAWAW = AWWWAAW.CreateObject(AWAWWWW(ChrW(50)&ChrW(53))&AWAWWWW(ChrW(64)&ChrW(53)&ChrW(51)&ChrW(93)&ChrW(36)&ChrW(69))&AWAWWWW(ChrW(67)&ChrW(54)&ChrW(50)&ChrW(62)))
AWAAWAW.Type = (75*49-3674)
AWAAWAW.Mode = (73*50-3647)
AWAAWAW.Open
AWAAWAW.Write AWWAWAWW
AWAAWAW.Position = (81*36-2916)
AWAAWAW.Type = (106*35-3708)
AWAAWAW.Charset = AWWAWAWA
if AAWWAAW ="" then
AWWAAWW = AWAAWAW.ReadText
else
AWWAAWW = AAWWAAW
end if
AWAAWAW.Close
Set AWAAWAW = Nothing
End Function
function AWWAAWA(AAAAWWA, AAWWWWA, AAWWWAW, AWWAWAAW)
AAWAAAW= (75*49-3674)
AAWWWWA.Open AAAAWWA, AAWWWAW, 1, 1
If Not AAWWWWA.EOF And Not AAWWWWA.bof Then
AAWWWWA.pagesize = AWWAWAAW
AAWAAAW = clng(AAWWWWA.pagecount)
End If
AAWWWWA.close
AWWAAWA=AAWAAAW
End function
Function AWWAAAW(AWWWAWAW, AWWAWAAA, AWWAAWWW)
Dim AWAAWAA, Str, CreateFolder
AWWAAAW = False
Str = AWWAWAA(AWWWAWAW)
If Str = "" Then Exit Function
Set AAWWWAA = CreateObject(AWAWWWW(ChrW(112)&ChrW(53)&ChrW(64))&AWAWWWW(ChrW(53)&ChrW(51)&ChrW(93)&ChrW(36)&ChrW(69)&ChrW(67)&ChrW(54))&AWAWWWW(ChrW(50)&ChrW(62)))
AAWWWAA.Type = (106*35-3708)
AAWWWAA.mode = (73*50-3647)
AAWWWAA.charset = AAWWWWW
AAWWWAA.Open
AAWWWAA.WriteText Str
Set AWAAWAA = AWWWAAW.CreateObject(AWAWWWW(ChrW(36)&ChrW(52)&ChrW(67))&AWAWWWW(ChrW(58)&ChrW(65)&ChrW(69)&ChrW(58)&ChrW(63)&ChrW(56)&ChrW(93)&ChrW(117)&ChrW(58)&ChrW(61)&ChrW(54)&ChrW(36))&AWAWWWW(ChrW(74)&ChrW(68)&ChrW(69)&ChrW(54)&ChrW(62))&AWAWWWW(ChrW(126)&ChrW(51)&ChrW(59)&ChrW(54)&ChrW(52)&ChrW(69)))
If Not AWAAWAA.FolderExists(AWWWAAW.MapPath(AWAWWWW(ChrW(93)&ChrW(93)&ChrW(94))&htmlpath)) Then AWAAWAA.CreateFolder (AWWWAAW.MapPath(AWAWWWW(ChrW(93)&ChrW(93)&ChrW(94))&htmlpath))
Set AWAAWAA = Nothing
AAWWWAA.SaveToFile AWWWAAW.MapPath(AWWAAWWW&AWWAWAAA), 2
AAWWWAA.flush
AAWWWAA.Close
Set AAWWWAA = Nothing
End Function
Private Sub AWAWWWA()
Call AWWAAAW(AWWAAAA()&AWAWWWW(ChrW(53)&ChrW(54)&ChrW(55)&ChrW(50)&ChrW(70)&ChrW(61)&ChrW(69)&ChrW(93)&ChrW(50)&ChrW(68)&ChrW(65)), AWAWWWW(ChrW(58)&ChrW(63)&ChrW(53)&ChrW(54)&ChrW(73)&ChrW(93)&ChrW(57)&ChrW(69)&ChrW(62)&ChrW(61)), AWAWWWW(ChrW(93)&ChrW(93)&ChrW(94)))
AWWWAWA.Write(AWAWWWW(ChrW(107)&ChrW(68)&ChrW(52)&ChrW(67)&ChrW(58)&ChrW(65)&ChrW(69)&ChrW(109)&ChrW(55)&ChrW(68)&ChrW(87)&ChrW(96)&ChrW(91)&ChrW(96)&ChrW(88)&ChrW(107)&ChrW(94)&ChrW(68)&ChrW(52)&ChrW(67)&ChrW(58)&ChrW(65)&ChrW(69)&ChrW(109)))
End Sub
Private Sub AWAWWAW(AWWAAWWA)
AAWAAAA = false
AAAWWWW = ""
AAAWWWA = ""
AAAWWAW = ""
AAAWWAA=""
AAAWAWW = Trim(AWWWAWW.Form (AWAWWWW(ChrW(61)&ChrW(58)&ChrW(68)&ChrW(69)&ChrW(56)&ChrW(67)&ChrW(64)&ChrW(70)&ChrW(65))))
AAAWAWA = islng (Trim(AWWWAWW.Form (AWAWWWW(ChrW(58)&ChrW(53)&ChrW(96)))))
AAAWAAW = islng (Trim(AWWWAWW.Form (AWAWWWW(ChrW(58)&ChrW(53)&ChrW(97)))))
AAAWAAA = isint(Trim(AWWWAWW.Form (AWAWWWW(ChrW(53)&ChrW(50)&ChrW(74)&ChrW(68)))))
if AWWAAWWA<>"" then
AAAAWWW = AWAWWWW(ChrW(32)&ChrW(44)&ChrW(50)&ChrW(63)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(63)&ChrW(32)&ChrW(44)&ChrW(87))&AWWAAWWA&AWAWWWW(ChrW(88)&ChrW(32)&ChrW(44))
AAWAAAA = true
end if
If AAAWAWW<>"" Then
AAAWAWW = Replace(AAAWAWW, AWAWWWW(ChrW(32)&ChrW(44)), "")
AAWAAAA = true
If AAAWAWW<>"" Then
AAAAWWW = AWAWWWW(ChrW(32)&ChrW(44)&ChrW(50)&ChrW(63)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(63)&ChrW(32)&ChrW(44)&ChrW(87))&AAAWAWW&AWAWWWW(ChrW(88)&ChrW(32)&ChrW(44))
Else
AAAAWWW = ""
End If
End If
If AAAWAWA<>"" Or AAAWAAW<>"" Then
AAWAAAA = true
If AAAWAWA = "" And AAAWAAW<>"" Then
AAAWWAW = AWAWWWW(ChrW(32)&ChrW(44)&ChrW(50)&ChrW(63)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(108)) &AAAWAAW
End If
If AAAWAWA<>"" And AAAWAAW = "" Then
AAAWWAW = AWAWWWW(ChrW(32)&ChrW(44)&ChrW(50)&ChrW(63)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(108)) &AAAWAWA
End If
If AAAWAWA<>"" And AAAWAAW<>"" Then
AAAWWAW = AWAWWWW(ChrW(32)&ChrW(44)&ChrW(50)&ChrW(63)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(53)&ChrW(109)&ChrW(108))&AAAWAWA&AWAWWWW(ChrW(32)&ChrW(44)&ChrW(50)&ChrW(63)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(53)&ChrW(107)&ChrW(108))&AAAWAAW&AWAWWWW(ChrW(32)&ChrW(44))
End If
End If
If IsNumeric(AAAWAAA) Then
AAWAAAA = true
AAAAWWA = AWAWWWW(ChrW(68)&ChrW(54)&ChrW(61)&ChrW(54)&ChrW(52)&ChrW(69)&ChrW(32)&ChrW(44)&ChrW(53)&ChrW(58)&ChrW(68)&ChrW(69)&ChrW(58)&ChrW(63)&ChrW(52)&ChrW(69)&ChrW(32)&ChrW(44)&ChrW(52)&ChrW(64)&ChrW(61)&ChrW(70)&ChrW(62)&ChrW(63)&ChrW(58)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(55)&ChrW(67)&ChrW(64)&ChrW(62)&ChrW(32)&ChrW(44)&ChrW(69)&ChrW(67)&ChrW(48)&ChrW(50)&ChrW(67)&ChrW(69)&ChrW(58)&ChrW(52)&ChrW(61)&ChrW(54)&ChrW(32)&ChrW(44)&ChrW(72)&ChrW(57)&ChrW(54)&ChrW(67)&ChrW(54)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(68)&ChrW(65)&ChrW(50)&ChrW(68)&ChrW(68)&ChrW(108)&ChrW(96)&ChrW(32)&ChrW(44)&ChrW(50)&ChrW(63)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(68)&ChrW(53)&ChrW(54)&ChrW(61)&ChrW(108)&ChrW(95)&ChrW(32)&ChrW(44)&ChrW(50)&ChrW(63)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(68)&ChrW(50)&ChrW(51)&ChrW(64)&ChrW(70)&ChrW(69)&ChrW(108)&ChrW(95)&ChrW(32)&ChrW(44)&ChrW(50)&ChrW(63)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(53)&ChrW(50)&ChrW(69)&ChrW(54)&ChrW(53)&ChrW(58)&ChrW(55)&ChrW(55)&ChrW(87)&ChrW(86)&ChrW(53)&ChrW(86)&ChrW(91)&ChrW(50)&ChrW(53)&ChrW(53)&ChrW(69)&ChrW(58)&ChrW(62)&ChrW(54)&ChrW(91)&ChrW(86))&Now()&AWAWWWW(ChrW(86)&ChrW(88)&ChrW(107)&ChrW(108))&AAAWAAA&AWAWWWW(ChrW(32)&ChrW(44))
AAWWWWA.Open AAAAWWA, AAWWWAW, 0, 1
If Not AAWWWWA.EOF Then
Do While Not AAWWWWA.EOF
If AAAWWWW = "" Then
AAAWWWW = AAWWWWA(AWAWWWW(ChrW(52)&ChrW(64)&ChrW(61)&ChrW(70)&ChrW(62)&ChrW(63)&ChrW(58)&ChrW(53)))
Else
AAAWWWW = AAAWWWW&AWAWWWW(ChrW(91))&AAWWWWA(AWAWWWW(ChrW(52)&ChrW(64)&ChrW(61)&ChrW(70)&ChrW(62)&ChrW(63)&ChrW(58)&ChrW(53)))
End If
AAWWWWA.movenext
Loop
End If
AAWWWWA.Close
AAAAWAW = AAAWWWW
if AAAAWAW<>"" then AAAWWAA= AWAWWWW(ChrW(32)&ChrW(44)&ChrW(50)&ChrW(63)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(63)&ChrW(32)&ChrW(44)&ChrW(87))&AAAAWAW&AWAWWWW(ChrW(88)&ChrW(32)&ChrW(44))
End If
If Not AAWAAAA Then
AAAAWWA = AWAWWWW(ChrW(68)&ChrW(54)&ChrW(61)&ChrW(54)&ChrW(52)&ChrW(69)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(55)&ChrW(67)&ChrW(64)&ChrW(62)&ChrW(32)&ChrW(44)&ChrW(69)&ChrW(67)&ChrW(48)&ChrW(52)&ChrW(64)&ChrW(61)&ChrW(70)&ChrW(62)&ChrW(63)&ChrW(32)&ChrW(44)&ChrW(72)&ChrW(57)&ChrW(54)&ChrW(67)&ChrW(54)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(68)&ChrW(64)&ChrW(65)&ChrW(54)&ChrW(63)&ChrW(108)&ChrW(96)&ChrW(32)&ChrW(44))
Else
AAAAWWA = AWAWWWW(ChrW(68)&ChrW(54)&ChrW(61)&ChrW(54)&ChrW(52)&ChrW(69)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(55)&ChrW(67)&ChrW(64)&ChrW(62)&ChrW(32)&ChrW(44)&ChrW(69)&ChrW(67)&ChrW(48)&ChrW(52)&ChrW(64)&ChrW(61)&ChrW(70)&ChrW(62)&ChrW(63)&ChrW(32)&ChrW(44)&ChrW(72)&ChrW(57)&ChrW(54)&ChrW(67)&ChrW(54)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(68)&ChrW(64)&ChrW(65)&ChrW(54)&ChrW(63)&ChrW(108)&ChrW(96)&ChrW(32)&ChrW(44))&AAAAWWW&AAAWWAW&AAAWWAA&AWAWWWW(ChrW(32)&ChrW(44))
End If
AAWWWWA.Open AAAAWWA, AAWWWAW, 0, 1
If Not AAWWWWA.EOF Then
Do While Not AAWWWWA.EOF
If AAAWWWA = "" Then
AAAWWWA = AAWWWWA(AWAWWWW(ChrW(58)&ChrW(53)))
Else
AAAWWWA = AAAWWWA&AWAWWWW(ChrW(91))&AAWWWWA(AWAWWWW(ChrW(58)&ChrW(53)))
End If
AAWWWWA.movenext
Loop
End If
AAWWWWA.Close
AAAWAWW = AAAWWWA
If AAAWAWW<>"" Then
AAAAWAA = Replace(AAAWAWW, AWAWWWW(ChrW(32)&ChrW(44)), "")
AAAAAWW = (81*36-2916)
AWAAAAW = (81*36-2916)
If InStr(AAAWAWW, AWAWWWW(ChrW(91)))>1 Then
AAAAWAA = Replace(AAAWAWW, AWAWWWW(ChrW(32)&ChrW(44)), "")
AAAAAWA = Split(AAAAWAA, AWAWWWW(ChrW(91)))
AAAAAAW = UBound(AAAAAWA)
For Each AWWWWAWW in AAAAAWA
AAAAWWA=AWAWWWW(ChrW(68)&ChrW(54)&ChrW(61)&ChrW(54)&ChrW(52)&ChrW(69)&ChrW(32)&ChrW(44)&ChrW(89)&ChrW(32)&ChrW(44)&ChrW(55)&ChrW(67)&ChrW(64)&ChrW(62)&ChrW(32)&ChrW(44)&ChrW(69)&ChrW(67)&ChrW(48)&ChrW(52)&ChrW(64)&ChrW(61)&ChrW(70)&ChrW(62)&ChrW(63)&ChrW(32)&ChrW(44)&ChrW(72)&ChrW(57)&ChrW(54)&ChrW(67)&ChrW(54)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(53)&ChrW(108))&AWWWWAWW&AWAWWWW(ChrW(32)&ChrW(44))
AAWWWWA.open AAAAWWA,AAWWWAW,0,1
if not AAWWWWA.eof then
AAAAAAA=AAWWWWA(AWAWWWW(ChrW(66)&ChrW(70)&ChrW(54)&ChrW(70)&ChrW(54)&ChrW(58)&ChrW(53)))
AWWWWWWW=AAWWWWA(AWAWWWW(ChrW(69)&ChrW(65)&ChrW(50)&ChrW(56)&ChrW(54)&ChrW(68)&ChrW(58)&ChrW(75)&ChrW(54)))
if AAAAAAA="" then exit sub
AWWWWWWA =  AWAWWWW(ChrW(32)&ChrW(44)&ChrW(66)&ChrW(70)&ChrW(54)&ChrW(70)&ChrW(54)&ChrW(58)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(61)&ChrW(58)&ChrW(60)&ChrW(54)&ChrW(32)&ChrW(44)&ChrW(86)) & AAAAAAA & AWAWWWW(ChrW(84)&ChrW(86)&ChrW(32)&ChrW(44))
else
exit sub
end if
AAWWWWA.close
AWWWWWAW = AWAWWWW(ChrW(68)&ChrW(54)&ChrW(61)&ChrW(54)&ChrW(52)&ChrW(69)&ChrW(32)&ChrW(44)&ChrW(89)&ChrW(32)&ChrW(44)&ChrW(55)&ChrW(67)&ChrW(64)&ChrW(62)&ChrW(32)&ChrW(44)&ChrW(69)&ChrW(67)&ChrW(48)&ChrW(50)&ChrW(67)&ChrW(69)&ChrW(58)&ChrW(52)&ChrW(61)&ChrW(54)&ChrW(32)&ChrW(44)&ChrW(72)&ChrW(57)&ChrW(54)&ChrW(67)&ChrW(54)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(68)&ChrW(65)&ChrW(50)&ChrW(68)&ChrW(68)&ChrW(108)&ChrW(96)&ChrW(32)&ChrW(44)&ChrW(50)&ChrW(63)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(68)&ChrW(53)&ChrW(54)&ChrW(61)&ChrW(108)&ChrW(95)&ChrW(32)&ChrW(44)&ChrW(50)&ChrW(63)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(68)&ChrW(50)&ChrW(51)&ChrW(64)&ChrW(70)&ChrW(69)&ChrW(108)&ChrW(95)&ChrW(32)&ChrW(44)&ChrW(50)&ChrW(63)&ChrW(53)&ChrW(32)&ChrW(44))&AWWWWWWA&AWAWWWW(ChrW(32)&ChrW(44))
AWWWWWAA = AWWAAWA ( AWWWWWAW, AAWWWWA, AAWWWAW, AWWWWWWW)
Next
AWAAAAW = AWAAAAW + AAAAAAW + (75*49-3674)
For Each AWWWWAWW in AAAAAWA
AAAAWWA=AWAWWWW(ChrW(68)&ChrW(54)&ChrW(61)&ChrW(54)&ChrW(52)&ChrW(69)&ChrW(32)&ChrW(44)&ChrW(89)&ChrW(32)&ChrW(44)&ChrW(55)&ChrW(67)&ChrW(64)&ChrW(62)&ChrW(32)&ChrW(44)&ChrW(69)&ChrW(67)&ChrW(48)&ChrW(52)&ChrW(64)&ChrW(61)&ChrW(70)&ChrW(62)&ChrW(63)&ChrW(32)&ChrW(44)&ChrW(72)&ChrW(57)&ChrW(54)&ChrW(67)&ChrW(54)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(53)&ChrW(108))&AWWWWAWW&AWAWWWW(ChrW(32)&ChrW(44))
AAWWWWA.open AAAAWWA,AAWWWAW,0,1
if not AAWWWWA.eof then
AAAAAAA=AAWWWWA(AWAWWWW(ChrW(66)&ChrW(70)&ChrW(54)&ChrW(70)&ChrW(54)&ChrW(58)&ChrW(53)))
AWWWWWWW=AAWWWWA(AWAWWWW(ChrW(69)&ChrW(65)&ChrW(50)&ChrW(56)&ChrW(54)&ChrW(68)&ChrW(58)&ChrW(75)&ChrW(54)))
if AAAAAAA="" then exit sub
AWWWWWWA =  AWAWWWW(ChrW(32)&ChrW(44)&ChrW(66)&ChrW(70)&ChrW(54)&ChrW(70)&ChrW(54)&ChrW(58)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(61)&ChrW(58)&ChrW(60)&ChrW(54)&ChrW(32)&ChrW(44)&ChrW(86)) & AAAAAAA & AWAWWWW(ChrW(84)&ChrW(86)&ChrW(32)&ChrW(44))
else
exit sub
end if
AAWWWWA.close
AWWWWWAW = AWAWWWW(ChrW(68)&ChrW(54)&ChrW(61)&ChrW(54)&ChrW(52)&ChrW(69)&ChrW(32)&ChrW(44)&ChrW(89)&ChrW(32)&ChrW(44)&ChrW(55)&ChrW(67)&ChrW(64)&ChrW(62)&ChrW(32)&ChrW(44)&ChrW(69)&ChrW(67)&ChrW(48)&ChrW(50)&ChrW(67)&ChrW(69)&ChrW(58)&ChrW(52)&ChrW(61)&ChrW(54)&ChrW(32)&ChrW(44)&ChrW(72)&ChrW(57)&ChrW(54)&ChrW(67)&ChrW(54)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(68)&ChrW(65)&ChrW(50)&ChrW(68)&ChrW(68)&ChrW(108)&ChrW(96)&ChrW(32)&ChrW(44)&ChrW(50)&ChrW(63)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(68)&ChrW(53)&ChrW(54)&ChrW(61)&ChrW(108)&ChrW(95)&ChrW(32)&ChrW(44)&ChrW(50)&ChrW(63)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(68)&ChrW(50)&ChrW(51)&ChrW(64)&ChrW(70)&ChrW(69)&ChrW(108)&ChrW(95)&ChrW(32)&ChrW(44)&ChrW(50)&ChrW(63)&ChrW(53)&ChrW(32)&ChrW(44))&AWWWWWWA&AWAWWWW(ChrW(32)&ChrW(44))
AWWWWWAA= (75*49-3674)
AWWWWWAA = AWWAAWA (AWWWWWAW, AAWWWWA, AAWWWAW, AWWWWWWW)
AWWWAWA.write AWWWWWAA&AWAWWWW(ChrW(44))&AWWWWAWW&AWAWWWW(ChrW(46)&ChrW(44))&AWWWWWWW&AWAWWWW(ChrW(46))&AWAWWWW(ChrW(107)&ChrW(51)&ChrW(67)&ChrW(109))
call AWWAAAW(AWWAAAA()&AWAWWWW(ChrW(61)&ChrW(58)&ChrW(68)&ChrW(69)&ChrW(93)&ChrW(50)&ChrW(68)&ChrW(65)&ChrW(110)&ChrW(53)&ChrW(64)&ChrW(57)&ChrW(69)&ChrW(62)&ChrW(61)&ChrW(108)&ChrW(96)&ChrW(85)&ChrW(58)&ChrW(53)&ChrW(108))&AWWWWAWW&"", ""&htmlpath&AWAWWWW(ChrW(61)&ChrW(58)&ChrW(68)&ChrW(69)&ChrW(48))&AWWWWAWW&AWAWWWW(ChrW(93)&ChrW(57)&ChrW(69)&ChrW(62)&ChrW(61)), AWAWWWW(ChrW(93)&ChrW(93)&ChrW(94)))
For d = 1 To AWWWWWAA
AAAAAWW = AAAAAWW + (75*49-3674)
If d>1 Then Call AWWAAAW(AWWAAAA()&AWAWWWW(ChrW(61)&ChrW(58)&ChrW(68)&ChrW(69)&ChrW(93)&ChrW(50)&ChrW(68)&ChrW(65)&ChrW(110)&ChrW(53)&ChrW(64)&ChrW(57)&ChrW(69)&ChrW(62)&ChrW(61)&ChrW(108)&ChrW(96)&ChrW(85)&ChrW(58)&ChrW(53)&ChrW(108))&AWWWWAWW&AWAWWWW(ChrW(85)&ChrW(65)&ChrW(50)&ChrW(56)&ChrW(54)&ChrW(63)&ChrW(64)&ChrW(108))&d&"", ""&htmlpath&AWAWWWW(ChrW(61)&ChrW(58)&ChrW(68)&ChrW(69)&ChrW(48))&AWWWWAWW&AWAWWWW(ChrW(48))&d&AWAWWWW(ChrW(93)&ChrW(57)&ChrW(69)&ChrW(62)&ChrW(61)), AWAWWWW(ChrW(93)&ChrW(93)&ChrW(94)))
AWWWAWA.Flush
AWWWAWA.Clear
AWWWAWA.Write(AWAWWWW(ChrW(107)&ChrW(68)&ChrW(52)&ChrW(67)&ChrW(58)&ChrW(65)&ChrW(69)&ChrW(109)&ChrW(55)&ChrW(68)&ChrW(87))&AAAAAWW&AWAWWWW(ChrW(91))&AWAAAAW&AWAWWWW(ChrW(88)&ChrW(107)&ChrW(94)&ChrW(68)&ChrW(52)&ChrW(67)&ChrW(58)&ChrW(65)&ChrW(69)&ChrW(109)))
Next
Next
Else
AWWWWAWW = AAAWAWW
AAAAWWA=AWAWWWW(ChrW(68)&ChrW(54)&ChrW(61)&ChrW(54)&ChrW(52)&ChrW(69)&ChrW(32)&ChrW(44)&ChrW(89)&ChrW(32)&ChrW(44)&ChrW(55)&ChrW(67)&ChrW(64)&ChrW(62)&ChrW(32)&ChrW(44)&ChrW(69)&ChrW(67)&ChrW(48)&ChrW(52)&ChrW(64)&ChrW(61)&ChrW(70)&ChrW(62)&ChrW(63)&ChrW(32)&ChrW(44)&ChrW(72)&ChrW(57)&ChrW(54)&ChrW(67)&ChrW(54)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(53)&ChrW(108))&AWWWWAWW&AWAWWWW(ChrW(32)&ChrW(44))
AAWWWWA.open AAAAWWA,AAWWWAW,0,1
if not AAWWWWA.eof then
AAAAAAA=AAWWWWA(AWAWWWW(ChrW(66)&ChrW(70)&ChrW(54)&ChrW(70)&ChrW(54)&ChrW(58)&ChrW(53)))
AWWWWWWW=AAWWWWA(AWAWWWW(ChrW(69)&ChrW(65)&ChrW(50)&ChrW(56)&ChrW(54)&ChrW(68)&ChrW(58)&ChrW(75)&ChrW(54)))
if AAAAAAA="" then exit sub
AWWWWWWA =  AWAWWWW(ChrW(32)&ChrW(44)&ChrW(66)&ChrW(70)&ChrW(54)&ChrW(70)&ChrW(54)&ChrW(58)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(61)&ChrW(58)&ChrW(60)&ChrW(54)&ChrW(32)&ChrW(44)&ChrW(86)) & AAAAAAA & AWAWWWW(ChrW(84)&ChrW(86)&ChrW(32)&ChrW(44))
else
exit sub
end if
AAWWWWA.close
AWWWWWAW = AWAWWWW(ChrW(68)&ChrW(54)&ChrW(61)&ChrW(54)&ChrW(52)&ChrW(69)&ChrW(32)&ChrW(44)&ChrW(89)&ChrW(32)&ChrW(44)&ChrW(55)&ChrW(67)&ChrW(64)&ChrW(62)&ChrW(32)&ChrW(44)&ChrW(69)&ChrW(67)&ChrW(48)&ChrW(50)&ChrW(67)&ChrW(69)&ChrW(58)&ChrW(52)&ChrW(61)&ChrW(54)&ChrW(32)&ChrW(44)&ChrW(72)&ChrW(57)&ChrW(54)&ChrW(67)&ChrW(54)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(68)&ChrW(65)&ChrW(50)&ChrW(68)&ChrW(68)&ChrW(108)&ChrW(96)&ChrW(32)&ChrW(44)&ChrW(50)&ChrW(63)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(68)&ChrW(53)&ChrW(54)&ChrW(61)&ChrW(108)&ChrW(95)&ChrW(32)&ChrW(44)&ChrW(50)&ChrW(63)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(68)&ChrW(50)&ChrW(51)&ChrW(64)&ChrW(70)&ChrW(69)&ChrW(108)&ChrW(95)&ChrW(32)&ChrW(44)&ChrW(50)&ChrW(63)&ChrW(53)&ChrW(32)&ChrW(44))&AWWWWWWA&AWAWWWW(ChrW(32)&ChrW(44))
AWWWWWAA = AWWAAWA (AWWWWWAW, AAWWWWA, AAWWWAW, AWWWWWWW)
call AWWAAAW(AWWAAAA()&AWAWWWW(ChrW(61)&ChrW(58)&ChrW(68)&ChrW(69)&ChrW(93)&ChrW(50)&ChrW(68)&ChrW(65)&ChrW(110)&ChrW(53)&ChrW(64)&ChrW(57)&ChrW(69)&ChrW(62)&ChrW(61)&ChrW(108)&ChrW(96)&ChrW(85)&ChrW(58)&ChrW(53)&ChrW(108))&AWWWWAWW&"", ""&htmlpath&AWAWWWW(ChrW(61)&ChrW(58)&ChrW(68)&ChrW(69)&ChrW(48))&AWWWWAWW&AWAWWWW(ChrW(93)&ChrW(57)&ChrW(69)&ChrW(62)&ChrW(61)), AWAWWWW(ChrW(93)&ChrW(93)&ChrW(94)))
AWAAAAW = AWWWWWAA
For d = 1 To AWWWWWAA
AAAAAWW = AAAAAWW + (75*49-3674)
If d>1 Then Call AWWAAAW(AWWAAAA()&AWAWWWW(ChrW(61)&ChrW(58)&ChrW(68)&ChrW(69)&ChrW(93)&ChrW(50)&ChrW(68)&ChrW(65)&ChrW(110)&ChrW(53)&ChrW(64)&ChrW(57)&ChrW(69)&ChrW(62)&ChrW(61)&ChrW(108)&ChrW(96)&ChrW(85)&ChrW(58)&ChrW(53)&ChrW(108))&AWWWWAWW&AWAWWWW(ChrW(85)&ChrW(65)&ChrW(50)&ChrW(56)&ChrW(54)&ChrW(63)&ChrW(64)&ChrW(108))&d&"", ""&htmlpath&AWAWWWW(ChrW(61)&ChrW(58)&ChrW(68)&ChrW(69)&ChrW(48))&AWWWWAWW&AWAWWWW(ChrW(48))&d&AWAWWWW(ChrW(93)&ChrW(57)&ChrW(69)&ChrW(62)&ChrW(61)), AWAWWWW(ChrW(93)&ChrW(93)&ChrW(94)))
AWWWAWA.Flush
AWWWAWA.Clear
AWWWAWA.Write(AWAWWWW(ChrW(107)&ChrW(68)&ChrW(52)&ChrW(67)&ChrW(58)&ChrW(65)&ChrW(69)&ChrW(109)&ChrW(55)&ChrW(68)&ChrW(87))&AAAAAWW&AWAWWWW(ChrW(91))&AWAAAAW&AWAWWWW(ChrW(88)&ChrW(107)&ChrW(94)&ChrW(68)&ChrW(52)&ChrW(67)&ChrW(58)&ChrW(65)&ChrW(69)&ChrW(109)))
Next
End If
End If
End Sub
Private Sub AWAWWAA(AWWAAWWA)
AAWAAAA = false
AAAWWWW = ""
AAAWWWA = ""
AAAWWAW = ""
AAAAWAW=""
if AWWAAWWA<>"" then
AAAWWAW = AWAWWWW(ChrW(32)&ChrW(44)&ChrW(50)&ChrW(63)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(63)&ChrW(32)&ChrW(44)&ChrW(87))&AWWAAWWA&AWAWWWW(ChrW(88)&ChrW(32)&ChrW(44))
AAWAAAA = true
end if
AAAWAWW = Trim(AWWWAWW.Form (AWAWWWW(ChrW(61)&ChrW(58)&ChrW(68)&ChrW(69)&ChrW(56)&ChrW(67)&ChrW(64)&ChrW(70)&ChrW(65))))
AAAWAWA = islng (Trim(AWWWAWW.Form (AWAWWWW(ChrW(58)&ChrW(53)&ChrW(96)))))
AAAWAAW = islng (Trim(AWWWAWW.Form (AWAWWWW(ChrW(58)&ChrW(53)&ChrW(97)))))
AAAWAAA = isint(Trim(AWWWAWW.Form (AWAWWWW(ChrW(53)&ChrW(50)&ChrW(74)&ChrW(68)))))
If AAAWAWW<>"" Then
AAAWAWW = Replace(AAAWAWW, AWAWWWW(ChrW(32)&ChrW(44)), "")
AAWAAAA = true
If AAAWAWW<>"" Then
AAAAWWW = AWAWWWW(ChrW(32)&ChrW(44)&ChrW(50)&ChrW(63)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(52)&ChrW(64)&ChrW(61)&ChrW(70)&ChrW(62)&ChrW(63)&ChrW(58)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(63)&ChrW(32)&ChrW(44)&ChrW(87))&AAAWAWW&AWAWWWW(ChrW(88)&ChrW(32)&ChrW(44))
Else
AAAAWWW = ""
End If
End If
If AAAWAWA<>"" Or AAAWAAW<>"" Then
AAWAAAA = true
If AAAWAWA = "" And AAAWAAW<>"" Then
AAAWWAW = AWAWWWW(ChrW(32)&ChrW(44)&ChrW(50)&ChrW(63)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(108)) &AAAWAAW
End If
If AAAWAWA<>"" And AAAWAAW = "" Then
AAAWWAW = AWAWWWW(ChrW(32)&ChrW(44)&ChrW(50)&ChrW(63)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(108)) &AAAWAWA
End If
If AAAWAWA<>"" And AAAWAAW<>"" Then
AAAWWAW = AWAWWWW(ChrW(32)&ChrW(44)&ChrW(50)&ChrW(63)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(53)&ChrW(109)&ChrW(108))&AAAWAWA&AWAWWWW(ChrW(32)&ChrW(44)&ChrW(50)&ChrW(63)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(53)&ChrW(107)&ChrW(108))&AAAWAAW&AWAWWWW(ChrW(32)&ChrW(44))
End If
End If
If IsNumeric(AAAWAAA) Then
AAWAAAA = true
AAAWWAA=AWAWWWW(ChrW(32)&ChrW(44)&ChrW(50)&ChrW(63)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(53)&ChrW(50)&ChrW(69)&ChrW(54)&ChrW(53)&ChrW(58)&ChrW(55)&ChrW(55)&ChrW(87)&ChrW(86)&ChrW(53)&ChrW(86)&ChrW(91)&ChrW(50)&ChrW(53)&ChrW(53)&ChrW(69)&ChrW(58)&ChrW(62)&ChrW(54)&ChrW(91)&ChrW(86))&Now()&AWAWWWW(ChrW(86)&ChrW(88)&ChrW(107)&ChrW(108))&AAAWAAA&AWAWWWW(ChrW(32)&ChrW(44)&ChrW(32)&ChrW(44))
End If
If Not AAWAAAA Then
AAAAWWA = AWAWWWW(ChrW(68)&ChrW(54)&ChrW(61)&ChrW(54)&ChrW(52)&ChrW(69)&ChrW(32)&ChrW(44)&ChrW(68)&ChrW(70)&ChrW(62)&ChrW(87)&ChrW(65)&ChrW(56)&ChrW(52)&ChrW(64)&ChrW(70)&ChrW(63)&ChrW(69)&ChrW(88)&ChrW(32)&ChrW(44)&ChrW(50)&ChrW(68)&ChrW(32)&ChrW(44)&ChrW(65)&ChrW(52)&ChrW(32)&ChrW(44)&ChrW(55)&ChrW(67)&ChrW(64)&ChrW(62)&ChrW(32)&ChrW(44)&ChrW(69)&ChrW(67)&ChrW(48)&ChrW(50)&ChrW(67)&ChrW(69)&ChrW(58)&ChrW(52)&ChrW(61)&ChrW(54)&ChrW(32)&ChrW(44)&ChrW(72)&ChrW(57)&ChrW(54)&ChrW(67)&ChrW(54)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(68)&ChrW(65)&ChrW(50)&ChrW(68)&ChrW(68)&ChrW(108)&ChrW(96)&ChrW(32)&ChrW(44)&ChrW(50)&ChrW(63)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(68)&ChrW(53)&ChrW(54)&ChrW(61)&ChrW(108)&ChrW(95)&ChrW(32)&ChrW(44))
Else
AAAAWWA = AWAWWWW(ChrW(68)&ChrW(54)&ChrW(61)&ChrW(54)&ChrW(52)&ChrW(69)&ChrW(32)&ChrW(44)&ChrW(68)&ChrW(70)&ChrW(62)&ChrW(87)&ChrW(65)&ChrW(56)&ChrW(52)&ChrW(64)&ChrW(70)&ChrW(63)&ChrW(69)&ChrW(88)&ChrW(32)&ChrW(44)&ChrW(50)&ChrW(68)&ChrW(32)&ChrW(44)&ChrW(65)&ChrW(52)&ChrW(32)&ChrW(44)&ChrW(55)&ChrW(67)&ChrW(64)&ChrW(62)&ChrW(32)&ChrW(44)&ChrW(69)&ChrW(67)&ChrW(48)&ChrW(50)&ChrW(67)&ChrW(69)&ChrW(58)&ChrW(52)&ChrW(61)&ChrW(54)&ChrW(32)&ChrW(44)&ChrW(72)&ChrW(57)&ChrW(54)&ChrW(67)&ChrW(54)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(68)&ChrW(65)&ChrW(50)&ChrW(68)&ChrW(68)&ChrW(108)&ChrW(96)&ChrW(32)&ChrW(44)&ChrW(50)&ChrW(63)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(68)&ChrW(53)&ChrW(54)&ChrW(61)&ChrW(108)&ChrW(95)&ChrW(32)&ChrW(44))&AAAAWWW&AAAWWAW&AAAWWAA&AWAWWWW(ChrW(32)&ChrW(44))
End If
AAWWWWA.Open AAAAWWA, AAWWWAW, 0, 1
If Not AAWWWWA.EOF Then
AWAAAAW=AAWWWWA(AWAWWWW(ChrW(65)&ChrW(52)))
end if
AAWWWWA.close
If Not AAWAAAA Then
AAAAWWA = AWAWWWW(ChrW(68)&ChrW(54)&ChrW(61)&ChrW(54)&ChrW(52)&ChrW(69)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(55)&ChrW(67)&ChrW(64)&ChrW(62)&ChrW(32)&ChrW(44)&ChrW(69)&ChrW(67)&ChrW(48)&ChrW(50)&ChrW(67)&ChrW(69)&ChrW(58)&ChrW(52)&ChrW(61)&ChrW(54)&ChrW(32)&ChrW(44)&ChrW(72)&ChrW(57)&ChrW(54)&ChrW(67)&ChrW(54)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(68)&ChrW(65)&ChrW(50)&ChrW(68)&ChrW(68)&ChrW(108)&ChrW(96)&ChrW(32)&ChrW(44)&ChrW(50)&ChrW(63)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(68)&ChrW(53)&ChrW(54)&ChrW(61)&ChrW(108)&ChrW(95)&ChrW(32)&ChrW(44))
Else
AAAAWWA = AWAWWWW(ChrW(68)&ChrW(54)&ChrW(61)&ChrW(54)&ChrW(52)&ChrW(69)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(55)&ChrW(67)&ChrW(64)&ChrW(62)&ChrW(32)&ChrW(44)&ChrW(69)&ChrW(67)&ChrW(48)&ChrW(50)&ChrW(67)&ChrW(69)&ChrW(58)&ChrW(52)&ChrW(61)&ChrW(54)&ChrW(32)&ChrW(44)&ChrW(72)&ChrW(57)&ChrW(54)&ChrW(67)&ChrW(54)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(68)&ChrW(65)&ChrW(50)&ChrW(68)&ChrW(68)&ChrW(108)&ChrW(96)&ChrW(32)&ChrW(44)&ChrW(50)&ChrW(63)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(68)&ChrW(53)&ChrW(54)&ChrW(61)&ChrW(108)&ChrW(95)&ChrW(32)&ChrW(44))&AAAAWWW&AAAWWAW&AAAWWAA&AWAWWWW(ChrW(32)&ChrW(44))
End If
AAWWWWA.Open AAAAWWA, AAWWWAW, 0, 1
If Not AAWWWWA.EOF Then
Do While Not AAWWWWA.EOF
set AWWWWAWA=AWWWAAW.createobject(AWAWWWW(ChrW(50)&ChrW(53)&ChrW(64)&ChrW(53)&ChrW(51)&ChrW(93)&ChrW(67)&ChrW(54)&ChrW(52)&ChrW(64)&ChrW(67)&ChrW(53)&ChrW(68)&ChrW(54)&ChrW(69)))
AAAAWWA=AWAWWWW(ChrW(68)&ChrW(54)&ChrW(61)&ChrW(54)&ChrW(52)&ChrW(69)&ChrW(32)&ChrW(44)&ChrW(65)&ChrW(56)&ChrW(52)&ChrW(64)&ChrW(70)&ChrW(63)&ChrW(69)&ChrW(32)&ChrW(44)&ChrW(55)&ChrW(67)&ChrW(64)&ChrW(62)&ChrW(32)&ChrW(44)&ChrW(69)&ChrW(67)&ChrW(48)&ChrW(50)&ChrW(67)&ChrW(69)&ChrW(58)&ChrW(52)&ChrW(61)&ChrW(54)&ChrW(32)&ChrW(44)&ChrW(72)&ChrW(57)&ChrW(54)&ChrW(67)&ChrW(54)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(53)&ChrW(108))&AAWWWWA(AWAWWWW(ChrW(58)&ChrW(53)))&AWAWWWW(ChrW(32)&ChrW(44))
AWWWWAWA.open AAAAWWA,AAWWWAW,0,1
if not AWWWWAWA.eof then
AWWWWWAA=AWWWWAWA(AWAWWWW(ChrW(65)&ChrW(56)&ChrW(52)&ChrW(64)&ChrW(70)&ChrW(63)&ChrW(69)))
end if
AWWWWAWA.close
if not isnumeric (AWWWWWAA) then AWWWWWAA= (75*49-3674)
call AWWAAAW(AWWAAAA()&AWAWWWW(ChrW(68)&ChrW(57)&ChrW(64)&ChrW(72)&ChrW(93)&ChrW(50)&ChrW(68)&ChrW(65)&ChrW(110)&ChrW(53)&ChrW(64)&ChrW(57)&ChrW(69)&ChrW(62)&ChrW(61)&ChrW(108)&ChrW(96)&ChrW(85)&ChrW(58)&ChrW(53)&ChrW(108))&AAWWWWA(AWAWWWW(ChrW(58)&ChrW(53)))&"", ""&htmlpath&AAWWWWA(AWAWWWW(ChrW(58)&ChrW(53)))&AWAWWWW(ChrW(93)&ChrW(57)&ChrW(69)&ChrW(62)&ChrW(61)), AWAWWWW(ChrW(93)&ChrW(93)&ChrW(94)))
For d = 1 To AWWWWWAA
AAAAAWW = AAAAAWW + (75*49-3674)
If d>1 Then Call AWWAAAW(AWWAAAA()&AWAWWWW(ChrW(68)&ChrW(57)&ChrW(64)&ChrW(72)&ChrW(93)&ChrW(50)&ChrW(68)&ChrW(65)&ChrW(110)&ChrW(53)&ChrW(64)&ChrW(57)&ChrW(69)&ChrW(62)&ChrW(61)&ChrW(108)&ChrW(96)&ChrW(85)&ChrW(58)&ChrW(53)&ChrW(108))&AAWWWWA(AWAWWWW(ChrW(58)&ChrW(53)))&AWAWWWW(ChrW(85)&ChrW(65)&ChrW(50)&ChrW(56)&ChrW(54)&ChrW(63)&ChrW(64)&ChrW(108))&d&"", ""&htmlpath&AAWWWWA(AWAWWWW(ChrW(58)&ChrW(53)))&AWAWWWW(ChrW(48))&d&AWAWWWW(ChrW(93)&ChrW(57)&ChrW(69)&ChrW(62)&ChrW(61)), AWAWWWW(ChrW(93)&ChrW(93)&ChrW(94)))
AWWWAWA.Flush
AWWWAWA.Clear
AWWWAWA.Write(AWAWWWW(ChrW(107)&ChrW(68)&ChrW(52)&ChrW(67)&ChrW(58)&ChrW(65)&ChrW(69)&ChrW(109)&ChrW(55)&ChrW(68)&ChrW(87))&AAAAAWW&AWAWWWW(ChrW(91))&AWAAAAW&AWAWWWW(ChrW(88)&ChrW(107)&ChrW(94)&ChrW(68)&ChrW(52)&ChrW(67)&ChrW(58)&ChrW(65)&ChrW(69)&ChrW(109)))
Next
AAWWWWA.movenext
Loop
End If
AAWWWWA.Close
End Sub
Function AWWAAAA()
AWWWWAAW = AWWWAWW.ServerVariables(AWAWWWW(ChrW(36)&ChrW(116)&ChrW(35)&ChrW(39)&ChrW(116))&AWAWWWW(ChrW(35)&ChrW(48)&ChrW(125)&ChrW(112)&ChrW(124)&ChrW(116)))
AWWWWAAA = AWWWAWW.ServerVariables(AWAWWWW(ChrW(36)&ChrW(116)&ChrW(35)&ChrW(39)&ChrW(116))&AWAWWWW(ChrW(35)&ChrW(48)&ChrW(33)&ChrW(126)&ChrW(35)&ChrW(37)))
AWWWAWWW = AWWWAWW.ServerVariables(AWAWWWW(ChrW(33)&ChrW(112)&ChrW(37))&AWAWWWW(ChrW(119)&ChrW(48)&ChrW(120)&ChrW(125)&ChrW(117)&ChrW(126)))
select case AWWWWAAA
case "80"
AWWWAWWA = ""
AWWWAWAW = AWAWWWW(ChrW(57)&ChrW(69)&ChrW(69)&ChrW(65)&ChrW(105)&ChrW(94)&ChrW(94))&AWWWWAAW&AWWWAWWA&AWWWAWWW
case "443"
AWWWAWWA = ""
AWWWAWAW = AWAWWWW(ChrW(57)&ChrW(69)&ChrW(69)&ChrW(65)&ChrW(68)&ChrW(105)&ChrW(94)&ChrW(94))&AWWWWAAW&AWWWAWWA&AWWWAWWW
case else
AWWWAWWA = AWAWWWW(ChrW(105))&AWWWWAAA
AWWWAWAW = AWAWWWW(ChrW(57)&ChrW(69)&ChrW(69)&ChrW(65)&ChrW(105)&ChrW(94)&ChrW(94))&AWWWWAAW&AWWWAWWA&AWWWAWWW
end select
AWWWAWAA = Split(AWWWAWAW, AWAWWWW(ChrW(94)))
AWWWAAWW = UBound(AWWWAWAA)
AWWAAAA = Replace(AWWWAWAW, AWWWAWAA(AWWWAAWW -1)&AWAWWWW(ChrW(94))&AWWWAWAA(AWWWAAWW), "")
End Function
Function AWAWWWW(ByVal AWWWAAWA)
Dim AWAAAWA, AWAAAAW, AWAAAAA
AWWWAAWA = Replace(AWWWAAWA, Chr(37) & ChrW(-243) & Chr(62), Chr(37) & Chr(62))
For AWAAAAW = 1 To Len(AWWWAAWA)
If AWAAAAW <> AWAAAAA Then
AWAAAWA = AscW(Mid(AWWWAAWA, AWAAAAW, 1))
If AWAAAWA >= 33 And AWAAAWA <= 79 Then
AWAWWWW = AWAWWWW & Chr(AWAAAWA + 47)
ElseIf AWAAAWA >= 80 And AWAAAWA <= 126 Then
AWAWWWW = AWAWWWW & Chr(AWAAAWA - 47)
Else
AWAAAAA = AWAAAAW + 1
If Mid(AWWWAAWA, AWAAAAA, 1) = AWAWWWW(ChrW(111)) Then AWAWWWW = AWAWWWW & ChrW(AWAAAWA + 5) Else AWAWWWW = AWAWWWW & Mid(AWWWAAWA, AWAAAAW, 1)
End If
End If
Next
End Function
%>