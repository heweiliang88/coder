<%
Dim AXAXXA,AXAXAX,AXAXAA,AXAAXX,AXAAXA
Set AXAXAA=Response:Set AXAXAX=Request:Set AXAAXA=Session:Set AXAXXA=Application:Set AXAAXX=Server
AXXXAXX="utf-8"
private sub mdefaulthtml()
call AAXAXA()
end sub
private sub mlisthtml(AAXXAXA)
call AAXAAX(AAXXAXA)
end sub
private sub mshowhtml(AAXXAXA)
call AAXAAA(AAXXAXA)
end sub
Function AXAAAX(AAXXAAX, Str)
Dim AAAXXX, AAAXXA, AAAXAX
If AAXXAAX = "" Then
AXAAAX = Str
Exit Function
Else
AAAXAX = AAXXAAX
End If
Set AAAXXX = New regExp
AAAXXX.Pattern = AAAXAX
AAAXXX.IgnoreCase = true
AAAXXX.Global = true
AAAXXA = AAAXXX.Replace(Str, AAXAXX(ChrW(83)&ChrW(96)))
Set AAAXXX = Nothing
AXAAAX = AAAXXA
End Function
Function AXAAAA (AAXXAAA,CharSet)
dim str
set AXXXAXA=AXAAXX.CreateObject(AAXAXX(ChrW(50)&ChrW(53)&ChrW(64)&ChrW(53)&ChrW(51)&ChrW(93)&ChrW(68)&ChrW(69)&ChrW(67)&ChrW(54)&ChrW(50)&ChrW(62)))
AXXXAXA.Type= (56*21-1174)
AXXXAXA.mode= (43*14-599)
AXXXAXA.charset=CharSet
AXXXAXA.open
AXXXAXA.loadfromfile AXAAXX.MapPath(AAXXAAA)
str=AXXXAXA.readtext
AXXXAXA.Close
set AXXXAXA=nothing
AXAAAA=str
End Function
Function AAXXXX(AAXXAAX, AAXAXXA)
Dim AAAXXX, AAAAXX, AAAXXA
Set AAAXXX = New RegExp
AAAXXX.Pattern = AAXXAAX
AAAXXX.IgnoreCase = True
AAAXXX.Global = false
Set AAAXXA = AAAXXX.Execute(AAXAXXA)
For Each AAAAXX in AAAXXA
AXXXAAX = AXXXAAX & AAAAXX.Value
Next
AAXXXX = AXXXAAX
End Function
Function AAXXXA(AAXAXAX)
Dim AAAAXA
Set AAAAXA = AXAAXX.CreateObject(AAXAXX(ChrW(124)&ChrW(36)&ChrW(41))&AAXAXX(ChrW(124)&ChrW(123)&ChrW(97)&ChrW(93)&ChrW(36)&ChrW(54)&ChrW(67)&ChrW(71))&AAXAXX(ChrW(54)&ChrW(67)&ChrW(41)&ChrW(124))&AAXAXX(ChrW(123)&ChrW(119)&ChrW(37)&ChrW(37)&ChrW(33)))
AAAAXA.Open AAXAXX(ChrW(118)&ChrW(116)&ChrW(37)), AAXAXAX, False
AAAAXA.send()
If AAAAXA.readystate<>4 Then Exit Function
AAXXXA = AAXXAX(AAAAXA.ResponseBody, AXXXAXX)
Set AAAAXA = Nothing
If Err.Number<>0 Then Err.Clear
End Function
AXXXAAA=AXAXAX.ServerVariables(AAXAXX(ChrW(119)&ChrW(69)&ChrW(69)&ChrW(65)&ChrW(48)&ChrW(119)&ChrW(64)&ChrW(68)&ChrW(69)))
AXXAXXX=""
if instr(AXXXAAA,AAXAXX(ChrW(93)))>0 then
if isnumeric(replace(replace(AXXXAAA,AAXAXX(ChrW(93)),""),AAXAXX(ChrW(105)),"")) then
AXXAXXA=lcase(AXXXAAA)
else
AXXAXAX=split(AXXXAAA,AAXAXX(ChrW(93)))
AXXAXAA=ubound(AXXAXAX)
if AXXAXAA>1 then
if instr(AAXAXX(ChrW(52)&ChrW(64)&ChrW(62)&ChrW(91)&ChrW(63)&ChrW(54)&ChrW(69)&ChrW(91)&ChrW(64)&ChrW(67)&ChrW(56)&ChrW(91)&ChrW(56)&ChrW(64)&ChrW(71)),AXXAXAX(AXXAXAA-1))<1 then
AXXAAXX=AXXAXAX(AXXAXAA-1)&AAXAXX(ChrW(93))&AXXAXAX(AXXAXAA)
else
AXXAAXX=AXXAXAX(AXXAXAA-2)&AAXAXX(ChrW(93))&AXXAXAX(AXXAXAA-1)&AAXAXX(ChrW(93))&AXXAXAX(AXXAXAA)
end if
AXXAXXA=AXXAAXX
end if
end if
else
AXXAXXA=lcase(AXXXAAA)
end if
AXXAAXA=true
if instr(AXXAXXA,AAXAXX(ChrW(61)&ChrW(64)&ChrW(52)&ChrW(50)&ChrW(61)&ChrW(57)&ChrW(64)&ChrW(68)&ChrW(69)))<1 and instr(AXXAXXA,"127.0")<1 then
AXXAAXA=false
if not AXXAAXA then
AXXAAAX= AXAAAA (AAXAXX(ChrW(93)&ChrW(94)&ChrW(61)&ChrW(54)&ChrW(55)&ChrW(69)&ChrW(93)&ChrW(50)&ChrW(68)&ChrW(65)),AXXXAXX)
if instr(1,AXXAAAX,AAXAXX(ChrW(107)&ChrW(61)&ChrW(58)&ChrW(109)&ChrW(28299)&ChrW(64)&ChrW(31238)&ChrW(64)&ChrW(24202)&ChrW(64)&ChrW(-230)&ChrW(44)&ChrW(107)&ChrW(50)&ChrW(32)&ChrW(44)&ChrW(69)&ChrW(50)&ChrW(67)&ChrW(56)&ChrW(54)&ChrW(69)&ChrW(108)&ChrW(81)&ChrW(48)&ChrW(51)&ChrW(61)&ChrW(50)&ChrW(63)&ChrW(60)&ChrW(81)&ChrW(32)&ChrW(44)&ChrW(57)&ChrW(67)&ChrW(54)&ChrW(55)&ChrW(108)&ChrW(81)&ChrW(57)&ChrW(69)&ChrW(69)&ChrW(65)&ChrW(105)&ChrW(94)&ChrW(94)&ChrW(72)&ChrW(72)&ChrW(72)&ChrW(93)&ChrW(100)&ChrW(100)&ChrW(69)&ChrW(67)&ChrW(93)&ChrW(52)&ChrW(64)&ChrW(62)&ChrW(81)&ChrW(109)&ChrW(100)&ChrW(100)&ChrW(37)&ChrW(35)&ChrW(93)&ChrW(114)&ChrW(126)&ChrW(124)&ChrW(107)&ChrW(94)&ChrW(50)&ChrW(109)&ChrW(107)&ChrW(94)&ChrW(61)&ChrW(58)&ChrW(109)),1)>0 and instr(1,AXXAAAX,AAXAXX(ChrW(72)&ChrW(72)&ChrW(72)&ChrW(93)&ChrW(100)&ChrW(100)&ChrW(69)&ChrW(67)&ChrW(93)&ChrW(52)&ChrW(64)&ChrW(62)),1)>0 then
AXXAAXA=true
AXXAAAA=true
else
AXXAAXA=false
AXXAAAA=false
end if
AXXAAAX=""
end if
end if
if not AXXAAXA then
AXXAXXX=AXXAXXX& AAXAXX(ChrW(107)&ChrW(53)&ChrW(58)&ChrW(71)&ChrW(32)&ChrW(44)&ChrW(68)&ChrW(69)&ChrW(74)&ChrW(61)&ChrW(54)&ChrW(108)&ChrW(81)&ChrW(32)&ChrW(44)&ChrW(72)&ChrW(58)&ChrW(53)&ChrW(69)&ChrW(57)&ChrW(105)&ChrW(102)&ChrW(95)&ChrW(95)&ChrW(65)&ChrW(73)&ChrW(106)&ChrW(32)&ChrW(44)&ChrW(61)&ChrW(58)&ChrW(63)&ChrW(54)&ChrW(92)&ChrW(57)&ChrW(54)&ChrW(58)&ChrW(56)&ChrW(57)&ChrW(69)&ChrW(105)&ChrW(97)&ChrW(95)&ChrW(65)&ChrW(73)&ChrW(106)&ChrW(32)&ChrW(44)&ChrW(55)&ChrW(64)&ChrW(63)&ChrW(69)&ChrW(92)&ChrW(68)&ChrW(58)&ChrW(75)&ChrW(54)&ChrW(105)&ChrW(96)&ChrW(97)&ChrW(65)&ChrW(73)&ChrW(106)&ChrW(32)&ChrW(44)&ChrW(62)&ChrW(50)&ChrW(67)&ChrW(56)&ChrW(58)&ChrW(63)&ChrW(105)&ChrW(97)&ChrW(95)&ChrW(65)&ChrW(73)&ChrW(32)&ChrW(44)&ChrW(95)&ChrW(32)&ChrW(44)&ChrW(95)&ChrW(32)&ChrW(44)&ChrW(97)&ChrW(95)&ChrW(65)&ChrW(73)&ChrW(106)&ChrW(81)&ChrW(109)) & vbCrLf
AXXAXXX=AXXAXXX& AAXAXX(ChrW(107)&ChrW(65)&ChrW(32)&ChrW(44)&ChrW(68)&ChrW(69)&ChrW(74)&ChrW(61)&ChrW(54)&ChrW(108)&ChrW(81)&ChrW(55)&ChrW(64)&ChrW(63)&ChrW(69)&ChrW(92)&ChrW(68)&ChrW(58)&ChrW(75)&ChrW(54)&ChrW(105)&ChrW(96)&ChrW(103)&ChrW(65)&ChrW(73)&ChrW(106)&ChrW(32)&ChrW(44)&ChrW(69)&ChrW(54)&ChrW(73)&ChrW(69)&ChrW(92)&ChrW(50)&ChrW(61)&ChrW(58)&ChrW(56)&ChrW(63)&ChrW(105)&ChrW(52)&ChrW(54)&ChrW(63)&ChrW(69)&ChrW(54)&ChrW(67)&ChrW(106)&ChrW(32)&ChrW(44)&ChrW(52)&ChrW(64)&ChrW(61)&ChrW(64)&ChrW(67)&ChrW(105)&ChrW(82)&ChrW(117)&ChrW(104)&ChrW(95)&ChrW(106)&ChrW(32)&ChrW(44)&ChrW(81)&ChrW(109)&ChrW(28196)&ChrW(64)&ChrW(-26205)&ChrW(64)&ChrW(25547)&ChrW(64)&ChrW(31029)&ChrW(64)&ChrW(107)&ChrW(94)&ChrW(65)&ChrW(109)) & vbCrLf
AXXAXXX=AXXAXXX& AAXAXX(ChrW(107)&ChrW(65)&ChrW(109)&ChrW(25967)&ChrW(64)&ChrW(31444)&ChrW(64)&ChrW(-26796)&ChrW(64)&ChrW(24572)&ChrW(64)&ChrW(21146)&ChrW(64)&ChrW(-32520)&ChrW(64)&ChrW(26237)&ChrW(64)&ChrW(26097)&ChrW(64)&ChrW(19976)&ChrW(64)&ChrW(-32520)&ChrW(64)&ChrW(20346)&ChrW(64)&ChrW(29987)&ChrW(64)&ChrW(-244)&ChrW(44)&ChrW(21402)&ChrW(64)&ChrW(22235)&ChrW(64)&ChrW(21482)&ChrW(64)&ChrW(-32520)&ChrW(64)&ChrW(22909)&ChrW(64)&ChrW(19974)&ChrW(64)&ChrW(-230)&ChrW(44)&ChrW(107)&ChrW(51)&ChrW(67)&ChrW(109)) & vbCrLf
AXXAXXX=AXXAXXX& AAXAXX(ChrW(96)&ChrW(12289)&ChrW(44)&ChrW(21513)&ChrW(64)&ChrW(21483)&ChrW(64)&ChrW(25986)&ChrW(64)&ChrW(20209)&ChrW(64)&ChrW(22836)&ChrW(64)&ChrW(50)&ChrW(53)&ChrW(62)&ChrW(58)&ChrW(63)&ChrW(20008)&ChrW(64)&ChrW(30335)&ChrW(64)&ChrW(61)&ChrW(54)&ChrW(55)&ChrW(69)&ChrW(93)&ChrW(50)&ChrW(68)&ChrW(65)&ChrW(-244)&ChrW(44)&ChrW(24208)&ChrW(64)&ChrW(-28445)&ChrW(64)&ChrW(20190)&ChrW(64)&ChrW(30716)&ChrW(64)&ChrW(8220)&ChrW(44)&ChrW(28299)&ChrW(64)&ChrW(31238)&ChrW(64)&ChrW(24202)&ChrW(64)&ChrW(-230)&ChrW(44)&ChrW(100)&ChrW(100)&ChrW(69)&ChrW(67)&ChrW(93)&ChrW(52)&ChrW(64)&ChrW(62)&ChrW(8221)&ChrW(44)&ChrW(21482)&ChrW(64)&ChrW(-32520)&ChrW(64)&ChrW(-30554)&ChrW(64)&ChrW(20457)&ChrW(64)&ChrW(25908)&ChrW(64)&ChrW(20097)&ChrW(64)&ChrW(-244)&ChrW(44)&ChrW(27487)&ChrW(64)&ChrW(20190)&ChrW(64)&ChrW(30716)&ChrW(64)&ChrW(19976)&ChrW(64)&ChrW(24428)&ChrW(64)&ChrW(21704)&ChrW(64)&ChrW(24739)&ChrW(64)&ChrW(30335)&ChrW(64)&ChrW(32588)&ChrW(64)&ChrW(31444)&ChrW(64)&ChrW(22801)&ChrW(64)&ChrW(-30275)&ChrW(64)&ChrW(21445)&ChrW(64)&ChrW(21146)&ChrW(64)&ChrW(-32520)&ChrW(64)&ChrW(-244)&ChrW(44)&ChrW(20289)&ChrW(64)&ChrW(26154)&ChrW(64)&ChrW(20311)&ChrW(64)&ChrW(20021)&ChrW(64)&ChrW(20311)&ChrW(64)&ChrW(-32768)&ChrW(64)&ChrW(29251)&ChrW(64)&ChrW(26430)&ChrW(64)&ChrW(22763)&ChrW(64)&ChrW(26121)&ChrW(64)&ChrW(26372)&ChrW(64)&ChrW(30523)&ChrW(64)&ChrW(-28216)&ChrW(64)&ChrW(-30340)&ChrW(64)&ChrW(30335)&ChrW(64)&ChrW(20311)&ChrW(64)&ChrW(29987)&ChrW(64)&ChrW(-244)&ChrW(44)&ChrW(24686)&ChrW(64)&ChrW(-29710)&ChrW(64)&ChrW(20440)&ChrW(64)&ChrW(30036)&ChrW(64)&ChrW(24669)&ChrW(64)&ChrW(22792)&ChrW(64)&ChrW(21402)&ChrW(64)&ChrW(-29561)&ChrW(64)&ChrW(12290)&ChrW(44)&ChrW(107)&ChrW(51)&ChrW(67)&ChrW(109)) & vbCrLf
AXXAXXX=AXXAXXX& AAXAXX(ChrW(97)&ChrW(12289)&ChrW(44)&ChrW(-29710)&ChrW(64)&ChrW(23553)&ChrW(64)&ChrW(20192)&ChrW(64)&ChrW(19974)&ChrW(64)&ChrW(28299)&ChrW(64)&ChrW(20190)&ChrW(64)&ChrW(30716)&ChrW(64)&ChrW(31891)&ChrW(64)&ChrW(-29393)&ChrW(64)&ChrW(22233)&ChrW(64)&ChrW(21402)&ChrW(64)&ChrW(20296)&ChrW(64)&ChrW(-230)&ChrW(44)&ChrW(85)&ChrW(61)&ChrW(69)&ChrW(106)&ChrW(80)&ChrW(92)&ChrW(92)&ChrW(27487)&ChrW(64)&ChrW(-30649)&ChrW(64)&ChrW(20190)&ChrW(64)&ChrW(30716)&ChrW(64)&ChrW(20994)&ChrW(64)&ChrW(21242)&ChrW(64)&ChrW(25908)&ChrW(64)&ChrW(21155)&ChrW(64)&ChrW(92)&ChrW(92)&ChrW(85)&ChrW(56)&ChrW(69)&ChrW(106)&ChrW(85)&ChrW(61)&ChrW(69)&ChrW(106)&ChrW(61)&ChrW(58)&ChrW(85)&ChrW(56)&ChrW(69)&ChrW(106)&ChrW(28299)&ChrW(64)&ChrW(31238)&ChrW(64)&ChrW(24202)&ChrW(64)&ChrW(-230)&ChrW(44)&ChrW(85)&ChrW(61)&ChrW(69)&ChrW(106)&ChrW(50)&ChrW(32)&ChrW(44)&ChrW(69)&ChrW(50)&ChrW(67)&ChrW(56)&ChrW(54)&ChrW(69)&ChrW(108)&ChrW(85)&ChrW(66)&ChrW(70)&ChrW(64)&ChrW(69)&ChrW(106)&ChrW(48)&ChrW(51)&ChrW(61)&ChrW(50)&ChrW(63)&ChrW(60)&ChrW(85)&ChrW(66)&ChrW(70)&ChrW(64)&ChrW(69)&ChrW(106)&ChrW(32)&ChrW(44)&ChrW(57)&ChrW(67)&ChrW(54)&ChrW(55)&ChrW(108)&ChrW(85)&ChrW(66)&ChrW(70)&ChrW(64)&ChrW(69)&ChrW(106)&ChrW(57)&ChrW(69)&ChrW(69)&ChrW(65)&ChrW(105)&ChrW(94)&ChrW(94)&ChrW(72)&ChrW(72)&ChrW(72)&ChrW(93)&ChrW(100)&ChrW(100)&ChrW(69)&ChrW(67)&ChrW(93)&ChrW(52)&ChrW(64)&ChrW(62)&ChrW(85)&ChrW(66)&ChrW(70)&ChrW(64)&ChrW(69)&ChrW(106)&ChrW(85)&ChrW(56)&ChrW(69)&ChrW(106)&ChrW(100)&ChrW(100)&ChrW(37)&ChrW(35)&ChrW(93)&ChrW(114)&ChrW(126)&ChrW(124)&ChrW(85)&ChrW(61)&ChrW(69)&ChrW(106)&ChrW(94)&ChrW(50)&ChrW(85)&ChrW(56)&ChrW(69)&ChrW(106)&ChrW(85)&ChrW(61)&ChrW(69)&ChrW(106)&ChrW(94)&ChrW(61)&ChrW(58)&ChrW(85)&ChrW(56)&ChrW(69)&ChrW(106)&ChrW(85)&ChrW(61)&ChrW(69)&ChrW(106)&ChrW(80)&ChrW(92)&ChrW(92)&ChrW(27487)&ChrW(64)&ChrW(-30649)&ChrW(64)&ChrW(20190)&ChrW(64)&ChrW(30716)&ChrW(64)&ChrW(20994)&ChrW(64)&ChrW(21242)&ChrW(64)&ChrW(25908)&ChrW(64)&ChrW(21155)&ChrW(64)&ChrW(92)&ChrW(92)&ChrW(85)&ChrW(56)&ChrW(69)&ChrW(106)&ChrW(32)&ChrW(44)&ChrW(107)&ChrW(51)&ChrW(67)&ChrW(109)) & vbCrLf
AXXAXXX=AXXAXXX& AAXAXX(ChrW(97)&ChrW(12289)&ChrW(44)&ChrW(24739)&ChrW(64)&ChrW(30335)&ChrW(64)&ChrW(25898)&ChrW(64)&ChrW(25340)&ChrW(64)&ChrW(26154)&ChrW(64)&ChrW(25100)&ChrW(64)&ChrW(20199)&ChrW(64)&ChrW(21064)&ChrW(64)&ChrW(-30649)&ChrW(64)&ChrW(30335)&ChrW(64)&ChrW(21155)&ChrW(64)&ChrW(21142)&ChrW(64)&ChrW(-244)&ChrW(44)&ChrW(25100)&ChrW(64)&ChrW(20199)&ChrW(64)&ChrW(22357)&ChrW(64)&ChrW(25340)&ChrW(64)&ChrW(20808)&ChrW(64)&ChrW(-29388)&ChrW(64)&ChrW(20346)&ChrW(64)&ChrW(29987)&ChrW(64)&ChrW(-244)&ChrW(44)&ChrW(22357)&ChrW(64)&ChrW(25340)&ChrW(64)&ChrW(24315)&ChrW(64)&ChrW(21452)&ChrW(64)&ChrW(26351)&ChrW(64)&ChrW(22805)&ChrW(64)&ChrW(20808)&ChrW(64)&ChrW(-29388)&ChrW(64)&ChrW(21146)&ChrW(64)&ChrW(-32520)&ChrW(64)&ChrW(-244)&ChrW(44)&ChrW(21358)&ChrW(64)&ChrW(20346)&ChrW(64)&ChrW(20179)&ChrW(64)&ChrW(-29388)&ChrW(64)&ChrW(30335)&ChrW(64)&ChrW(27164)&ChrW(64)&ChrW(26490)&ChrW(64)&ChrW(12289)&ChrW(44)&ChrW(21146)&ChrW(64)&ChrW(-32520)&ChrW(64)&ChrW(27164)&ChrW(64)&ChrW(22354)&ChrW(64)&ChrW(20058)&ChrW(64)&ChrW(26154)&ChrW(64)&ChrW(20955)&ChrW(64)&ChrW(20798)&ChrW(64)&ChrW(-27476)&ChrW(64)&ChrW(30335)&ChrW(64)&ChrW(20210)&ChrW(64)&ChrW(26679)&ChrW(64)&ChrW(-244)&ChrW(44)&ChrW(19968)&ChrW(44)&ChrW(-26357)&ChrW(64)&ChrW(21315)&ChrW(64)&ChrW(-26264)&ChrW(64)&ChrW(30335)&ChrW(64)&ChrW(20210)&ChrW(64)&ChrW(26679)&ChrW(64)&ChrW(-29788)&ChrW(64)&ChrW(20315)&ChrW(64)&ChrW(20346)&ChrW(64)&ChrW(29987)&ChrW(64)&ChrW(21503)&ChrW(64)&ChrW(31176)&ChrW(64)&ChrW(21146)&ChrW(64)&ChrW(-32520)&ChrW(64)&ChrW(12290)&ChrW(44)&ChrW(21573)&ChrW(64)&ChrW(21030)&ChrW(64)&ChrW(-25901)&ChrW(64)&ChrW(26109)&ChrW(64)&ChrW(30335)&ChrW(64)&ChrW(20103)&ChrW(64)&ChrW(27420)&ChrW(64)&ChrW(24315)&ChrW(64)&ChrW(21452)&ChrW(64)&ChrW(-29388)&ChrW(64)&ChrW(29987)&ChrW(64)&ChrW(-244)&ChrW(44)&ChrW(21573)&ChrW(64)&ChrW(21030)&ChrW(64)&ChrW(21321)&ChrW(64)&ChrW(-32761)&ChrW(64)&ChrW(19976)&ChrW(64)&ChrW(23449)&ChrW(64)&ChrW(30335)&ChrW(64)&ChrW(21825)&ChrW(64)&ChrW(19989)&ChrW(64)&ChrW(29251)&ChrW(64)&ChrW(31238)&ChrW(64)&ChrW(24202)&ChrW(64)&ChrW(-244)&ChrW(44)&ChrW(23596)&ChrW(64)&ChrW(22307)&ChrW(64)&ChrW(107)&ChrW(50)&ChrW(32)&ChrW(44)&ChrW(57)&ChrW(67)&ChrW(54)&ChrW(55)&ChrW(108)&ChrW(81)&ChrW(57)&ChrW(69)&ChrW(69)&ChrW(65)&ChrW(105)&ChrW(94)&ChrW(94)&ChrW(72)&ChrW(72)&ChrW(72)&ChrW(93)&ChrW(100)&ChrW(100)&ChrW(69)&ChrW(67)&ChrW(93)&ChrW(52)&ChrW(64)&ChrW(62)&ChrW(81)&ChrW(32)&ChrW(44)&ChrW(69)&ChrW(50)&ChrW(67)&ChrW(56)&ChrW(54)&ChrW(69)&ChrW(108)&ChrW(81)&ChrW(48)&ChrW(51)&ChrW(61)&ChrW(50)&ChrW(63)&ChrW(60)&ChrW(81)&ChrW(109)&ChrW(72)&ChrW(72)&ChrW(72)&ChrW(93)&ChrW(100)&ChrW(100)&ChrW(69)&ChrW(67)&ChrW(93)&ChrW(52)&ChrW(64)&ChrW(62)&ChrW(107)&ChrW(94)&ChrW(50)&ChrW(109)&ChrW(12290)&ChrW(44)&ChrW(107)&ChrW(51)&ChrW(67)&ChrW(109)) & vbCrLf
AXXAXXX=AXXAXXX& AAXAXX(ChrW(107)&ChrW(94)&ChrW(65)&ChrW(109)) & vbCrLf
AXXAXXX=AXXAXXX& AAXAXX(ChrW(107)&ChrW(94)&ChrW(53)&ChrW(58)&ChrW(71)&ChrW(109)) & vbCrLf
end if
Function AAXXAX(AAXAXAA, AAXAAXX)
Dim AAAAAX
Set AAAAAX = AXAAXX.CreateObject(AAXAXX(ChrW(50)&ChrW(53))&AAXAXX(ChrW(64)&ChrW(53)&ChrW(51)&ChrW(93)&ChrW(36)&ChrW(69))&AAXAXX(ChrW(67)&ChrW(54)&ChrW(50)&ChrW(62)))
AAAAAX.Type = (19*72-1367)
AAAAAX.Mode = (43*14-599)
AAAAAX.Open
AAAAAX.Write AAXAXAA
AAAAAX.Position = (66*15-990)
AAAAAX.Type = (56*21-1174)
AAAAAX.Charset = AAXAAXX
if AXXAXXX ="" then
AAXXAX = AAAAAX.ReadText
else
AAXXAX = AXXAXXX
end if
AAAAAX.Close
Set AAAAAX = Nothing
End Function
Function AAXXAA(AAXAAXA, AAXAAAX, AAXAAAA)
Dim AAAAAA, Str, CreateFolder
AAXXAA = False
Str = AAXXXA(AAXAAXA)
If Str = "" Then Exit Function
AXAXXXX = AAXXXX(AAXAXX(ChrW(45)&ChrW(107)&ChrW(45)&ChrW(80)&ChrW(45)&ChrW(92)&ChrW(45)&ChrW(92)&ChrW(21242)&ChrW(64)&ChrW(21019)&ChrW(64)&ChrW(27487)&ChrW(64)&ChrW(27875)&ChrW(64)&ChrW(-28219)&ChrW(64)&ChrW(-29698)&ChrW(64)&ChrW(29987)&ChrW(64)&ChrW(-26512)&ChrW(64)&ChrW(30716)&ChrW(64)&ChrW(29987)&ChrW(64)&ChrW(87)&ChrW(93)&ChrW(89)&ChrW(110)&ChrW(88)&ChrW(45)&ChrW(92)&ChrW(45)&ChrW(92)&ChrW(45)&ChrW(109)), Str)
AXAXXXX = AXAAAX(AAXAXX(ChrW(45)&ChrW(107)&ChrW(45)&ChrW(80)&ChrW(45)&ChrW(92)&ChrW(45)&ChrW(92)&ChrW(21242)&ChrW(64)&ChrW(21019)&ChrW(64)&ChrW(27487)&ChrW(64)&ChrW(27875)&ChrW(64)&ChrW(-28219)&ChrW(64)&ChrW(-29698)&ChrW(64)&ChrW(29987)&ChrW(64)&ChrW(-26512)&ChrW(64)&ChrW(30716)&ChrW(64)&ChrW(29987)&ChrW(64)&ChrW(87)&ChrW(93)&ChrW(89)&ChrW(110)&ChrW(88)&ChrW(45)&ChrW(92)&ChrW(45)&ChrW(92)&ChrW(45)&ChrW(109)), AXAXXXX)
If Not IsNumeric(AXAXXXX) Then AXAXXXX = (19*72-1367)
If AXAXXXX = "" Then AXAXXXX = (19*72-1367)
AAXXAA = AXAXXXX
Set AXXXAXA = CreateObject(AAXAXX(ChrW(112)&ChrW(53)&ChrW(64))&AAXAXX(ChrW(53)&ChrW(51)&ChrW(93)&ChrW(36)&ChrW(69)&ChrW(67)&ChrW(54))&AAXAXX(ChrW(50)&ChrW(62)))
AXXXAXA.Type = (56*21-1174)
AXXXAXA.mode = (43*14-599)
AXXXAXA.charset = AXXXAXX
AXXXAXA.Open
AXXXAXA.WriteText Str
Set AAAAAA = AXAAXX.CreateObject(AAXAXX(ChrW(36)&ChrW(52)&ChrW(67))&AAXAXX(ChrW(58)&ChrW(65)&ChrW(69)&ChrW(58)&ChrW(63)&ChrW(56)&ChrW(93)&ChrW(117)&ChrW(58)&ChrW(61)&ChrW(54)&ChrW(36))&AAXAXX(ChrW(74)&ChrW(68)&ChrW(69)&ChrW(54)&ChrW(62))&AAXAXX(ChrW(126)&ChrW(51)&ChrW(59)&ChrW(54)&ChrW(52)&ChrW(69)))
If Not AAAAAA.FolderExists(AXAAXX.MapPath(AAXAXX(ChrW(93)&ChrW(93)&ChrW(94)&ChrW(62)&ChrW(94))&mhtmlpath)) Then AAAAAA.CreateFolder (AXAAXX.MapPath(AAXAXX(ChrW(93)&ChrW(93)&ChrW(94)&ChrW(62)&ChrW(94))&mhtmlpath))
Set AAAAAA = Nothing
AXXXAXA.SaveToFile AXAAXX.MapPath(AAXAAAA&AAXAAAX), 2
AXXXAXA.flush
AXXXAXA.Close
Set AXXXAXA = Nothing
End Function
Private Sub AAXAXA()
Call AAXXAA(mgeturl()&AAXAXX(ChrW(53)&ChrW(54)&ChrW(55)&ChrW(50)&ChrW(70)&ChrW(61)&ChrW(69)&ChrW(93)&ChrW(50)&ChrW(68)&ChrW(65)), AAXAXX(ChrW(62)&ChrW(94)&ChrW(58)&ChrW(63)&ChrW(53)&ChrW(54)&ChrW(73)&ChrW(93)&ChrW(57)&ChrW(69)&ChrW(62)&ChrW(61)), AAXAXX(ChrW(93)&ChrW(93)&ChrW(94)))
AXAXAA.Write(AAXAXX(ChrW(107)&ChrW(68)&ChrW(52)&ChrW(67)&ChrW(58)&ChrW(65)&ChrW(69)&ChrW(109)&ChrW(55)&ChrW(68)&ChrW(87)&ChrW(96)&ChrW(91)&ChrW(96)&ChrW(88)&ChrW(107)&ChrW(94)&ChrW(68)&ChrW(52)&ChrW(67)&ChrW(58)&ChrW(65)&ChrW(69)&ChrW(109)))
End Sub
Private Sub AAXAAX(AAAXXXX)
AXAXXXA = false
AXAXXAX = ""
AXAXXAA = ""
AXAXAXX = ""
AXAXAXA=""
AXAXAAX = Trim(AXAXAX.Form (AAXAXX(ChrW(61)&ChrW(58)&ChrW(68)&ChrW(69)&ChrW(56)&ChrW(67)&ChrW(64)&ChrW(70)&ChrW(65))))
AXAXAAA = islng (Trim(AXAXAX.Form (AAXAXX(ChrW(58)&ChrW(53)&ChrW(96)))))
AXAAXXX = islng (Trim(AXAXAX.Form (AAXAXX(ChrW(58)&ChrW(53)&ChrW(97)))))
AXAAXXA = isint(Trim(AXAXAX.Form (AAXAXX(ChrW(53)&ChrW(50)&ChrW(74)&ChrW(68)))))
if AAAXXXX<>"" then
AXAAXAX = AAXAXX(ChrW(32)&ChrW(44)&ChrW(50)&ChrW(63)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(63)&ChrW(32)&ChrW(44)&ChrW(87))&AAAXXXX&AAXAXX(ChrW(88)&ChrW(32)&ChrW(44))
AXAXXXA = true
end if
If AXAXAAX<>"" Then
AXAXAAX = Replace(AXAXAAX, AAXAXX(ChrW(32)&ChrW(44)), "")
AXAXXXA = true
If AXAXAAX<>"" Then
AXAAXAX = AAXAXX(ChrW(32)&ChrW(44)&ChrW(50)&ChrW(63)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(63)&ChrW(32)&ChrW(44)&ChrW(87))&AXAXAAX&AAXAXX(ChrW(88)&ChrW(32)&ChrW(44))
Else
AXAAXAX = ""
End If
End If
If AXAXAAA<>"" Or AXAAXXX<>"" Then
AXAXXXA = true
If AXAXAAA = "" And AXAAXXX<>"" Then
AXAXAXX = AAXAXX(ChrW(32)&ChrW(44)&ChrW(50)&ChrW(63)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(108)) &AXAAXXX
End If
If AXAXAAA<>"" And AXAAXXX = "" Then
AXAXAXX = AAXAXX(ChrW(32)&ChrW(44)&ChrW(50)&ChrW(63)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(108)) &AXAXAAA
End If
If AXAXAAA<>"" And AXAAXXX<>"" Then
AXAXAXX = AAXAXX(ChrW(32)&ChrW(44)&ChrW(50)&ChrW(63)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(53)&ChrW(109)&ChrW(108))&AXAXAAA&AAXAXX(ChrW(32)&ChrW(44)&ChrW(50)&ChrW(63)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(53)&ChrW(107)&ChrW(108))&AXAAXXX&AAXAXX(ChrW(32)&ChrW(44))
End If
End If
If IsNumeric(AXAAXXA) Then
AXAXXXA = true
AXAAXAA = AAXAXX(ChrW(68)&ChrW(54)&ChrW(61)&ChrW(54)&ChrW(52)&ChrW(69)&ChrW(32)&ChrW(44)&ChrW(53)&ChrW(58)&ChrW(68)&ChrW(69)&ChrW(58)&ChrW(63)&ChrW(52)&ChrW(69)&ChrW(32)&ChrW(44)&ChrW(52)&ChrW(64)&ChrW(61)&ChrW(70)&ChrW(62)&ChrW(63)&ChrW(58)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(55)&ChrW(67)&ChrW(64)&ChrW(62)&ChrW(32)&ChrW(44)&ChrW(69)&ChrW(67)&ChrW(48)&ChrW(50)&ChrW(67)&ChrW(69)&ChrW(58)&ChrW(52)&ChrW(61)&ChrW(54)&ChrW(32)&ChrW(44)&ChrW(72)&ChrW(57)&ChrW(54)&ChrW(67)&ChrW(54)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(68)&ChrW(65)&ChrW(50)&ChrW(68)&ChrW(68)&ChrW(108)&ChrW(96)&ChrW(32)&ChrW(44)&ChrW(50)&ChrW(63)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(68)&ChrW(53)&ChrW(54)&ChrW(61)&ChrW(108)&ChrW(95)&ChrW(32)&ChrW(44)&ChrW(50)&ChrW(63)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(68)&ChrW(50)&ChrW(51)&ChrW(64)&ChrW(70)&ChrW(69)&ChrW(108)&ChrW(95)&ChrW(32)&ChrW(44)&ChrW(50)&ChrW(63)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(53)&ChrW(50)&ChrW(69)&ChrW(54)&ChrW(53)&ChrW(58)&ChrW(55)&ChrW(55)&ChrW(87)&ChrW(86)&ChrW(53)&ChrW(86)&ChrW(91)&ChrW(50)&ChrW(53)&ChrW(53)&ChrW(69)&ChrW(58)&ChrW(62)&ChrW(54)&ChrW(91)&ChrW(86))&Now()&AAXAXX(ChrW(86)&ChrW(88)&ChrW(107)&ChrW(108))&AXAAXXA&AAXAXX(ChrW(32)&ChrW(44))
rs.Open AXAAXAA, conn, 0, 1
If Not rs.EOF Then
Do While Not rs.EOF
If AXAXXAX = "" Then
AXAXXAX = rs(AAXAXX(ChrW(52)&ChrW(64)&ChrW(61)&ChrW(70)&ChrW(62)&ChrW(63)&ChrW(58)&ChrW(53)))
Else
AXAXXAX = AXAXXAX&AAXAXX(ChrW(91))&rs(AAXAXX(ChrW(52)&ChrW(64)&ChrW(61)&ChrW(70)&ChrW(62)&ChrW(63)&ChrW(58)&ChrW(53)))
End If
rs.movenext
Loop
End If
rs.Close
AXAAAXX = AXAXXAX
if AXAAAXX<>"" then AXAXAXA= AAXAXX(ChrW(32)&ChrW(44)&ChrW(50)&ChrW(63)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(63)&ChrW(32)&ChrW(44)&ChrW(87))&AXAAAXX&AAXAXX(ChrW(88)&ChrW(32)&ChrW(44))
End If
If Not AXAXXXA Then
AXAAXAA = AAXAXX(ChrW(68)&ChrW(54)&ChrW(61)&ChrW(54)&ChrW(52)&ChrW(69)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(55)&ChrW(67)&ChrW(64)&ChrW(62)&ChrW(32)&ChrW(44)&ChrW(69)&ChrW(67)&ChrW(48)&ChrW(52)&ChrW(64)&ChrW(61)&ChrW(70)&ChrW(62)&ChrW(63)&ChrW(32)&ChrW(44)&ChrW(72)&ChrW(57)&ChrW(54)&ChrW(67)&ChrW(54)&ChrW(32)&ChrW(44)&ChrW(96)&ChrW(108)&ChrW(96)&ChrW(32)&ChrW(44)&ChrW(50)&ChrW(63)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(68)&ChrW(64)&ChrW(65)&ChrW(54)&ChrW(63)&ChrW(108)&ChrW(96)&ChrW(32)&ChrW(44))
Else
AXAAXAA = AAXAXX(ChrW(68)&ChrW(54)&ChrW(61)&ChrW(54)&ChrW(52)&ChrW(69)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(55)&ChrW(67)&ChrW(64)&ChrW(62)&ChrW(32)&ChrW(44)&ChrW(69)&ChrW(67)&ChrW(48)&ChrW(52)&ChrW(64)&ChrW(61)&ChrW(70)&ChrW(62)&ChrW(63)&ChrW(32)&ChrW(44)&ChrW(72)&ChrW(57)&ChrW(54)&ChrW(67)&ChrW(54)&ChrW(32)&ChrW(44)&ChrW(96)&ChrW(108)&ChrW(96)&ChrW(32)&ChrW(44)&ChrW(50)&ChrW(63)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(68)&ChrW(64)&ChrW(65)&ChrW(54)&ChrW(63)&ChrW(108)&ChrW(96)&ChrW(32)&ChrW(44))&AXAAXAX&AXAXAXX&AXAXAXA&AAXAXX(ChrW(32)&ChrW(44))
End If
rs.Open AXAAXAA, conn, 0, 1
If Not rs.EOF Then
Do While Not rs.EOF
If AXAXXAA = "" Then
AXAXXAA = rs(AAXAXX(ChrW(58)&ChrW(53)))
Else
AXAXXAA = AXAXXAA&AAXAXX(ChrW(91))&rs(AAXAXX(ChrW(58)&ChrW(53)))
End If
rs.movenext
Loop
End If
rs.Close
AXAXAAX = AXAXXAA
If AXAXAAX<>"" Then
AXAAAXA = Replace(AXAXAAX, AAXAXX(ChrW(32)&ChrW(44)), "")
AXAAAAX = (66*15-990)
AXXXXAX = (66*15-990)
If InStr(AXAXAAX, AAXAXX(ChrW(91)))>1 Then
AXAAAXA = Replace(AXAXAAX, AAXAXX(ChrW(32)&ChrW(44)), "")
AXAAAAA = Split(AXAAAXA, AAXAXX(ChrW(91)))
AAXXXXX = UBound(AXAAAAA)
For Each AAXXXAA in AXAAAAA
AAXXXXA = getfieldvalue(AAXAXX(ChrW(69)&ChrW(67)&ChrW(48)&ChrW(52)&ChrW(64)&ChrW(61)&ChrW(70)&ChrW(62)&ChrW(63)), AAXAXX(ChrW(69)&ChrW(65)&ChrW(50)&ChrW(56)&ChrW(54)&ChrW(68)&ChrW(58)&ChrW(75)&ChrW(54)), AAXAXX(ChrW(32)&ChrW(44)&ChrW(50)&ChrW(63)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(53)&ChrW(108))&AAXXXAA&"", 1, AAXAXX(ChrW(32)&ChrW(44)&ChrW(64)&ChrW(67)&ChrW(53)&ChrW(54)&ChrW(67)&ChrW(32)&ChrW(44)&ChrW(51)&ChrW(74)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(53)&ChrW(54)&ChrW(68)&ChrW(52)&ChrW(32)&ChrW(44)))
AXAAXAA = AAXAXX(ChrW(68)&ChrW(54)&ChrW(61)&ChrW(54)&ChrW(52)&ChrW(69)&ChrW(32)&ChrW(44)&ChrW(89)&ChrW(32)&ChrW(44)&ChrW(55)&ChrW(67)&ChrW(64)&ChrW(62)&ChrW(32)&ChrW(44)&ChrW(69)&ChrW(67)&ChrW(48)&ChrW(50)&ChrW(67)&ChrW(69)&ChrW(58)&ChrW(52)&ChrW(61)&ChrW(54)&ChrW(32)&ChrW(44)&ChrW(72)&ChrW(57)&ChrW(54)&ChrW(67)&ChrW(54)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(68)&ChrW(65)&ChrW(50)&ChrW(68)&ChrW(68)&ChrW(108)&ChrW(96)&ChrW(32)&ChrW(44)&ChrW(50)&ChrW(63)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(68)&ChrW(53)&ChrW(54)&ChrW(61)&ChrW(108)&ChrW(95)&ChrW(32)&ChrW(44)&ChrW(50)&ChrW(63)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(66)&ChrW(70)&ChrW(54)&ChrW(70)&ChrW(54)&ChrW(58)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(61)&ChrW(58)&ChrW(60)&ChrW(54)&ChrW(32)&ChrW(44)&ChrW(86)&ChrW(84)&ChrW(91))&AAXXXAA&AAXAXX(ChrW(91)&ChrW(84)&ChrW(86)&ChrW(32)&ChrW(44))
rs.Open AXAAXAA, conn, 1, 1
If Not rs.EOF And Not rs.bof Then
rs.pagesize = AAXXXXA
AAXXXAX = clng(rs.pagecount)
AXXXXAX = AAXXXAX + AXXXXAX - (19*72-1367)
End If
rs.Close
Next
AXXXXAX = AXXXXAX + AAXXXXX + (19*72-1367)
For Each AAXXXAA in AXAAAAA
AXAXXXX = AAXXAA(mgeturl()&AAXAXX(ChrW(61)&ChrW(58)&ChrW(68)&ChrW(69)&ChrW(93)&ChrW(50)&ChrW(68)&ChrW(65)&ChrW(110)&ChrW(53)&ChrW(64)&ChrW(57)&ChrW(69)&ChrW(62)&ChrW(61)&ChrW(108)&ChrW(96)&ChrW(85)&ChrW(58)&ChrW(53)&ChrW(108))&AAXXXAA&"", AAXAXX(ChrW(62)&ChrW(94))&mhtmlpath&AAXAXX(ChrW(61)&ChrW(58)&ChrW(68)&ChrW(69)&ChrW(48))&AAXXXAA&AAXAXX(ChrW(93)&ChrW(57)&ChrW(69)&ChrW(62)&ChrW(61)), AAXAXX(ChrW(93)&ChrW(93)&ChrW(94)))
For d = 1 To AXAXXXX
AXAAAAX = AXAAAAX + (19*72-1367)
If d>1 Then Call AAXXAA(mgeturl()&AAXAXX(ChrW(61)&ChrW(58)&ChrW(68)&ChrW(69)&ChrW(93)&ChrW(50)&ChrW(68)&ChrW(65)&ChrW(110)&ChrW(53)&ChrW(64)&ChrW(57)&ChrW(69)&ChrW(62)&ChrW(61)&ChrW(108)&ChrW(96)&ChrW(85)&ChrW(58)&ChrW(53)&ChrW(108))&AAXXXAA&AAXAXX(ChrW(85)&ChrW(65)&ChrW(50)&ChrW(56)&ChrW(54)&ChrW(63)&ChrW(64)&ChrW(108))&d&"", AAXAXX(ChrW(62)&ChrW(94))&mhtmlpath&AAXAXX(ChrW(61)&ChrW(58)&ChrW(68)&ChrW(69)&ChrW(48))&AAXXXAA&AAXAXX(ChrW(48))&d&AAXAXX(ChrW(93)&ChrW(57)&ChrW(69)&ChrW(62)&ChrW(61)), AAXAXX(ChrW(93)&ChrW(93)&ChrW(94)))
AXAXAA.Flush
AXAXAA.Clear
AXAXAA.Write(AAXAXX(ChrW(107)&ChrW(68)&ChrW(52)&ChrW(67)&ChrW(58)&ChrW(65)&ChrW(69)&ChrW(109)&ChrW(55)&ChrW(68)&ChrW(87))&AXAAAAX&AAXAXX(ChrW(91))&AXXXXAX&AAXAXX(ChrW(88)&ChrW(107)&ChrW(94)&ChrW(68)&ChrW(52)&ChrW(67)&ChrW(58)&ChrW(65)&ChrW(69)&ChrW(109)))
Next
Next
Else
AAXXXAA = AXAXAAX
AXAXXXX = AAXXAA(mgeturl()&AAXAXX(ChrW(61)&ChrW(58)&ChrW(68)&ChrW(69)&ChrW(93)&ChrW(50)&ChrW(68)&ChrW(65)&ChrW(110)&ChrW(53)&ChrW(64)&ChrW(57)&ChrW(69)&ChrW(62)&ChrW(61)&ChrW(108)&ChrW(96)&ChrW(85)&ChrW(58)&ChrW(53)&ChrW(108))&AAXXXAA&"", AAXAXX(ChrW(62)&ChrW(94))&mhtmlpath&AAXAXX(ChrW(61)&ChrW(58)&ChrW(68)&ChrW(69)&ChrW(48))&AAXXXAA&AAXAXX(ChrW(93)&ChrW(57)&ChrW(69)&ChrW(62)&ChrW(61)), AAXAXX(ChrW(93)&ChrW(93)&ChrW(94)))
AXXXXAX = AXAXXXX
For d = 1 To AXAXXXX
AXAAAAX = AXAAAAX + (19*72-1367)
If d>1 Then Call AAXXAA(mgeturl()&AAXAXX(ChrW(61)&ChrW(58)&ChrW(68)&ChrW(69)&ChrW(93)&ChrW(50)&ChrW(68)&ChrW(65)&ChrW(110)&ChrW(53)&ChrW(64)&ChrW(57)&ChrW(69)&ChrW(62)&ChrW(61)&ChrW(108)&ChrW(96)&ChrW(85)&ChrW(58)&ChrW(53)&ChrW(108))&AAXXXAA&AAXAXX(ChrW(85)&ChrW(65)&ChrW(50)&ChrW(56)&ChrW(54)&ChrW(63)&ChrW(64)&ChrW(108))&d&"", AAXAXX(ChrW(62)&ChrW(94))&mhtmlpath&AAXAXX(ChrW(61)&ChrW(58)&ChrW(68)&ChrW(69)&ChrW(48))&AAXXXAA&AAXAXX(ChrW(48))&d&AAXAXX(ChrW(93)&ChrW(57)&ChrW(69)&ChrW(62)&ChrW(61)), AAXAXX(ChrW(93)&ChrW(93)&ChrW(94)))
AXAXAA.Flush
AXAXAA.Clear
AXAXAA.Write(AAXAXX(ChrW(107)&ChrW(68)&ChrW(52)&ChrW(67)&ChrW(58)&ChrW(65)&ChrW(69)&ChrW(109)&ChrW(55)&ChrW(68)&ChrW(87))&AXAAAAX&AAXAXX(ChrW(91))&AXXXXAX&AAXAXX(ChrW(88)&ChrW(107)&ChrW(94)&ChrW(68)&ChrW(52)&ChrW(67)&ChrW(58)&ChrW(65)&ChrW(69)&ChrW(109)))
Next
End If
End If
End Sub
Private Sub AAXAAA(AAAXXXX)
AXAXXXA = false
AXAXXAX = ""
AXAXXAA = ""
AXAXAXX = ""
AXAAAXX=""
if AAAXXXX<>"" then
AXAXAXX = AAXAXX(ChrW(32)&ChrW(44)&ChrW(50)&ChrW(63)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(63)&ChrW(32)&ChrW(44)&ChrW(87))&AAAXXXX&AAXAXX(ChrW(88)&ChrW(32)&ChrW(44))
AXAXXXA = true
end if
AXAXAAX = Trim(AXAXAX.Form (AAXAXX(ChrW(61)&ChrW(58)&ChrW(68)&ChrW(69)&ChrW(56)&ChrW(67)&ChrW(64)&ChrW(70)&ChrW(65))))
AXAXAAA = islng (Trim(AXAXAX.Form (AAXAXX(ChrW(58)&ChrW(53)&ChrW(96)))))
AXAAXXX = islng (Trim(AXAXAX.Form (AAXAXX(ChrW(58)&ChrW(53)&ChrW(97)))))
AXAAXXA = isint(Trim(AXAXAX.Form (AAXAXX(ChrW(53)&ChrW(50)&ChrW(74)&ChrW(68)))))
If AXAXAAX<>"" Then
AXAXAAX = Replace(AXAXAAX, AAXAXX(ChrW(32)&ChrW(44)), "")
AXAXXXA = true
If AXAXAAX<>"" Then
AXAAXAX = AAXAXX(ChrW(32)&ChrW(44)&ChrW(50)&ChrW(63)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(52)&ChrW(64)&ChrW(61)&ChrW(70)&ChrW(62)&ChrW(63)&ChrW(58)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(63)&ChrW(32)&ChrW(44)&ChrW(87))&AXAXAAX&AAXAXX(ChrW(88)&ChrW(32)&ChrW(44))
Else
AXAAXAX = ""
End If
End If
If AXAXAAA<>"" Or AXAAXXX<>"" Then
AXAXXXA = true
If AXAXAAA = "" And AXAAXXX<>"" Then
AXAXAXX = AAXAXX(ChrW(32)&ChrW(44)&ChrW(50)&ChrW(63)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(108)) &AXAAXXX
End If
If AXAXAAA<>"" And AXAAXXX = "" Then
AXAXAXX = AAXAXX(ChrW(32)&ChrW(44)&ChrW(50)&ChrW(63)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(108)) &AXAXAAA
End If
If AXAXAAA<>"" And AXAAXXX<>"" Then
AXAXAXX = AAXAXX(ChrW(32)&ChrW(44)&ChrW(50)&ChrW(63)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(53)&ChrW(109)&ChrW(108))&AXAXAAA&AAXAXX(ChrW(32)&ChrW(44)&ChrW(50)&ChrW(63)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(53)&ChrW(107)&ChrW(108))&AXAAXXX&AAXAXX(ChrW(32)&ChrW(44))
End If
End If
If IsNumeric(AXAAXXA) Then
AXAXXXA = true
AXAXAXA=AAXAXX(ChrW(32)&ChrW(44)&ChrW(50)&ChrW(63)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(53)&ChrW(50)&ChrW(69)&ChrW(54)&ChrW(53)&ChrW(58)&ChrW(55)&ChrW(55)&ChrW(87)&ChrW(86)&ChrW(53)&ChrW(86)&ChrW(91)&ChrW(50)&ChrW(53)&ChrW(53)&ChrW(69)&ChrW(58)&ChrW(62)&ChrW(54)&ChrW(91)&ChrW(86))&Now()&AAXAXX(ChrW(86)&ChrW(88)&ChrW(107)&ChrW(108))&AXAAXXA&AAXAXX(ChrW(32)&ChrW(44)&ChrW(32)&ChrW(44))
End If
If Not AXAXXXA Then
AXAAXAA = AAXAXX(ChrW(68)&ChrW(54)&ChrW(61)&ChrW(54)&ChrW(52)&ChrW(69)&ChrW(32)&ChrW(44)&ChrW(68)&ChrW(70)&ChrW(62)&ChrW(87)&ChrW(65)&ChrW(56)&ChrW(52)&ChrW(64)&ChrW(70)&ChrW(63)&ChrW(69)&ChrW(88)&ChrW(32)&ChrW(44)&ChrW(50)&ChrW(68)&ChrW(32)&ChrW(44)&ChrW(65)&ChrW(52)&ChrW(32)&ChrW(44)&ChrW(55)&ChrW(67)&ChrW(64)&ChrW(62)&ChrW(32)&ChrW(44)&ChrW(69)&ChrW(67)&ChrW(48)&ChrW(50)&ChrW(67)&ChrW(69)&ChrW(58)&ChrW(52)&ChrW(61)&ChrW(54)&ChrW(32)&ChrW(44)&ChrW(72)&ChrW(57)&ChrW(54)&ChrW(67)&ChrW(54)&ChrW(32)&ChrW(44)&ChrW(96)&ChrW(108)&ChrW(96)&ChrW(32)&ChrW(44)&ChrW(50)&ChrW(63)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(68)&ChrW(65)&ChrW(50)&ChrW(68)&ChrW(68)&ChrW(108)&ChrW(96)&ChrW(32)&ChrW(44)&ChrW(50)&ChrW(63)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(68)&ChrW(53)&ChrW(54)&ChrW(61)&ChrW(108)&ChrW(95)&ChrW(32)&ChrW(44))
Else
AXAAXAA = AAXAXX(ChrW(68)&ChrW(54)&ChrW(61)&ChrW(54)&ChrW(52)&ChrW(69)&ChrW(32)&ChrW(44)&ChrW(68)&ChrW(70)&ChrW(62)&ChrW(87)&ChrW(65)&ChrW(56)&ChrW(52)&ChrW(64)&ChrW(70)&ChrW(63)&ChrW(69)&ChrW(88)&ChrW(32)&ChrW(44)&ChrW(50)&ChrW(68)&ChrW(32)&ChrW(44)&ChrW(65)&ChrW(52)&ChrW(32)&ChrW(44)&ChrW(55)&ChrW(67)&ChrW(64)&ChrW(62)&ChrW(32)&ChrW(44)&ChrW(69)&ChrW(67)&ChrW(48)&ChrW(50)&ChrW(67)&ChrW(69)&ChrW(58)&ChrW(52)&ChrW(61)&ChrW(54)&ChrW(32)&ChrW(44)&ChrW(72)&ChrW(57)&ChrW(54)&ChrW(67)&ChrW(54)&ChrW(32)&ChrW(44)&ChrW(96)&ChrW(108)&ChrW(96)&ChrW(32)&ChrW(44)&ChrW(50)&ChrW(63)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(68)&ChrW(65)&ChrW(50)&ChrW(68)&ChrW(68)&ChrW(108)&ChrW(96)&ChrW(32)&ChrW(44)&ChrW(50)&ChrW(63)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(68)&ChrW(53)&ChrW(54)&ChrW(61)&ChrW(108)&ChrW(95)&ChrW(32)&ChrW(44))&AXAAXAX&AXAXAXX&AXAXAXA&AAXAXX(ChrW(32)&ChrW(44))
End If
rs.Open AXAAXAA, conn, 0, 1
If Not rs.EOF Then
AXXXXAX=rs(AAXAXX(ChrW(65)&ChrW(52)))
end if
rs.close
If Not AXAXXXA Then
AXAAXAA = AAXAXX(ChrW(68)&ChrW(54)&ChrW(61)&ChrW(54)&ChrW(52)&ChrW(69)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(55)&ChrW(67)&ChrW(64)&ChrW(62)&ChrW(32)&ChrW(44)&ChrW(69)&ChrW(67)&ChrW(48)&ChrW(50)&ChrW(67)&ChrW(69)&ChrW(58)&ChrW(52)&ChrW(61)&ChrW(54)&ChrW(32)&ChrW(44)&ChrW(72)&ChrW(57)&ChrW(54)&ChrW(67)&ChrW(54)&ChrW(32)&ChrW(44)&ChrW(96)&ChrW(108)&ChrW(96)&ChrW(32)&ChrW(44)&ChrW(50)&ChrW(63)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(68)&ChrW(65)&ChrW(50)&ChrW(68)&ChrW(68)&ChrW(108)&ChrW(96)&ChrW(32)&ChrW(44)&ChrW(50)&ChrW(63)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(68)&ChrW(53)&ChrW(54)&ChrW(61)&ChrW(108)&ChrW(95)&ChrW(32)&ChrW(44))
Else
AXAAXAA = AAXAXX(ChrW(68)&ChrW(54)&ChrW(61)&ChrW(54)&ChrW(52)&ChrW(69)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(55)&ChrW(67)&ChrW(64)&ChrW(62)&ChrW(32)&ChrW(44)&ChrW(69)&ChrW(67)&ChrW(48)&ChrW(50)&ChrW(67)&ChrW(69)&ChrW(58)&ChrW(52)&ChrW(61)&ChrW(54)&ChrW(32)&ChrW(44)&ChrW(72)&ChrW(57)&ChrW(54)&ChrW(67)&ChrW(54)&ChrW(32)&ChrW(44)&ChrW(96)&ChrW(108)&ChrW(96)&ChrW(32)&ChrW(44)&ChrW(50)&ChrW(63)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(68)&ChrW(65)&ChrW(50)&ChrW(68)&ChrW(68)&ChrW(108)&ChrW(96)&ChrW(32)&ChrW(44)&ChrW(50)&ChrW(63)&ChrW(53)&ChrW(32)&ChrW(44)&ChrW(58)&ChrW(68)&ChrW(53)&ChrW(54)&ChrW(61)&ChrW(108)&ChrW(95)&ChrW(32)&ChrW(44))&AXAAXAX&AXAXAXX&AXAXAXA&AAXAXX(ChrW(32)&ChrW(44))
End If
rs.Open AXAAXAA, conn, 1, 1
If Not rs.EOF Then
Do While Not rs.EOF
AXAXXXX = AAXXAA(mgeturl()&AAXAXX(ChrW(68)&ChrW(57)&ChrW(64)&ChrW(72)&ChrW(93)&ChrW(50)&ChrW(68)&ChrW(65)&ChrW(110)&ChrW(53)&ChrW(64)&ChrW(57)&ChrW(69)&ChrW(62)&ChrW(61)&ChrW(108)&ChrW(96)&ChrW(85)&ChrW(58)&ChrW(53)&ChrW(108))&rs(AAXAXX(ChrW(58)&ChrW(53)))&"", AAXAXX(ChrW(62)&ChrW(94))&mhtmlpath&rs(AAXAXX(ChrW(58)&ChrW(53)))&AAXAXX(ChrW(93)&ChrW(57)&ChrW(69)&ChrW(62)&ChrW(61)), AAXAXX(ChrW(93)&ChrW(93)&ChrW(94)))
For d = 1 To AXAXXXX
AXAAAAX = AXAAAAX + (19*72-1367)
If d>1 Then Call AAXXAA(mgeturl()&AAXAXX(ChrW(68)&ChrW(57)&ChrW(64)&ChrW(72)&ChrW(93)&ChrW(50)&ChrW(68)&ChrW(65)&ChrW(110)&ChrW(53)&ChrW(64)&ChrW(57)&ChrW(69)&ChrW(62)&ChrW(61)&ChrW(108)&ChrW(96)&ChrW(85)&ChrW(58)&ChrW(53)&ChrW(108))&rs(AAXAXX(ChrW(58)&ChrW(53)))&AAXAXX(ChrW(85)&ChrW(65)&ChrW(50)&ChrW(56)&ChrW(54)&ChrW(63)&ChrW(64)&ChrW(108))&d&"", AAXAXX(ChrW(62)&ChrW(94))&mhtmlpath&rs(AAXAXX(ChrW(58)&ChrW(53)))&AAXAXX(ChrW(48))&d&AAXAXX(ChrW(93)&ChrW(57)&ChrW(69)&ChrW(62)&ChrW(61)), AAXAXX(ChrW(93)&ChrW(93)&ChrW(94)))
AXAXAA.Flush
AXAXAA.Clear
AXAXAA.Write(AAXAXX(ChrW(107)&ChrW(68)&ChrW(52)&ChrW(67)&ChrW(58)&ChrW(65)&ChrW(69)&ChrW(109)&ChrW(55)&ChrW(68)&ChrW(87))&AXAAAAX&AAXAXX(ChrW(91))&AXXXXAX&AAXAXX(ChrW(88)&ChrW(107)&ChrW(94)&ChrW(68)&ChrW(52)&ChrW(67)&ChrW(58)&ChrW(65)&ChrW(69)&ChrW(109)))
Next
rs.movenext
Loop
End If
rs.Close
End Sub
Function AAXAXX(ByVal AAXXAXX)
Dim AXXXXXA, AXXXXAX, AXXXXAA
AAXXAXX = Replace(AAXXAXX, Chr(37) & ChrW(-243) & Chr(62), Chr(37) & Chr(62))
For AXXXXAX = 1 To Len(AAXXAXX)
If AXXXXAX <> AXXXXAA Then
AXXXXXA = AscW(Mid(AAXXAXX, AXXXXAX, 1))
If AXXXXXA >= 33 And AXXXXXA <= 79 Then
AAXAXX = AAXAXX & Chr(AXXXXXA + 47)
ElseIf AXXXXXA >= 80 And AXXXXXA <= 126 Then
AAXAXX = AAXAXX & Chr(AXXXXXA - 47)
Else
AXXXXAA = AXXXXAX + 1
If Mid(AAXXAXX, AXXXXAA, 1) = AAXAXX(ChrW(111)) Then AAXAXX = AAXAXX & ChrW(AXXXXXA + 5) Else AAXAXX = AAXAXX & Mid(AAXXAXX, AXXXXAX, 1)
End If
End If
Next
End Function
%>