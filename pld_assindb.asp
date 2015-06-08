<!--#include file="Inc/Sysconn.asp" -->
<!--#include file="inc/Skin_CSS.asp"-->
<!--#include file="Inc/function.asp"-->
<%
	Dim action
	action=Request("action")

	Dim strlsh, strTsrwlx, strtsnr, stryxtsy, strcc, strrksj,strbz, strTsNum, j
	Dim strtsks, strtsjs
	strlsh=Request("TsLsh")		'流水号
	strTsrwlx=Request("Tsrwlx") '任务类型--厂内，厂外任务
	strtsnr=Request("tsnr")		'调试内容
	stryxtsy=Request("yxtsy")	'已选调试员
	strzrr=Request("zrr")		'厂外调试员
	strTsNum=Csng(Request("TsNum"))	'厂外调试参与调试员人数
	strcc=Request("ccfs")		'出厂方式\厂外验收方式
	strbz=Request("beiz")		'厂外调试验收备注
	strrksj=Request("rksj")		'入库时间
	errmsg=""					'错误提示

	Select case action
		Case "Massign"		'厂外调试分配
			Call Tsindb()
		Case "YL"			'用料
			Call YLindb()
		Case "Ruku"			'入库
			Call Rukuindb()
		Case "Chuuku"			'入库
			Call Chuukuindb()
		Case "Change"		'更改
			Call Changeindb()
		Case "ExtF"			'厂外调试分值入库
			Call ExtFindb()

	End Select

Function Tsindb()
	If strTsrwlx="1" Then
		'厂内调试任务
		Dim strcycs
		strcycs=Request("cycs")
		If 	strcycs="" Then
			errmsg="参与调试次数不能为空！"
			Call WriteErrMsg()
			Response.End
		End If
		If stryxtsy="" Then
			'不增加调试次数直接结束任务
			strSql="select * from [ins_tsxx] where lsh='"&strlsh&"' and cc='正在调试' order by ID Desc"
			set Rs=server.createobject("adodb.recordset")
			Rs.open strSql,Conn,1,3
			If Rs.eof Or Rs.bof Then
				errmsg="任务还没开始或已经验收结束!"
				Call WriteErrMsg()
				Response.End
			else
					Rs("cc")=strcc
					Rs.update
			End If
		else
			'添加厂内调试任务调试员
			strSql="select * from [ins_tsxx]"
			set Rs=server.createobject("adodb.recordset")
			Rs.open strSql,Conn,1,3
			for i=1 to CSng(strcycs)
				Rs.AddNew
				Rs("lsh")=strlsh
				Rs("tsy")=stryxtsy
				Rs("tsnr")=strtsnr
				if CSng(strcycs)>1 and i<CSng(strcycs) Then
					Rs("cc")="正在调试"
				else
					Rs("cc")=strcc
				End if
				Rs.update
			next
		End If
		'更新厂内任务书中的调试次数、结束时间、出厂方式
		Dim strDxcs, strTscs
		strTscs=0 : strDxcs=0	'模头、定型调试次数
		strSql="select * from [ins_tsxx] where lsh='"&strlsh&"'"
		set Rs=server.createobject("adodb.recordset")
		Rs.open strSql,Conn,1,1
		Do while not(Rs.eof Or Rs.bof)
			strTscs=strTscs+1
			strDxcs=strDxcs+1
			Rs.movenext
		Loop
		strTscs=strTscs &":"& strDxcs	'调试次数
		strSql="select * from [mission] where lsh='"&strlsh&"'"
		set Rs=server.createobject("adodb.recordset")
		Rs.open strSql,Conn,1,3
		If Not(Rs.eof Or Rs.bof) Then
				Rs("sjtscs")=strTscs
				Rs("ccfs")=strcc
				If strcc<>"正在调试" Then Rs("sjjssj")=Date() End If
				Rs.update
		End If
	else	'厂外调试任务
		'更新厂外任务书中的验收方式、备注
		Dim LshArr,strddh
		LshArr=Split(strlsh,",")
		For i=o to ubound(LshArr)
 			strSql="select * from [mission] where lsh='"&LshArr(i)&"' and isNull(cwkssj)"
 			set Rs=server.createobject("adodb.recordset")
 			Rs.open strSql,Conn,1,3
 			If Not(Rs.eof Or Rs.bof) Then
 				strddh=Rs("ddh")
 				Rs("ysjg")=strcc
 				If strbz<>"" Then
 					If IsNull(Rs("bz")) Then
 						Rs("bz")="※ "&strbz
 					else
 						Rs("bz")=Rs("bz") &"<br>※ "& strbz
 					End If
 				End If
 			Rs.update
 			else
 				errmsg=errmsg&LshArr(i)&"模具已分配，请核实后再次输入！<br>"
 			End If
 			Rs.Close
		Next
		If errmsg<>"" Then
			Call WriteErrMsg()
			Exit Function
		End If
		'更新调试信息
		strSql="select * from [ext_tsxx]"
		set Rs=server.createobject("adodb.recordset")
		Rs.open strSql,Conn,1,3
		For i=0 to ubound(LshArr)
    		Rs.AddNew
    		Rs("ddh")=strddh
    		Rs("lsh")=LshArr(i)
    		Rs("tsy")=Request("zrr")
    		Rs("kssj")=Request("tskssj")
    		Rs("jssj")=Request("tsjssj")
    		Rs("cc")=strcc
    		Rs.update
    	Next
		If strTsNum>1 Then
			for i = 1 to strTsNum-1
				For j=0 to ubound(LshArr)
					Rs.AddNew
					Rs("ddh")=strddh
					Rs("lsh")=LshArr(j)
					Rs("tsy")=Request(i&"zrr")
					Rs("kssj")=Request(i&"tskssj")
					Rs("jssj")=Request(i&"tsjssj")
					Rs("cc")=strcc
					Rs.update
				Next
			next
		End If
		Rs.Close

		'更新模具分值系数
		Dim imjxs
		For i=0 to ubound(LshArr)
			imjxs=Request("mjxs"&i)
			if imjxs="" Then imjxs=1
			strSql="select * from [mission] where lsh='"&LshArr(i)&"'"
			set Rs=server.createobject("adodb.recordset")
			Rs.open strSql,Conn,1,3
    		Rs("djxs")=imjxs
    		Rs.update
			Rs.Close
    	Next

	'结算厂外调试员得分,临时（不能被统计）
		If strcc="验收合格" or strcc="通知放行"	Then
'		strSql="select * from [ext_tsfz] where lsh='"&LshArr(i)&"'"
'		set Rs=server.createobject("adodb.recordset")
'		Rs.open strSql,Conn,1,1
'		If not(Rs.eof Or Rs.bof) Then	'防止重复计分
'			Call WriteErrMsg()
'		else
			Call ext_fz(strlsh)
'		End If
'		Rs.Close
		End If
	End If
	Call WriteSuccessMsg("流水号 " & strlsh &  " 任务书分配成功!","")
	Rs.Close
End Function

Function YLindb()
	Dim stryxtsyjs
	stryxtsyjs=Request("yxtsyjs")
	If stryxtsyjs="" Then errmsg="时间不能为空！" End If
	If errmsg <> "" Then
		Call WriteErrMsg()
		Response.End
	End If
	Dim stryl
	stryl=Request("smyl")
	'更新调试用料
	strSql="select * from [mission] where lsh='"&strlsh&"'"
	set Rs=server.createobject("adodb.recordset")
	Rs.open strSql,Conn,1,3
	If Not(Rs.eof Or Rs.bof) Then
		Rs("sjtsyl")=stryl
		Rs("rksj")=stryxtsyjs
		Rs("cksj")=stryxtsyjs
		Rs.update
	End If
'	'结算厂内调试员得分
'	strSql="select * from [ins_tsfz] where lsh='"&strlsh&"'"
'	set Rs=server.createobject("adodb.recordset")
'	Rs.open strSql,Conn,1,3
'	If Rs.eof Or Rs.bof	Then	'防止重复计分
		Call ins_fz(stryxtsyjs)
'	End If
	Call WriteSuccessMsg("流水号 " & strlsh &  " 任务书分配成功!","pld_assign.asp?rwlx=1")
	Rs.Close
End Function

Function ins_fz(stryxtsyjs)
	Dim strEdcs, strSjcs, TsGrxs, XsGrxs, strCsxs, strYlxs, strTszf, strTscs, strTsy, TsyArray
	Dim TmpSql, Rs2, ArrCs, ArrGrxs, Tmpxs, TsZgrxs, XsZgrxs, strGzz, Grxs, TmpArr, strXszf
	strTsy="" : ArrCs="" : ArrGrxs=""
	'根据用料、调试次数、和出厂方式求的总的奖励
	strSql="select * from [mission] where lsh='"&strlsh&"'"
	set Rs=server.createobject("adodb.recordset")
	Rs.open strSql,Conn,1,3
	If not(Rs.eof Or Rs.bof) Then
		strTszf=Rs("cntsfz")	'厂内调试分值
		strXszf=Rs("scxfz")		'生产线分值
		strEdcs=Split(Rs("edtscs"),":")
		strSjcs=Split(Rs("sjtscs"),":")
'		strCsxs=(strEdcs(0)-strSjcs(0))*0.02+(strEdcs(1)-strSjcs(1))*0.05	'调试次数系数
'		strYlxs=Fix((Rs("edtsyl")-Rs("sjtsyl"))/(Rs("edtsyl")*0.25))*0.05	'用料系数
		if strSjcs(0)>strEdcs(1) Then
			strCsxs=(strEdcs(1)-strSjcs(0))*0.05
		else if strSjcs(0)<strEdcs(0) Then
				strCsxs=(strEdcs(0)-strSjcs(0))*0.05
			else
				strCsxs=0
			End If
		End If
		strYlxs=0
		Select Case Rs("ccfs")
			Case "产品合格"
				Tmpxs=1+strCsxs+strYlxs
			Case "一项不合格"
				Tmpxs=1+strCsxs+strYlxs-0.1
			Case "两项不合格"
				Tmpxs=1+strCsxs+strYlxs-0.2
			Case "三项不合格"
				Tmpxs=1+strCsxs+strYlxs-0.3
			Case "通知放行"
				Tmpxs=0.6
			Case "粗调"
				Tmpxs=0.6
		End Select
		If Tmpxs<0.6 Then Tmpxs=0.6 End If
		strTszf=strTszf*Tmpxs
		strXszf=strXszf*Tmpxs
		Rs("cnsjfz")=Round(strTszf,1)
		Rs.update
	End If
	Rs.Close
'	'求出参与调试的每个人的系数
'	strSql="select * from [ins_tsxx] where lsh='"&strlsh&"' order by ID"
'	set Rs=server.createobject("adodb.recordset")
'	Rs.open strSql,Conn,1,1
'	ArrCs=Rs.recordcount
'	Do while not(Rs.eof Or Rs.bof)
'		Grxs="" : TsGrxs="" : XsGrxs="" : TsZgrxs=0 : XsZgrxs=0 : strGzz=""
'		'调试员个人系数，线上个人系数,本次调试参与调试员系数之和,本次调试参与调试线系数之和
'		strTsy=Rs("tsy")
'		strTsy=split(strTsy,", ")
'		For i = 0 to ubound(strTsy)
'			TmpSql="select * from [user] where UserName='"&strTsy(i)&"'"
'			set Rs2=server.createobject("adodb.recordset")
'			Rs2.open TmpSql,Conn,1,1
'			If not(Rs2.eof Or Rs2.bof) Then
'				Select case Rs2("Gzz")
'					case 1
'						If TsGrxs="" Then
'							TsGrxs=Rs2("Grxs")
'						else
'							TsGrxs=TsGrxs & "," & Rs2("Grxs")
'						End If
'					case 3
'						If XsGrxs="" Then
'							XsGrxs=Rs2("Grxs")
'						else
'							XsGrxs=XsGrxs & "," & Rs2("Grxs")
'						End If
'				End Select
'				If strGzz="" Then
'					strGzz=Rs2("GZZ")
'				else
'					strGzz=strGzz & "," & Rs2("GZZ")
'				End If
'				If Grxs="" Then
'					Grxs=Rs2("Grxs")
'				else
'					Grxs=Grxs & "," & Rs2("Grxs")
'				End If
'			End If
'			Rs2.Close
'		Next
'		TsGrxs=split(TsGrxs,",")
'		XsGrxs=split(XsGrxs,",")
'		strGzz=split(strGzz,",")
'		Grxs=split(Grxs,",")
'		for i = 0 to ubound(TsGrxs)
'			If TsZgrxs=0 Then
'				TsZgrxs=CSng(TsGrxs(i))
'			else
'				TsZgrxs=TsZgrxs + CSng(TsGrxs(i))
'			End If
'		next
'		for i = 0 to ubound(XsGrxs)
'			If XsZgrxs=0 Then
'				XsZgrxs=CSng(XsGrxs(i))
'			else
'				XsZgrxs=XsZgrxs + CSng(XsGrxs(i))
'			End If
'		next
'		for i = 0 to ubound(strTsy)
'			Select case strGzz(i)
'				case 1
'					If ArrGrxs="" Then
'						ArrGrxs=strTsy(i) &":"& CSng(Grxs(i))/TsZgrxs &":"& strGzz(i)
'					else
'						ArrGrxs=ArrGrxs &";"& strTsy(i) &":"& CSng(Grxs(i))/TsZgrxs &":"& strGzz(i)
'					End If
'				case 3
'					If ArrGrxs="" Then
'						ArrGrxs=strTsy(i) &":"& CSng(Grxs(i))/XsZgrxs &":"& strGzz(i)
'					else
'						ArrGrxs=ArrGrxs &";"& strTsy(i) &":"& CSng(Grxs(i))/XsZgrxs &":"& strGzz(i)
'						'组成形如"甲:0.20:1;乙:0.52:3;丙:0.31:1;甲:0.77:1"的字符串
'					End If
'			End Select
'		Next
'	Rs.movenext
'	Loop
'	Rs.Close
'	ArrGrxs=split(ArrGrxs,";")
'	For i=0 to Ubound(ArrGrxs)
'		TmpArr=split(ArrGrxs(i),":")
'		strSql="select * from [ins_tsfz] where lsh='"&strlsh&"' and zrr='"&TmpArr(0)&"'"
'		set Rs=server.createobject("adodb.recordset")
'		Rs.open strSql,Conn,1,3
'		If Rs.eof Or Rs.bof Then
'			Rs.AddNew
'			Rs("lsh")=strlsh
'			Rs("zrr")=TmpArr(0)
'			Rs("cycs")=1
'			If TmpArr(2)=1 Then
'				Rs("fz")=strTszf/ArrCs*TmpArr(1)
'				Rs("js")="调试"
'			else
'				Rs("fz")=strXszf/ArrCs*TmpArr(1)
'				Rs("js")="生产"
'			End If
'			Rs("sj")=stryxtsyjs
'			Rs.update
'		else
'			Rs("cycs")=Rs("cycs")+1
'			If TmpArr(2)=1 Then
'				Rs("fz")=Rs("fz")+strTszf/ArrCs*TmpArr(1)
'			else
'				Rs("fz")=Rs("fz")+strXszf/ArrCs*TmpArr(1)
'			End If
'			Rs("sj")=stryxtsyjs
'			Rs.update
'		End If
'	Rs.Close
'	Next
'	strSql="select * from [ins_tsfz] where lsh='"&strlsh&"' Order BY ID"
'	set rs=server.createobject("adodb.recordset")
'	rs.Open strSql,conn,1,3
'	do while not rs.eof
'		Rs("fz")=Round(Rs("fz"),1)
'		rs.movenext
'	loop
'	Rs.Close
End Function

Function Rukuindb()
	Dim strzxfz
	If strrksj="" Then errmsg="入库时间不能为空！" End If
	If errmsg <> "" Then
		Call WriteErrMsg()
		Response.End
	End If
	'读取装箱分值
	strSql="select * from [mission] where lsh='"&strlsh&"'"
	set Rs=server.createobject("adodb.recordset")
	Rs.open strSql,Conn,1,1
	If Not(Rs.eof Or Rs.bof) Then
		strzxfz=Rs("bgxfz")
	End If
	Rs.Close
	'更新入库时间
	strSql="select * from [mission] where lsh='"&strlsh&"'"
	set Rs=server.createobject("adodb.recordset")
	Rs.open strSql,Conn,1,3
	If Not(Rs.eof Or Rs.bof) Then
		Rs("rksj")=strrksj
		Rs.update
	End If
	Rs.Close
	'更新入库所得分值
	strSql="select * from [ins_tsfz] where lsh='"&strlsh&"' and zrr='装箱'"
	set Rs=server.createobject("adodb.recordset")
	Rs.open strSql,Conn,1,3
	If Rs.eof Or Rs.bof Then
		Rs.AddNew
		Rs("lsh")=strlsh
		Rs("zrr")="装箱"
		Rs("js")="装箱"
		Rs("cycs")=1
		Rs("fz")=strzxfz
		Rs("sj")=strrksj
		Rs.update
	End If
	Call WriteSuccessMsg("流水号 " & strlsh &  " 成功入库!","")
	Rs.Close
End Function

Function Chuukuindb()
	Dim strcksj
	strcksj=Request("cksj")
	If strcksj="" Then errmsg="出库时间不能为空！" End If
	If errmsg <> "" Then
		Call WriteErrMsg()
		Response.End
	End If
	'更新出库时间
	strSql="select * from [mission] where lsh='"&strlsh&"'"
	set Rs=server.createobject("adodb.recordset")
	Rs.open strSql,Conn,1,3
	If Not(Rs.eof Or Rs.bof) Then
		Rs("cksj")=strcksj
		Rs.update
	End If
	Rs.Close
	Call WriteSuccessMsg("流水号 " & strlsh &  " 成功出库!","")
End Function

Function Changeindb()
	Dim Number,tsy, strid, strtsnr, strTscs,strcs
	Number=Request("num")
	errmsg=""
	tsy=Request("zrr"&Number)
	If tsy="" Then
		errmsg="最后一次调试出厂不允许删除！"
		Call WriteErrMsg()
		exit Function
	End If
	For i=1 to Number
		strid=Request("id"&i)
		tsy=Request("zrr"&i)
		If tsy<>"" Then
			tsy=Replace(tsy,",",", ")
			'更新调试信息
			strSql="select * from [ins_tsxx] where id=" & Clng(strid)
			set Rs=server.createobject("adodb.recordset")
			Rs.open strSql,Conn,1,3
			If Not(Rs.eof Or Rs.bof) Then
				Rs("tsy")=tsy
				Rs.update
			End If
		else
			'调试员为空时删除这次调试信息
			''首先查询本次调试内容更新模具调试次数
			strSql="select * from [ins_tsxx] where id=" & Clng(strid)
			set Rs=server.createobject("adodb.recordset")
			Rs.open strSql,Conn,1,1
			If Not(Rs.eof Or Rs.bof) Then
				strtsnr=Rs("tsnr")'调试内容
				strlsh=Rs("lsh")'流水号
			End If
			strSql="select * from [mission] where lsh='"&strlsh&"'"
			set Rs=server.createobject("adodb.recordset")
			Rs.open strSql,Conn,1,1
			strTscs=Rs("sjtscs")
			strTscs=split(strTscs,":")
			If strtsnr="定型" Then'修改调试次数
				strcs= strTscs(0) &":"& Csng(strTscs(1))-1
			else
				strcs= Csng(strTscs(0))-1 &":"& strTscs(1)
			End If

			strSql="select * from [mission] where lsh='"&strlsh&"'"
			set Rs=server.createobject("adodb.recordset")
			Rs.open strSql,Conn,1,3
			If Not(Rs.eof Or Rs.bof) Then
					Rs("sjtscs")=strcs
					Rs.update
			End If

			'删除调试信息
			sql="delete from ins_tsxx where ID=" & Clng(strid)
			conn.Execute sql
		End If
		Rs.Close
	next
	Call WriteSuccessMsg("更改责任人成功更新!","")
End Function

Function ext_fz(slsh)
	Dim ilsh,strTsy,strkssj,strjssj,TmpArr,Tmpstr,TmpXs,Tmptsy,Tmpfz,j,n
	ilsh=slsh : strTsy="" :	strkssj="" : strjssj="" : TmpArr="" : Tmptsy="" : Tmpfz=0 : errmsg=""
	'------------------------------------------求厂外调试时间、参与人员
	strlsh=Split(slsh,",")
	strSql="select * from [ext_tsxx] where lsh='"&strlsh(0)&"'"
	set Rs=server.createobject("adodb.recordset")
	Rs.open strSql,Conn,1,1
	Do while not rs.eof
		If strkssj="" Then
			strkssj=Rs("kssj")		'开始调试时间
		else
			If datediff("d",strkssj,Rs("kssj"))<0 Then strkssj=Rs("kssj")
		End If
		If strjssj="" Then
			strjssj=Rs("jssj")		'结束调试时间
		else
			If datediff("d",strjssj,Rs("jssj"))>0 Then strjssj=Rs("jssj")
		End If
		If strTsy="" Then
			strTsy=Rs("tsy")		'参与调试人员
		else
			strTsy=strTsy&","&Rs("tsy")
		End If
	Rs.movenext
	loop
	Rs.Close
	strTsy=Split(MoveR(strTsy),",")	'过滤调试员列表中重复项
	'----------------------------------参与调试人员系数列表
	For i=0 to Ubound(strTsy)
		strSql="select * from [user] where UserName='"&strTsy(i)&"'"
		set Rs=server.createobject("adodb.recordset")
		Rs.open strSql,Conn,1,1
		If not rs.eof Then
			strTsy(i)=strTsy(i)&":"&Rs("Grxs")
		End If
		Rs.Close
	Next
	'-----------------------------------求厂外调试分值,更新结束不统计时间
	strtsfz=0
	For i=0 to ubound(strlsh)
		strSql="select * from [mission] where lsh='"&strlsh(i)&"'"
		set Rs=server.createobject("adodb.recordset")
		Rs.open strSql,Conn,1,3
		If Not(Rs.eof Or Rs.bof) Then
			strtsfz=strtsfz + Rs("cwtsfz")*Rs("djxs")	'厂外调试分值
			Rs("cwkssj")=strkssj						'更新开始时间
			Rs("btjsj")=strjssj	 						'更新结束不统计时间
			Rs.update
		End If
		Rs.Close
	Next
	'-----------------------------------调试员得分
	Dim TmpKs,TmpJs
	TmpKs="" : TmpJs=""
'	errmsg="共"&datediff("d",strkssj,strjssj)&"天,strtsfz="&strtsfz&"<br>"
	For i=0 to datediff("d",strkssj,strjssj)
		TmpArr=""
		Tmpstr=dateadd("d",i,strkssj)
		strSql="select * from [ext_tsxx] where lsh='"&strlsh(0)&"'"
		set Rs=server.createobject("adodb.recordset")
		Rs.open strSql,Conn,1,1
		Do while not Rs.eof
			If datediff("d",Tmpstr,Rs("kssj"))<=0 and datediff("d",Tmpstr,Rs("jssj"))>=0 Then
				If TmpArr="" Then
					For each x in strTsy
						If Left(x,Instr(x,":")-1)=Rs("tsy") Then TmpArr=x
					Next
				else				'得出具体某一天的参与调试人员
					For each x in strTsy
						If Left(x,Instr(x,":")-1)=Rs("tsy") Then TmpArr=TmpArr&"||"&x
					Next
				End If
			End If
			Rs.movenext
		Loop
		Rs.Close
		If TmpArr<>"" Then
    		TmpArr=Split(TmpArr,"||")
    		TmpXs=0
    		For j=0 to ubound(TmpArr)
    			TmpXs=TmpXs+Mid(TmpArr(j),Instr(TmpArr(j),":")+1)
    		Next
    		For j=0 to ubound(TmpArr)
    			Tmptsy=Left(TmpArr(j),Instr(TmpArr(j),":")-1)
    			Tmpfz=Round(strtsfz/(datediff("d",strkssj,strjssj)+1)*Mid(TmpArr(j),Instr(TmpArr(j),":")+1)/TmpXs,3)
'    			errmsg=errmsg&Tmpstr&"这一天:"&Tmptsy&",系数:"&TmpXs&",分值:"&Tmpfz&"<br>"
				strSql="select * from [ext_tsxx] where lsh='"&ilsh&"' and tsy='"&Tmptsy&"' order by ID Desc"
    			set Rs=server.createobject("adodb.recordset")
    			Rs.open strSql,Conn,1,1
    			If Not Rs.eof Then
    				TmpKs=Rs("kssj")
    				TmpJs=Rs("jssj")
    			End If
    			Rs.Close

    			strSql="select * from [ext_tsfz] where (lsh='"&ilsh&"' or lsh like '%,"&ilsh&"%' or lsh like '%"&ilsh&",%')and zrr='"&Tmptsy&"'"
    			set Rs=server.createobject("adodb.recordset")
    			Rs.open strSql,Conn,1,3
    			If Rs.eof Or Rs.bof Then
    				Rs.AddNew
    				Rs("lsh")=ilsh
    				Rs("zrr")=Tmptsy
    				If TmpKs<>"" Then Rs("kssj")=TmpKs
    				If TmpJs<>"" Then Rs("jssj")=TmpJs
    				Rs("fz")=Tmpfz
    				Rs.update
    			else
    				Rs("fz")=Rs("fz")+Tmpfz
    				Rs.update
    			End If
    			Rs.Close
    		next
		End If
	next
'	Call WriteErrMsg()
	Call WriteSuccessMsg("流水号 "&ilsh&" 厂外任务分配成功!","")
End Function

Function ExtFindb()		'打完系数后重新更新分值并写入时间，使能被统计
	Dim strnumb,strid,strxs,strfz
	strnumb=Request("ExtFNB")
	For i=0 to strnumb
		strid=Request("id"&i)
		strxs=Request("xs"&i)
		If strxs="" Then strxs=0
		strfz=Request("fz"&i)
		strSql="select * from [ext_tsfz] where id=" & Clng(strid)
		set Rs=server.createobject("adodb.recordset")
		Rs.open strSql,Conn,1,3
		If Not(Rs.eof Or Rs.bof) Then
			Rs("xs")=strxs
			Rs("fz")=strfz
			Rs("rksj")=Date()
			Rs.update
		End If
		Rs.Close
	Next
	'更新任务实际结束时间
	Dim ilsh
	ilsh=Split(strlsh,",")
	For i=0 to ubound(ilsh)
		strSql="select * from [mission] where lsh='"&ilsh(i)&"'"
		set Rs=server.createobject("adodb.recordset")
		Rs.open strSql,Conn,1,3
		If Not(Rs.eof Or Rs.bof) Then
			Rs("cwjssj")=Rs("btjsj")
			Rs("btjsj")=null
			Rs.update
		End If
		Rs.Close
	Next
	Call WriteSuccessMsg("流水号 "&strlsh&" 厂外任务分配成功!","")
End Function

'******************************
'函数：MoveR(Rstr)
'参数：Rstr,以逗号做分隔的字符串
'描述：去除数组中重复的项
'示例：MoveR("abc,abc,dge,gcg,dge,dgi,die,dir,fgk,dir,gis,sgier,ssir")
'******************************
Function MoveR(Rstr)
 Dim i,SpStr
 SpStr = Split(Rstr,",")
 For i = 0 To Ubound(Spstr)
  If i = 0 then
   MoveR = MoveR & SpStr(i)
  Else
   If instr(MoveR,SpStr(i))=0 Then
    MoveR = MoveR & "," & SpStr(i)
   End If
  End If
 Next
End Function
%>
