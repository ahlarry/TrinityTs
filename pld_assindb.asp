<!--#include file="Inc/Sysconn.asp" -->
<!--#include file="inc/Skin_CSS.asp"-->
<!--#include file="Inc/function.asp"-->
<%
	Dim action
	action=Request("action")

	Dim strlsh, strTsrwlx, strtsnr, stryxtsy, strcc, strrksj,strbz, strTsNum, j
	Dim strtsks, strtsjs
	strlsh=Request("TsLsh")		'��ˮ��
	strTsrwlx=Request("Tsrwlx") '��������--���ڣ���������
	strtsnr=Request("tsnr")		'��������
	stryxtsy=Request("yxtsy")	'��ѡ����Ա
	strzrr=Request("zrr")		'�������Ա
	strTsNum=Csng(Request("TsNum"))	'������Բ������Ա����
	strcc=Request("ccfs")		'������ʽ\�������շ�ʽ
	strbz=Request("beiz")		'����������ձ�ע
	strrksj=Request("rksj")		'���ʱ��
	errmsg=""					'������ʾ

	Select case action
		Case "Massign"		'������Է���
			Call Tsindb()
		Case "YL"			'����
			Call YLindb()
		Case "Ruku"			'���
			Call Rukuindb()
		Case "Chuuku"			'���
			Call Chuukuindb()
		Case "Change"		'����
			Call Changeindb()
		Case "ExtF"			'������Է�ֵ���
			Call ExtFindb()

	End Select

Function Tsindb()
	If strTsrwlx="1" Then
		'���ڵ�������
		Dim strcycs
		strcycs=Request("cycs")
		If 	strcycs="" Then
			errmsg="������Դ�������Ϊ�գ�"
			Call WriteErrMsg()
			Response.End
		End If
		If stryxtsy="" Then
			'�����ӵ��Դ���ֱ�ӽ�������
			strSql="select * from [ins_tsxx] where lsh='"&strlsh&"' and cc='���ڵ���' order by ID Desc"
			set Rs=server.createobject("adodb.recordset")
			Rs.open strSql,Conn,1,3
			If Rs.eof Or Rs.bof Then
				errmsg="����û��ʼ���Ѿ����ս���!"
				Call WriteErrMsg()
				Response.End
			else
					Rs("cc")=strcc
					Rs.update
			End If
		else
			'��ӳ��ڵ����������Ա
			strSql="select * from [ins_tsxx]"
			set Rs=server.createobject("adodb.recordset")
			Rs.open strSql,Conn,1,3
			for i=1 to CSng(strcycs)
				Rs.AddNew
				Rs("lsh")=strlsh
				Rs("tsy")=stryxtsy
				Rs("tsnr")=strtsnr
				if CSng(strcycs)>1 and i<CSng(strcycs) Then
					Rs("cc")="���ڵ���"
				else
					Rs("cc")=strcc
				End if
				Rs.update
			next
		End If
		'���³����������еĵ��Դ���������ʱ�䡢������ʽ
		Dim strDxcs, strTscs
		strTscs=0 : strDxcs=0	'ģͷ�����͵��Դ���
		strSql="select * from [ins_tsxx] where lsh='"&strlsh&"'"
		set Rs=server.createobject("adodb.recordset")
		Rs.open strSql,Conn,1,1
		Do while not(Rs.eof Or Rs.bof)
			strTscs=strTscs+1
			strDxcs=strDxcs+1
			Rs.movenext
		Loop
		strTscs=strTscs &":"& strDxcs	'���Դ���
		strSql="select * from [mission] where lsh='"&strlsh&"'"
		set Rs=server.createobject("adodb.recordset")
		Rs.open strSql,Conn,1,3
		If Not(Rs.eof Or Rs.bof) Then
				Rs("sjtscs")=strTscs
				Rs("ccfs")=strcc
				If strcc<>"���ڵ���" Then Rs("sjjssj")=Date() End If
				Rs.update
		End If
	else	'�����������
		'���³����������е����շ�ʽ����ע
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
 						Rs("bz")="�� "&strbz
 					else
 						Rs("bz")=Rs("bz") &"<br>�� "& strbz
 					End If
 				End If
 			Rs.update
 			else
 				errmsg=errmsg&LshArr(i)&"ģ���ѷ��䣬���ʵ���ٴ����룡<br>"
 			End If
 			Rs.Close
		Next
		If errmsg<>"" Then
			Call WriteErrMsg()
			Exit Function
		End If
		'���µ�����Ϣ
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

		'����ģ�߷�ֵϵ��
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

	'���㳧�����Ա�÷�,��ʱ�����ܱ�ͳ�ƣ�
		If strcc="���պϸ�" or strcc="֪ͨ����"	Then
'		strSql="select * from [ext_tsfz] where lsh='"&LshArr(i)&"'"
'		set Rs=server.createobject("adodb.recordset")
'		Rs.open strSql,Conn,1,1
'		If not(Rs.eof Or Rs.bof) Then	'��ֹ�ظ��Ʒ�
'			Call WriteErrMsg()
'		else
			Call ext_fz(strlsh)
'		End If
'		Rs.Close
		End If
	End If
	Call WriteSuccessMsg("��ˮ�� " & strlsh &  " ���������ɹ�!","")
	Rs.Close
End Function

Function YLindb()
	Dim stryxtsyjs
	stryxtsyjs=Request("yxtsyjs")
	If stryxtsyjs="" Then errmsg="ʱ�䲻��Ϊ�գ�" End If
	If errmsg <> "" Then
		Call WriteErrMsg()
		Response.End
	End If
	Dim stryl
	stryl=Request("smyl")
	'���µ�������
	strSql="select * from [mission] where lsh='"&strlsh&"'"
	set Rs=server.createobject("adodb.recordset")
	Rs.open strSql,Conn,1,3
	If Not(Rs.eof Or Rs.bof) Then
		Rs("sjtsyl")=stryl
		Rs("rksj")=stryxtsyjs
		Rs("cksj")=stryxtsyjs
		Rs.update
	End If
'	'���㳧�ڵ���Ա�÷�
'	strSql="select * from [ins_tsfz] where lsh='"&strlsh&"'"
'	set Rs=server.createobject("adodb.recordset")
'	Rs.open strSql,Conn,1,3
'	If Rs.eof Or Rs.bof	Then	'��ֹ�ظ��Ʒ�
		Call ins_fz(stryxtsyjs)
'	End If
	Call WriteSuccessMsg("��ˮ�� " & strlsh &  " ���������ɹ�!","pld_assign.asp?rwlx=1")
	Rs.Close
End Function

Function ins_fz(stryxtsyjs)
	Dim strEdcs, strSjcs, TsGrxs, XsGrxs, strCsxs, strYlxs, strTszf, strTscs, strTsy, TsyArray
	Dim TmpSql, Rs2, ArrCs, ArrGrxs, Tmpxs, TsZgrxs, XsZgrxs, strGzz, Grxs, TmpArr, strXszf
	strTsy="" : ArrCs="" : ArrGrxs=""
	'�������ϡ����Դ������ͳ�����ʽ����ܵĽ���
	strSql="select * from [mission] where lsh='"&strlsh&"'"
	set Rs=server.createobject("adodb.recordset")
	Rs.open strSql,Conn,1,3
	If not(Rs.eof Or Rs.bof) Then
		strTszf=Rs("cntsfz")	'���ڵ��Է�ֵ
		strXszf=Rs("scxfz")		'�����߷�ֵ
		strEdcs=Split(Rs("edtscs"),":")
		strSjcs=Split(Rs("sjtscs"),":")
'		strCsxs=(strEdcs(0)-strSjcs(0))*0.02+(strEdcs(1)-strSjcs(1))*0.05	'���Դ���ϵ��
'		strYlxs=Fix((Rs("edtsyl")-Rs("sjtsyl"))/(Rs("edtsyl")*0.25))*0.05	'����ϵ��
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
			Case "��Ʒ�ϸ�"
				Tmpxs=1+strCsxs+strYlxs
			Case "һ��ϸ�"
				Tmpxs=1+strCsxs+strYlxs-0.1
			Case "����ϸ�"
				Tmpxs=1+strCsxs+strYlxs-0.2
			Case "����ϸ�"
				Tmpxs=1+strCsxs+strYlxs-0.3
			Case "֪ͨ����"
				Tmpxs=0.6
			Case "�ֵ�"
				Tmpxs=0.6
		End Select
		If Tmpxs<0.6 Then Tmpxs=0.6 End If
		strTszf=strTszf*Tmpxs
		strXszf=strXszf*Tmpxs
		Rs("cnsjfz")=Round(strTszf,1)
		Rs.update
	End If
	Rs.Close
'	'���������Ե�ÿ���˵�ϵ��
'	strSql="select * from [ins_tsxx] where lsh='"&strlsh&"' order by ID"
'	set Rs=server.createobject("adodb.recordset")
'	Rs.open strSql,Conn,1,1
'	ArrCs=Rs.recordcount
'	Do while not(Rs.eof Or Rs.bof)
'		Grxs="" : TsGrxs="" : XsGrxs="" : TsZgrxs=0 : XsZgrxs=0 : strGzz=""
'		'����Ա����ϵ�������ϸ���ϵ��,���ε��Բ������Աϵ��֮��,���ε��Բ��������ϵ��֮��
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
'						'�������"��:0.20:1;��:0.52:3;��:0.31:1;��:0.77:1"���ַ���
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
'				Rs("js")="����"
'			else
'				Rs("fz")=strXszf/ArrCs*TmpArr(1)
'				Rs("js")="����"
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
	If strrksj="" Then errmsg="���ʱ�䲻��Ϊ�գ�" End If
	If errmsg <> "" Then
		Call WriteErrMsg()
		Response.End
	End If
	'��ȡװ���ֵ
	strSql="select * from [mission] where lsh='"&strlsh&"'"
	set Rs=server.createobject("adodb.recordset")
	Rs.open strSql,Conn,1,1
	If Not(Rs.eof Or Rs.bof) Then
		strzxfz=Rs("bgxfz")
	End If
	Rs.Close
	'�������ʱ��
	strSql="select * from [mission] where lsh='"&strlsh&"'"
	set Rs=server.createobject("adodb.recordset")
	Rs.open strSql,Conn,1,3
	If Not(Rs.eof Or Rs.bof) Then
		Rs("rksj")=strrksj
		Rs.update
	End If
	Rs.Close
	'����������÷�ֵ
	strSql="select * from [ins_tsfz] where lsh='"&strlsh&"' and zrr='װ��'"
	set Rs=server.createobject("adodb.recordset")
	Rs.open strSql,Conn,1,3
	If Rs.eof Or Rs.bof Then
		Rs.AddNew
		Rs("lsh")=strlsh
		Rs("zrr")="װ��"
		Rs("js")="װ��"
		Rs("cycs")=1
		Rs("fz")=strzxfz
		Rs("sj")=strrksj
		Rs.update
	End If
	Call WriteSuccessMsg("��ˮ�� " & strlsh &  " �ɹ����!","")
	Rs.Close
End Function

Function Chuukuindb()
	Dim strcksj
	strcksj=Request("cksj")
	If strcksj="" Then errmsg="����ʱ�䲻��Ϊ�գ�" End If
	If errmsg <> "" Then
		Call WriteErrMsg()
		Response.End
	End If
	'���³���ʱ��
	strSql="select * from [mission] where lsh='"&strlsh&"'"
	set Rs=server.createobject("adodb.recordset")
	Rs.open strSql,Conn,1,3
	If Not(Rs.eof Or Rs.bof) Then
		Rs("cksj")=strcksj
		Rs.update
	End If
	Rs.Close
	Call WriteSuccessMsg("��ˮ�� " & strlsh &  " �ɹ�����!","")
End Function

Function Changeindb()
	Dim Number,tsy, strid, strtsnr, strTscs,strcs
	Number=Request("num")
	errmsg=""
	tsy=Request("zrr"&Number)
	If tsy="" Then
		errmsg="���һ�ε��Գ���������ɾ����"
		Call WriteErrMsg()
		exit Function
	End If
	For i=1 to Number
		strid=Request("id"&i)
		tsy=Request("zrr"&i)
		If tsy<>"" Then
			tsy=Replace(tsy,",",", ")
			'���µ�����Ϣ
			strSql="select * from [ins_tsxx] where id=" & Clng(strid)
			set Rs=server.createobject("adodb.recordset")
			Rs.open strSql,Conn,1,3
			If Not(Rs.eof Or Rs.bof) Then
				Rs("tsy")=tsy
				Rs.update
			End If
		else
			'����ԱΪ��ʱɾ����ε�����Ϣ
			''���Ȳ�ѯ���ε������ݸ���ģ�ߵ��Դ���
			strSql="select * from [ins_tsxx] where id=" & Clng(strid)
			set Rs=server.createobject("adodb.recordset")
			Rs.open strSql,Conn,1,1
			If Not(Rs.eof Or Rs.bof) Then
				strtsnr=Rs("tsnr")'��������
				strlsh=Rs("lsh")'��ˮ��
			End If
			strSql="select * from [mission] where lsh='"&strlsh&"'"
			set Rs=server.createobject("adodb.recordset")
			Rs.open strSql,Conn,1,1
			strTscs=Rs("sjtscs")
			strTscs=split(strTscs,":")
			If strtsnr="����" Then'�޸ĵ��Դ���
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

			'ɾ��������Ϣ
			sql="delete from ins_tsxx where ID=" & Clng(strid)
			conn.Execute sql
		End If
		Rs.Close
	next
	Call WriteSuccessMsg("���������˳ɹ�����!","")
End Function

Function ext_fz(slsh)
	Dim ilsh,strTsy,strkssj,strjssj,TmpArr,Tmpstr,TmpXs,Tmptsy,Tmpfz,j,n
	ilsh=slsh : strTsy="" :	strkssj="" : strjssj="" : TmpArr="" : Tmptsy="" : Tmpfz=0 : errmsg=""
	'------------------------------------------�������ʱ�䡢������Ա
	strlsh=Split(slsh,",")
	strSql="select * from [ext_tsxx] where lsh='"&strlsh(0)&"'"
	set Rs=server.createobject("adodb.recordset")
	Rs.open strSql,Conn,1,1
	Do while not rs.eof
		If strkssj="" Then
			strkssj=Rs("kssj")		'��ʼ����ʱ��
		else
			If datediff("d",strkssj,Rs("kssj"))<0 Then strkssj=Rs("kssj")
		End If
		If strjssj="" Then
			strjssj=Rs("jssj")		'��������ʱ��
		else
			If datediff("d",strjssj,Rs("jssj"))>0 Then strjssj=Rs("jssj")
		End If
		If strTsy="" Then
			strTsy=Rs("tsy")		'���������Ա
		else
			strTsy=strTsy&","&Rs("tsy")
		End If
	Rs.movenext
	loop
	Rs.Close
	strTsy=Split(MoveR(strTsy),",")	'���˵���Ա�б����ظ���
	'----------------------------------���������Աϵ���б�
	For i=0 to Ubound(strTsy)
		strSql="select * from [user] where UserName='"&strTsy(i)&"'"
		set Rs=server.createobject("adodb.recordset")
		Rs.open strSql,Conn,1,1
		If not rs.eof Then
			strTsy(i)=strTsy(i)&":"&Rs("Grxs")
		End If
		Rs.Close
	Next
	'-----------------------------------������Է�ֵ,���½�����ͳ��ʱ��
	strtsfz=0
	For i=0 to ubound(strlsh)
		strSql="select * from [mission] where lsh='"&strlsh(i)&"'"
		set Rs=server.createobject("adodb.recordset")
		Rs.open strSql,Conn,1,3
		If Not(Rs.eof Or Rs.bof) Then
			strtsfz=strtsfz + Rs("cwtsfz")*Rs("djxs")	'������Է�ֵ
			Rs("cwkssj")=strkssj						'���¿�ʼʱ��
			Rs("btjsj")=strjssj	 						'���½�����ͳ��ʱ��
			Rs.update
		End If
		Rs.Close
	Next
	'-----------------------------------����Ա�÷�
	Dim TmpKs,TmpJs
	TmpKs="" : TmpJs=""
'	errmsg="��"&datediff("d",strkssj,strjssj)&"��,strtsfz="&strtsfz&"<br>"
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
				else				'�ó�����ĳһ��Ĳ��������Ա
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
'    			errmsg=errmsg&Tmpstr&"��һ��:"&Tmptsy&",ϵ��:"&TmpXs&",��ֵ:"&Tmpfz&"<br>"
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
	Call WriteSuccessMsg("��ˮ�� "&ilsh&" �����������ɹ�!","")
End Function

Function ExtFindb()		'����ϵ�������¸��·�ֵ��д��ʱ�䣬ʹ�ܱ�ͳ��
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
	'��������ʵ�ʽ���ʱ��
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
	Call WriteSuccessMsg("��ˮ�� "&strlsh&" �����������ɹ�!","")
End Function

'******************************
'������MoveR(Rstr)
'������Rstr,�Զ������ָ����ַ���
'������ȥ���������ظ�����
'ʾ����MoveR("abc,abc,dge,gcg,dge,dgi,die,dir,fgk,dir,gis,sgier,ssir")
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
