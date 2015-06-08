<!--#include file="Inc/Sysconn.asp" -->
<!--#include file="inc/Skin_CSS.asp"-->
<!--#include file="Inc/function.asp"-->
<%
	Dim action
	action=Request("action")

	Dim strlsh, strddh, strkhmc,strdmdj,strmjlx, strxcjy, strrwlx,strtsyl
	Dim strmz, strbh, strtscs, strjhjs, strscxh, strqysd, strsml, strgjl,strrwnr,strqs
	Dim StrCntsfz, StrScxfz, StrCwtsfz, StrZxfz
	Dim strcjqy, strsd

	'任务书所有变量初始化
strlsh=Trim(Request("lsh")) : strddh=Trim(Request("ddh")) : strkhmc=Trim(Request("khmc")) : strrwnr=Request("rwnr") : strdmdj=Request("dmdj") : strmjlx=Request("mjlx") : strmz=Request("mz") : strbh=Request("bh") : strtscs=Request("tscs") : strjhjs=Request("jhjs") : strscxh=Request("scxh") : strqysd=Request("qysd") : strxcjy=Request("xcjy") : StrCntsfz=Request("Cntsfz") : StrScxfz=Request("Scxfz") : StrCwtsfz=Request("Cwtsfz") : StrZxfz=Request("Zxfz") : strtsyl=Request("tsyl") : strqs=Request("qs")
strcjqy=Request("cjqy") : strsd=Request("sd")
	If Request("smlqk")="代用料厂家" Then
		strsml=Request("smltxt")
	else
		strsml=Request("smlqk")
	End If
	If Request("gjlqk")="代用料厂家" Then
		strgjl=Request("gjltxt")
	else
		strgjl=Request("gjlqk")
	End If
	If strbh<>1 Then strbh=0 End If
	If strqs<>1 Then strqs=0 End If
	If strsd<>1 Then strsd=0 End If
	If strxcjy<>1 Then strxcjy=0 End If
	'对待入库数据进行处理
	errmsg=""
	If strlsh="" Then errmsg="流水号为空!<br>"
	If strddh="" Then errmsg=errmsg & "订单号为空!<br>"
	If strkhmc="" Then errmsg=errmsg & "客户名称为空!<br>"
	If strtscs=""  Then errmsg=errmsg & "额定调试次数为空!<br>"
	'If strjhjs=""  Then errmsg=errmsg & "计划结束时间为空!<br>"
	If strsml=""  Then errmsg=errmsg & "试模料情况不能为空!<br>"
	If strqysd=""  Then errmsg=errmsg & "牵引速度不能为空!<br>"

	If errmsg <> "" Then
		Call WriteErrMsg()
	else
		'数据入库函数从这里开始
		Select Case action
			Case "add"
				strrwlx=Request("rwlx")
				Call mtask_add()
			Case "change"
				strrwlx=Request("rwlx")
				Call mtask_change()
			Case "del"
				Call mtask_del()
			Case else
				response.write "action=" & action
		End select
	End If


	'添加任务书入库
	Function mtask_add()
		'检测流水号是否已存在
		strSql="select lsh from mission where lsh='"&strlsh&"'"
		set rs=server.createobject("adodb.recordset")
		rs.open strSql,conn,1,1
		If Not(Rs.eof Or Rs.bof) Then
			errmsg="流水号 " & strlsh & " 任务书已存在!请更改流水号!"
			Call WriteErrMsg()
			Exit Function
		End If
		Rs.Close

		strSql="select * from [mission]"
		set Rs=server.createobject("adodb.recordset")
		Rs.open strSql,Conn,1,3
		Rs.AddNew
			Rs("lsh")=strlsh
			Rs("ddh")=strddh
			Rs("khmc")=strkhmc
			Rs("cjqy")=strcjqy
			Rs("rwlx")=strrwnr
			Rs("dmdj")=strdmdj
			Rs("mjlx")=strmjlx
			If strmz<>"" Then Rs("mz")=strmz End If
			Rs("bh")=strbh
			Rs("qs")=strqs
			Rs("sd")=strsd
			Rs("qysd")=strqysd
			Rs("cwtsfz")=StrCwtsfz
			Rs("jy")=strxcjy
			Rs("edtscs")=strtscs
	 		If strscxh<>"" Then Rs("scx")=strscxh End If
			Rs("smlqk")=strsml
			If strgjl<>"" Then Rs("gjlqk")=strgjl End If
			If strtsyl<>"" Then Rs("edtsyl")=strtsyl End If
			Rs("cntsfz")=StrCntsfz
			Rs("scxfz")=StrScxfz
			Rs("bgxfz")=StrZxfz
		Rs.update
		Rs.Close
		Call WriteSuccessMsg("流水号 " & strlsh &  " 任务书添加成功!","")
		Response.End
	End Function

	'更改任务书入库
	Function mtask_change()
		'检测流水号是否已存在
		Dim strid
		strid=Clng(Request("iid"))
		sql="select * from [mission] where lsh='"&strlsh&"'"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,conn,1,1
		If not Rs.eof Then
			If Rs("id")	<> strid Then
				errmsg="流水号 " & strlsh & " 任务书已存在，不能改为该流水号！！"
				Call WriteErrMsg()
				Exit Function
			End If
		End If
		Rs.Close

		strSql="select * from [mission] where ID=" & strid
		set rs=server.createobject("adodb.recordset")
		Rs.open strSql,Conn,1,3
			Rs("ddh")=strddh
			Rs("khmc")=strkhmc
			Rs("cjqy")=strcjqy
			Rs("rwlx")=strrwnr
			Rs("dmdj")=strdmdj
			Rs("mjlx")=strmjlx
			If strmz<>"" Then Rs("mz")=strmz End If
			Rs("bh")=strbh
			Rs("qs")=strqs
			Rs("sd")=strsd
			Rs("qysd")=strqysd
			Rs("cwtsfz")=StrCwtsfz
			Rs("jy")=strxcjy
			Rs("edtscs")=strtscs
	 		If strscxh<>"" Then Rs("scx")=strscxh End If
			Rs("smlqk")=strsml
			If strgjl<>"" Then Rs("gjlqk")=strgjl End If
			If strtsyl<>"" Then Rs("edtsyl")=strtsyl End If
			Rs("cntsfz")=StrCntsfz
			Rs("scxfz")=StrScxfz
			Rs("bgxfz")=StrZxfz
	 		If Rs("lsh")<>strlsh Then
	 			Call TsxxChg(Rs("lsh"),strlsh)
	 			Rs("lsh")=strlsh
			End If
		Rs.update
		Rs.Close
		Call WriteSuccessMsg("流水号 " & strlsh &  " 任务书更改成功!","")
	End Function

	Function TsxxChg(Yllsh,Xzlsh)
		TmpSql="select * from [ins_tsxx] where [lsh]='"&Yllsh&"'"
		set rs2=server.createobject("adodb.recordset")
		Rs2.open TmpSql,Conn,1,3
		If not(Rs2.eof) Then
			Rs2("lsh")=Xzlsh
			Rs2.update
		End If
		Rs2.Close

		TmpSql="select * from [ins_tsfz] where [lsh]='"&Yllsh&"'"
		set rs2=server.createobject("adodb.recordset")
		Rs2.open TmpSql,Conn,1,3
		If not(Rs2.eof) Then
			Rs2("lsh")=Xzlsh
			Rs2.update
		End If
		Rs2.Close
	End Function
%>
