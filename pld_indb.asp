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

	'���������б�����ʼ��
strlsh=Trim(Request("lsh")) : strddh=Trim(Request("ddh")) : strkhmc=Trim(Request("khmc")) : strrwnr=Request("rwnr") : strdmdj=Request("dmdj") : strmjlx=Request("mjlx") : strmz=Request("mz") : strbh=Request("bh") : strtscs=Request("tscs") : strjhjs=Request("jhjs") : strscxh=Request("scxh") : strqysd=Request("qysd") : strxcjy=Request("xcjy") : StrCntsfz=Request("Cntsfz") : StrScxfz=Request("Scxfz") : StrCwtsfz=Request("Cwtsfz") : StrZxfz=Request("Zxfz") : strtsyl=Request("tsyl") : strqs=Request("qs")
strcjqy=Request("cjqy") : strsd=Request("sd")
	If Request("smlqk")="�����ϳ���" Then
		strsml=Request("smltxt")
	else
		strsml=Request("smlqk")
	End If
	If Request("gjlqk")="�����ϳ���" Then
		strgjl=Request("gjltxt")
	else
		strgjl=Request("gjlqk")
	End If
	If strbh<>1 Then strbh=0 End If
	If strqs<>1 Then strqs=0 End If
	If strsd<>1 Then strsd=0 End If
	If strxcjy<>1 Then strxcjy=0 End If
	'�Դ�������ݽ��д���
	errmsg=""
	If strlsh="" Then errmsg="��ˮ��Ϊ��!<br>"
	If strddh="" Then errmsg=errmsg & "������Ϊ��!<br>"
	If strkhmc="" Then errmsg=errmsg & "�ͻ�����Ϊ��!<br>"
	If strtscs=""  Then errmsg=errmsg & "����Դ���Ϊ��!<br>"
	'If strjhjs=""  Then errmsg=errmsg & "�ƻ�����ʱ��Ϊ��!<br>"
	If strsml=""  Then errmsg=errmsg & "��ģ���������Ϊ��!<br>"
	If strqysd=""  Then errmsg=errmsg & "ǣ���ٶȲ���Ϊ��!<br>"

	If errmsg <> "" Then
		Call WriteErrMsg()
	else
		'������⺯�������￪ʼ
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


	'������������
	Function mtask_add()
		'�����ˮ���Ƿ��Ѵ���
		strSql="select lsh from mission where lsh='"&strlsh&"'"
		set rs=server.createobject("adodb.recordset")
		rs.open strSql,conn,1,1
		If Not(Rs.eof Or Rs.bof) Then
			errmsg="��ˮ�� " & strlsh & " �������Ѵ���!�������ˮ��!"
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
		Call WriteSuccessMsg("��ˮ�� " & strlsh &  " ��������ӳɹ�!","")
		Response.End
	End Function

	'�������������
	Function mtask_change()
		'�����ˮ���Ƿ��Ѵ���
		Dim strid
		strid=Clng(Request("iid"))
		sql="select * from [mission] where lsh='"&strlsh&"'"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,conn,1,1
		If not Rs.eof Then
			If Rs("id")	<> strid Then
				errmsg="��ˮ�� " & strlsh & " �������Ѵ��ڣ����ܸ�Ϊ����ˮ�ţ���"
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
		Call WriteSuccessMsg("��ˮ�� " & strlsh &  " ��������ĳɹ�!","")
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
