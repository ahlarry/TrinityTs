<!--#include file="Inc/Sysconn.asp" -->
<!--#include file="inc/Skin_CSS.asp"-->
<!--#include file="Inc/function.asp"-->
<%
	Dim action, strkpxm, strkpnr, strkpfz, strkpsj, strzrr, strbeiz, strkpr
	action=Request("action")
	strkpxm=Request("kpxm")
	strkpnr=Request("kpnr")
	strkpfz=Csng(Request("kpf"))
	strkpsj=Request("kpsj")
	strzrr=Request("zrr")
	strbeiz=Request("beiz")
	strkpr=session("UserName")

	Select case action
		Case "add"
			Call Addindb()
		Case "del"
			Call Delindb()
	End Select
		
	Function Addindb()
			strSql="select * from [kplb]"
			set Rs=server.createobject("adodb.recordset")
			Rs.open strSql,Conn,1,3
				Rs.AddNew
				Rs("kpnr")=strkpnr
				Rs("kpxm")=strkpxm
				Rs("fz")=strkpfz
				Rs("kpsj")=strkpsj
				Rs("zrr")=strzrr
				If strbeiz<>"" Then Rs("bz")=strbeiz End If
				Rs("kpr")=strkpr
				Rs.update
			Rs.Close
		Call WriteSuccessMsg("������ӳɹ�!","")				
	End Function
	
	Function Delindb()
		Dim strid
		strid=Request("id")
		'ɾ������
		if strid<>"" then
			sql="delete from kplb where ID=" & Clng(strid)
			conn.Execute sql
		end if
		Call CloseConn()      		
		Call WriteSuccessMsg("ɾ���ɹ�","")	
End Function
%>