<!--#include file="Inc/Sysconn.asp" -->
<!--#include file="inc/Skin_CSS.asp"-->
<!--#include file="Inc/function.asp"-->
<%
	Dim action
	action=Request("action")

	Dim strzrr, strkhm, strlsh, strsj, strzywt, stryyfx, strjzcs
	strzrr=Request("zrr")
	strkhm=Request("khm")
	strlsh=Request("lsh")
	strsj=Request("sj")	
	strzywt=Request("zywt")
	stryyfx=Request("yyfx")
	strjzcs=Request("jzcs")
	errmsg=""
	
	Select case action
		Case "add"
			Call Faq_add()
	End Select
		
Function Faq_add()
	Dim tmpzrr
	If strkhm="" Then errmsg=errmsg & "�ͻ�������Ϊ��!<br>"
   	If strlsh="" Then errmsg=errmsg & "��ˮ�Ų���Ϊ��!<br>"
   	If strsj="" Then errmsg=errmsg & "ʱ�䲻��Ϊ��!<br>"    	
   	If strzywt="" Then errmsg=errmsg & "��Ҫ���ⲻ��Ϊ��!<br>"    	
   	If strzrr="" Then 
  		errmsg=errmsg & "�����˲���Ϊ��!<br>"
   	else 
     	tmpzrr=Replace(strzrr,"��",",")
       	tmpzrr=split(tmpzrr,",")	
   	End If
    '����û����Ƿ���ȷ
   	For i=0 to ubound(tmpzrr)
   		strSql="select * from [user] where UserName='"&tmpzrr(i)&"'"
   		set Rs=server.createobject("adodb.recordset")
   		Rs.open strSql,Conn,1,1
    	If rs.eof Then
   			errmsg=errmsg &"����Ա <b>"& tmpzrr(i) &"</b> �����ڣ���˶Ժ���������!<br>"
    	End If
    	Rs.Close
   	Next

   	If errmsg <> "" Then
   		Call WriteErrMsg()
   		Response.End
    End If
    	
   	'����������
   	strSql="select * from [Faq]"
   	set Rs=server.createobject("adodb.recordset")
    Rs.open strSql,Conn,1,3
       	Rs.AddNew
       	Rs("zrr")=strzrr
       	Rs("Khmc")=strkhm
       	Rs("Lsh")=strlsh
       	Rs("Sj")=strsj
       	Rs("WenTi")=strzywt
       	If stryyfx<>"" Then Rs("YuanY")=stryyfx
       	If strjzcs<>"" Then Rs("CuoS")=strjzcs
       	Rs.update
   	Rs.Close	
   	Call WriteSuccessMsg("��ˮ��Ϊ <strong>" & strlsh &  "</strong> �����������ӳɹ�!","")
End Function
%>
