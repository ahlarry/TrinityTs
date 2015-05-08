<!--#include file="Inc/Sysconn.asp" -->
<!--#include file="inc/Skin_CSS.asp"-->
<!--#include file="Inc/function.asp"-->
<%
	Dim action
	action=Request("action")

	Dim strrwlx, strxldh, strrwnr, strfz, strjssj, strzrr, strLNum, strTNum, k
	strrwlx=Request("rwlx")
	strxldh=Request("xldh")
	strrwnr=Request("rwnr")
	strfz=Request("fz")
	strjssj=Request("jssj")
	strzrr=Request("zrr")
	strLNum=Csng(Request("LshNum"))	'调试参与流水号
	strTNum=Csng(Request("TsyNum"))	'调试参与调试员人数
	errmsg=""
	
	Select case action
		Case "add"
			Call spc_add()
		Case "change"
			Call spc_cha()
		Case "Tscha"
			Call spc_Tsc()
		Case "del"
			Call spc_del()
	End Select
		
Function spc_add()
	Dim tmpzrr,Xs,TmpArr
	Xs=0
	if strrwlx<>"调试" Then 
    	If strxldh="" Then errmsg="修理单号不能为空!<br>"
    	If strrwnr="" Then errmsg=errmsg & "任务内容不能为空!<br>"
    	If strfz="" Then errmsg=errmsg & "分值不能为空!<br>"
    	If strjssj="" Then errmsg=errmsg & "结束时间不能为空!<br>"    	
    	If strzrr="" Then 
    		errmsg=errmsg & "责任人不能为空!<br>"
    	else 
       		tmpzrr=Replace(strzrr,"，",",")
       		tmpzrr=split(tmpzrr,",")	
       	End If
       	'检测用户名是否正确
    	For i=0 to ubound(tmpzrr)
    		strSql="select * from [user] where UserName='"&tmpzrr(i)&"'"
    		set Rs=server.createobject("adodb.recordset")
    		Rs.open strSql,Conn,1,1
    		If not rs.eof Then
    			tmpzrr(i)=tmpzrr(i)&"||"&Rs("Grxs")
    			Xs=Xs + Rs("Grxs")
    		else
    			errmsg=errmsg &"调试员 <b>"& tmpzrr(i) &"</b> 不存在，请核对后重新输入!<br>"
    		End If
    		Rs.Close
    	Next

    	If errmsg <> "" Then
    		Call WriteErrMsg()
    		Response.End
    	End If
    	
    	'添加零星任务
    	strSql="select * from [spc_mission]"
    	set Rs=server.createobject("adodb.recordset")
    	Rs.open strSql,Conn,1,3
    	For i=0 to ubound(tmpzrr)
        	Rs.AddNew
        	Rs("rwlx")=strrwlx
        	Rs("xldh")=strxldh
        	Rs("rwnr")=strrwnr
        	Rs("wcsj")=strjssj
    		Rs("rksj")=strjssj 	
    		TmpArr=split(tmpzrr(i),"||")
        	Rs("zrr")=TmpArr(0)
        	Rs("fz")=Round(strfz*TmpArr(1)/Xs,2)
        	Rs.update
    	Next
    	Rs.Close	
    	Call WriteSuccessMsg("修理单号 " & strxldh &  " 零星任务书添加成功!","")
    else
   		'添加调试信息
	  	Dim strmjlb,strkssj,strbz
   		strlsh=Request("lsh")
   		strmjlb=Request("mjlb")
   		strfz=Request("rwzf")
   		strzrr=Request("tsy")
   		strkssj=Request("tskssj")
  		strjssj=Request("tsjssj")
   		strbz=Request("beiz")
   		
   		strSql="select * from spc_mission where rwlx='调试' and xldh='" &strlsh& "'"
		set Rs=server.createobject("adodb.recordset")
		Rs.open strSql,Conn,1,1
		If not rs.eof Then
			errmsg=errmsg &"单号 <b>"& strlsh &"</b> 已存在，请核对后重新输入!<br>"
		End If
		Rs.Close

    	If errmsg <> "" Then
    		Call WriteErrMsg()
    		Response.End
    	End If

   		If strLNum>1 Then
			for i = 1 to strLNum-1
				strlsh=strlsh&"||"&Request(i&"lsh")
				strmjlb=strmjlb&"||"&Request(i&"mjlb")
				strfz=strfz&"||"&Request(i&"rwzf")
			next
		End If
   		If strTNum>1 Then
			for i = 1 to strTNum-1
				strzrr=strzrr&"||"&Request(i&"tsy")
				strkssj=strkssj&"||"&Request(i&"tskssj")
				strjssj=strjssj&"||"&Request(i&"tsjssj")
			next
		End If
		strlsh=split(strlsh,"||")
		strmjlb=split(strmjlb,"||")
		strfz=split(strfz,"||")
		Dim TmpLsh, Tmpmjlb,Tmpfz
		For k=0 to ubound(strlsh)
			TmpLsh=strlsh(k)
			Tmpmjlb=strmjlb(k)
			Tmpfz=strfz(k)
			Call ext_fz(TmpLsh,Tmpmjlb,Tmpfz,strzrr,strkssj,strjssj,strbz)
		Next
		Call WriteSuccessMsg("修理单号 " & strxldh &  " 零星任务书添加成功!","")
	End If
End Function

Function spc_cha()	'更改零星任务
	Dim strydh,strysj,tmpzrr,TmpArr,Xs,strid
	strid=Request("id") : strydh=Request("ydh") : strysj=Request("ysj")
	Xs=0
   	If strxldh="" Then errmsg="修理单号不能为空!<br>"	
	If strrwnr="" Then errmsg=errmsg & "任务内容不能为空!<br>"
	If strfz="" Then errmsg=errmsg & "分值不能为空!<br>"
	If strjssj="" Then errmsg=errmsg & "结束时间不能为空!<br>"
	If strzrr="" Then 
		errmsg=errmsg & "责任人不能为空!<br>"
	else
		strzrr=Replace(strzrr,"，",",")
	End If
	tmpzrr=split(strzrr,",")
	For i=0 to ubound(tmpzrr)
		strSql="select * from [user] where UserName='"&tmpzrr(i)&"'"
		set Rs=server.createobject("adodb.recordset")
		Rs.open strSql,Conn,1,1
		If not rs.eof Then
			tmpzrr(i)=tmpzrr(i)&"||"&Rs("Grxs")
			Xs=Xs + Rs("Grxs")
		else
			errmsg=errmsg &"调试员 <b>"& tmpzrr(i) &"</b> 不存在，请核对后重新输入!<br>"
		End If
		Rs.Close
	Next

'	errmsg="strydh="&strydh&",strxldh="&strxldh&"strrwlx="&strrwlx
	
	If strydh<>strxldh Then 
    	strSql="select * from spc_mission where xldh='" &strxldh& "' and rwlx='" &strrwlx& "'"
    	set Rs=server.createobject("adodb.recordset")
    	Rs.open strSql,Conn,1,1
    	If not Rs.eof and datediff("d",Rs("wcsj"),strysj)<>0 Then 
    		errmsg=errmsg &strid&"单号"&strxldh&"已经存在，无法改为该单号!!<br>"
    	End if
    	Rs.Close
	End If 
	
	If errmsg <> "" Then
		Call WriteErrMsg()
		Response.End
	End If
	
	'删除旧的记录
	sql="delete from spc_mission where rwlx='" &strrwlx& "' and xldh='" &strydh& "' and datediff('d',wcsj,'"&strysj&"')=0"
	conn.Execute sql
	'添加新记录
   	strSql="select * from [spc_mission]"
   	set Rs=server.createobject("adodb.recordset")
   	Rs.open strSql,Conn,1,3
   	For i=0 to ubound(tmpzrr)
       	Rs.AddNew
       	Rs("rwlx")=strrwlx
       	Rs("xldh")=strxldh
       	Rs("rwnr")=strrwnr
       	Rs("wcsj")=strjssj
   		Rs("rksj")=strjssj 	
  		TmpArr=split(tmpzrr(i),"||")
        Rs("zrr")=TmpArr(0)
       	Rs("fz")=Round(strfz*TmpArr(1)/Xs,2)
       	Rs.update
   	Next
   	Rs.Close	
	Call WriteSuccessMsg(" 零星任务书修改成功!","spc_list.asp")
End Function

Function spc_Tsc()	'更改厂外调试任务
	Dim strid,strrwlx,strydh,strmjlb,strkssj,strbz
	strid=Request("id")	
   	strSql="select * from spc_mission where ID=" & Clng(strid)
   	set Rs=server.createobject("adodb.recordset")
   	Rs.open strSql,Conn,1,3
	If not Rs.eof Then 
		strrwlx=Rs("rwlx")
		strydh=Rs("xldh")
	End if
	Rs.Close	
	
   	strlsh=Request("lsh")
	strmjlb=Request("mjlb")
   	strfz=Request("rwzf")
 	strzrr=Request("tsy")
   	strkssj=Request("tskssj")
 	strjssj=Request("tsjssj")
   	strbz=Request("beiz")
 
 	If strydh<>strlsh Then 
    	strSql="select * from spc_mission where rwlx='调试' and xldh='" &strlsh& "'"
    	set Rs=server.createobject("adodb.recordset")
    	Rs.open strSql,Conn,1,1
    	If not rs.eof Then
    		errmsg=errmsg &"单号 <b>"& strlsh &"</b> 已存在，请核对后重新输入!<br>"
    	End If
    	Rs.Close
	End If 

    If errmsg <> "" Then
   		Call WriteErrMsg()
   		Response.End
   	End If
	'删除旧的记录
	sql="delete from spc_mission where rwlx='" &strrwlx& "' and xldh='" &strydh& "'"
	conn.Execute sql
	'添加新记录
   	If strTNum>1 Then
		for i = 1 to strTNum-1
			strzrr=strzrr&"||"&Request(i&"tsy")
			strkssj=strkssj&"||"&Request(i&"tskssj")
			strjssj=strjssj&"||"&Request(i&"tsjssj")
		next
	End If	
	Call ext_fz(strlsh,strmjlb,strfz,strzrr,strkssj,strjssj,strbz)
	Call WriteSuccessMsg("单号 " & strlsh &  " 零星任务书修改成功!","spc_list.asp")
End Function

Function spc_del()
	Dim strid,strrwlx,strydh,strysj
	strid=Request("id")
	
   	strSql="select * from spc_mission where ID=" & Clng(strid)
   	set Rs=server.createobject("adodb.recordset")
   	Rs.open strSql,Conn,1,3
	If not Rs.eof Then 
		strrwlx=Rs("rwlx")
		strydh=Rs("xldh")
		strysj=Rs("wcsj")
	End if
	Rs.Close
	
	'删除记录
	sql="delete from spc_mission where rwlx='" &strrwlx& "' and xldh='" &strydh& "' and datediff('d',wcsj,'"&strysj&"')=0"
	conn.Execute sql

	Call CloseConn()      		
	Call WriteSuccessMsg("删除成功","")	
End Function

Function ext_fz(slsh,smjlb,sfz,szrr,skssj,sjssj,sbz)
	Dim strTsy,strkssj,strjssj,KssjArr,JssjArr,TmpArr,Tmpsj,TmpXs,Tmptsy,Tmpfz,j,n
	strTsy="" :	strkssj="" : strjssj="" : TmpArr="" : Tmptsy="" : Tmpfz=0 : errmsg=""
'	errmsg=slsh&","&smjlb&","&sfz&","&szrr&","&skssj&","&sjssj&","&sbz&"<br>"&strrwlx&"<br>strTNum="&strTNum
	'----------------------------------参与调试人员系数列表
	strTsy=split(szrr,"||")
	For i=0 to Ubound(strTsy)	
		strSql="select * from [user] where UserName='"&strTsy(i)&"'"
		set Rs=server.createobject("adodb.recordset")
		Rs.open strSql,Conn,1,1
		If not rs.eof Then
			strTsy(i)=strTsy(i)&":"&Rs("Grxs")
		End If
		Rs.Close
	Next
	'-----------------------------------调试时间
	KssjArr=split(skssj,"||")
	JssjArr=split(sjssj,"||")	
	Tmpsj=KssjArr(0)
	For i=1 to Ubound(KssjArr)	
		If datediff("d",Tmpsj,KssjArr(i))<0 Then Tmpsj=KssjArr(i)
	Next
	strkssj=Tmpsj			'开始调试时间
	Tmpsj=JssjArr(0)
	For i=1 to Ubound(JssjArr)	
		If datediff("d",Tmpsj,JssjArr(i))>0 Then Tmpsj=JssjArr(i)
	Next
	strjssj=Tmpsj			'结束调试时间
	'errmsg=errmsg&strkssj&"/"&strjssj&"<br>"
	'-----------------------------------调试员得分
	For i=0 to datediff("d",strkssj,strjssj)
		TmpArr=""
		Tmpsj=dateadd("d",i,strkssj)	
		For j=0 to ubound(KssjArr)	
			If datediff("d",Tmpsj,KssjArr(j))<=0 and datediff("d",Tmpsj,JssjArr(j))>=0 Then
				If TmpArr="" Then
					TmpArr=strTsy(j)
				else				'得出具体某一天的参与调试人员
					TmpArr=TmpArr&"||"&strTsy(j)
				End If 
			End If
		next
		If TmpArr<>"" Then 
    		TmpArr=Split(TmpArr,"||")
    		TmpXs=0
    		For j=0 to ubound(TmpArr)
    			TmpXs=TmpXs+Mid(TmpArr(j),Instr(TmpArr(j),":")+1)
    		Next
    		For j=0 to ubound(TmpArr)
    			Tmptsy=Left(TmpArr(j),Instr(TmpArr(j),":")-1)
    			Tmpfz=Round(sfz/(datediff("d",strkssj,strjssj)+1)*Mid(TmpArr(j),Instr(TmpArr(j),":")+1)/TmpXs,3)
    		'	errmsg=errmsg&Tmpsj&"这一天:"&Tmptsy&",系数:"&TmpXs&",分值:"&Tmpfz&"<br>"
    			strSql="select * from [spc_mission] where rwlx='"&strrwlx&"' and xldh='"&slsh&"' and zrr='"&Tmptsy&"'"
    			set Rs=server.createobject("adodb.recordset")
    			Rs.open strSql,Conn,1,3
   			 	If Rs.eof Or Rs.bof Then
       			 	Rs.AddNew
      		  		Rs("rwlx")=strrwlx
       		 		Rs("xldh")=slsh
        		  	Rs("rwnr")="模具类别:"&smjlb&",模具分:"&sfz&"。"&sbz
        			Rs("fz")=Tmpfz
        			Rs("rksj")=Date()
        			Rs("zrr")=Tmptsy
       				'更新此调试员参与的时间段
        			For n=0 to ubound(strTsy)
    					if Tmptsy = Left(strTsy(n),Instr(strTsy(n),":")-1) Then 
    					  	Rs("kssj")=KssjArr(n)  				
    					  	Rs("wcsj")=JssjArr(n)  	
    					End If
    				Next	
        			Rs.update
   				else 
     				Rs("fz")=Rs("fz") + Tmpfz
    				Rs.update			    				
    			End If
    			Rs.Close						
  			next
		End If
	next 
	if errmsg<>"" Then
		Call WriteErrMsg()
	End If
End Function
%>
