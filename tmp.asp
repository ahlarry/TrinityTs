<%
		If Not(Rs.eof Or Rs.bof) Then
			If IsNull(Rs("sjjssj")) and Rs("rwlx")<>"�ǵ�" Then				
				Call InsAssign()
			else
				if Rs("sjtsyl")=0  and Rs("rwlx")<>"�ǵ�" Then
					Call YlAssign()
				else
					if isnull(Rs("rksj")) Then 
						Call RkAssign()
					else
						If IsNull(Rs("cksj")) Then
						 	Call CkAssign()
						End If					
					End If 
				End If	
			End if
		else
			Response.write("<font color='red'>���񲻴��ڻ�����ɣ�</font>")
		End If

If Not(Rs.eof Or Rs.bof) Then
	If Rs("rwlx")="�ǵ�" or Rs("rwlx")="TB" Then
		If IsNull(Rs("cksj")) Then
		 	Call CkAssign()
		End If					
	else
		If IsNull(Rs("sjjssj")) Then 
			Call InsAssign()
		else 
			if Rs("sjtsyl")=0 Then 
				Call YlAssign()
			else
				if isnull(Rs("rksj")) Then 
					Call RkAssign()
				else	
					If IsNull(Rs("cksj")) Then
						Call CkAssign()
					End If	
				End If
			End If
		End If
	End If 
End If
%>
