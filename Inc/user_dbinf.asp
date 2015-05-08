<%
	'10:46 2007-7-10-星期二
	dim c_alluser, c_alljcx, c_allscx, c_alljl, c_allzy, c_alladm,UserSql,tmpname, tmpGroup
	'所有人员，所有组员，挤出线，生产线，所有经理，所有管理员
	c_alluser="" : c_allzy="" : c_alljcx="" : c_allscx="" : c_alljl="" : c_alladm=""
	UserSql="select UserName,Grxs,Gzz from [User] order by Gzz,UserName,Grxs Desc"
	set rs=server.createobject("adodb.recordset")
	rs.open UserSql,conn,1,1
	do while not rs.eof
		tmpname=replace(rs("UserName"),"|||","")
		tmpGroup=Rs("Gzz")
		'所有人员
		if c_alluser <> "" then 
			c_alluser=c_alluser & "|||" & tmpname
		else
			c_alluser=tmpname
		end if
		'所有组员
		If tmpGroup>0 Then
			if c_allzy <> "" then 
				c_allzy=c_allzy & "|||" & tmpname
			else
				c_allzy=tmpname
			end if
		End If
		'组员数组
			Select Case tmpgroup
				Case -1			'管理员
					If c_alladm<>"" Then
						c_alladm=c_alladm & "|||" & tmpname
					Else
						c_alladm=tmpname
					End If			
				Case 0			'经理
					If c_alljl<>"" Then
						c_alljl=c_alljl & "|||" & tmpname
					Else
						c_alljl=tmpname
					End If
				Case 1			'挤出线
					If c_alljcx<>"" Then
						c_alljcx=c_alljcx & "|||" & tmpname
					Else
						c_alljcx=tmpname
				End If
				Case 3			'生产线
					If c_allscx<>"" Then
						c_allscx=c_allscx & "|||" & tmpname
					Else
						c_allscx=tmpname
					End If
			End Select
		rs.movenext
	loop
	rs.close
	If c_alluser="" Then c_alluser=" " End If
	If c_allzy="" Then c_allzy=" " End If
	If c_alladm="" Then c_alladm=" " End If
	If c_alljl="" Then c_alljl=" " End If
	If c_alljcx="" Then c_alljcx=" " End If
	If c_allscx="" Then c_allscx=" " End If
	c_alluser = split(c_alluser, "|||")
	c_allzy = split(c_allzy, "|||")
	c_alladm = split(c_alladm, "|||")
	c_alljl = split(c_alljl, "|||")
	c_alljcx = split(c_alljcx, "|||")
	c_allscx = split(c_allscx, "|||")
%>