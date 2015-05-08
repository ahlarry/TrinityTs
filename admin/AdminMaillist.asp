<%@language=VBScript codepage=936 %>
<%
response.buffer=true
%>
<!--#include file="conn.asp"-->
<!--#include file="../inc/config.asp"-->
<!--#include file="Admin.asp"-->
<!--#include file="inc/function.asp"-->
<%
dim sql,rs,Action,FoundErr,ErrMsg
dim JMObjInstalled
Action=trim(request("Action"))
Email=trim(request("Email"))
mane=trim(request("Name"))
JMObjInstalled=IsObjInstalled("JMail.Message")

dim FSObjInstalled
FSObjInstalled=IsObjInstalled("Scripting.FileSystemObject")


%>
<!-- #include file="Inc/Head.asp" -->
<table width="90%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#000000" class="border">
  <tr class="topbg"> 
    <td class="back_southidc" height="28" colspan="2" align="center" bgcolor="#FFFFFF"><strong>邮 
      件 列 表 管 理</strong></td>
  </tr>
  <tr bgcolor="#FFFFFF" class="tdbg"> 
    <td width="101" height="30" bgcolor="#A4B6D7"><div align="right">管理导航：</div></td>
    <td width="595" height="30"><a href="AdminMaillist.asp">发送邮件列表</a> | <a href="AdminMaillist.asp?Action=Export">导出邮件列表</a> 
  </tr>
</table>
<br>
<%
if Action="Send" then
	call SendMaillist()
elseif Action="Export" then
	call ExportMail()
elseif Action="DoExport" then
	call DoExportList()
else
	call main()
end if

if FoundErr=True then
	call WriteErrMsg()
end if

sub main()
%>
<form method="POST" action="AdminMaillist.asp?Action=Send">
  <table width="90%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#000000" Class="border">
    <tr bgcolor="#FFFFFF" class="title"> 
      <td class="back_southidc" height="28" colspan=2 align=center><b> 邮 件 列 表</b></td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tdbg"> 
      <td rowspan="3" align="right" bgcolor="#A4B6D7">收件人：</td>
      <td width="85%"> 
        <input type="radio" name="incepttype" value="1">
        邮件发给所有注册用户</td>
    </tr>
    <tr class="tdbg"> 
      <td width="85%" bgcolor="#FFFFFF"> 
        <input type="radio" name="incepttype" value="2">
        按用户姓名发送邮件&nbsp;
        <input type="text" name="inceptname" size="35">
        多个用户名请用<font color="#0000FF">英文的逗号</font>分隔。</td>
    </tr>
    <tr class="tdbg"> 
      <td width="85%" bgcolor="#FFFFFF"> 
        <input name="incepttype" type="radio" value="3" checked>
        按用户Email发送邮件 		
        <input name="inceptemail" type="text" value="<%=Email%>" size="35">
        多个用户Email请用<font color="#0000FF">英文的逗号</font>分隔。</td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tdbg"> 
      <td width="15%" align="right" bgcolor="#A4B6D7">邮件主题：</td>
      <td width="85%"> 
        <input type=text name=subject size=64>
      </td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tdbg"> 
      <td align="right" bgcolor="#A4B6D7">邮件内容：</td>
      <td> 
        <textarea cols=80 rows=8 name="content"></textarea>
      </td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tdbg"> 
      <td width="15%" align="right" bgcolor="#A4B6D7">发件人：</td>
      <td width="85%"> 
        <input type="text" name="sendername" size="64" value="<%=SiteName%>">
        </td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tdbg"> 
      <td width="15%" align="right" bgcolor="#A4B6D7">发件人Email：</td>
      <td width="85%"> 
        <input type="text" name="senderemail" size="64" value="<%=WebMasterEmail%>">
        </td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tdbg"> 
      <td align="right" bgcolor="#A4B6D7">邮件优先级：</td>
      <td> 
        <input type="radio" name="Priority" value="1">
        高 
        <input type="radio" name="Priority" value="3" checked>
        普通 
        <input type="radio" name="Priority" value="5">
        低</td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tdbg"> 
      <td width="15%" align="right" bgcolor="#A4B6D7">注意事项：</td>
      <td width="85%"> 
        <%
		If JMObjInstalled=false Then
			Response.Write "<b><font color=red>对不起，因为服务器不支持 JMail组件! 所以不能使用本功能。</font></b>"
		else
			Response.Write "信息将发送到所有注册时完整填写了信箱的用户，邮件列表的使用将消耗大量的服务器资源，请慎重使用。"
		End If
		%>
      </td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tdbg"> 
      <td colspan=2 align=center> 
        <input name="Action" type="hidden" id="Action" value="Send">
        <input name="Submit" type="submit" id="Submit" value=" 发 送 " <% If JMObjInstalled=false Then response.write "disabled" end if%>>
        &nbsp; 
        <input  name="Reset" type="reset" id="Reset2" value=" 清 除 ">
      </td>
    </tr>
  </table>
</form>
<%
end sub

sub SendMaillist()
	dim Sendername,Senderemail,Subject,Content,Priority,InceptType,InceptName,InceptEmail,i,j
	Sendername=trim(request("sendername"))
	Senderemail=trim(request("senderemail"))
	Subject=trim(request("Subject"))
	Content=trim(request("Content"))
	Priority=trim(request("Priority"))
	if Sendername="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>发件人不能为空！</li>"
	end if
	if Senderemail="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>发件人Email不能为空！</li>"
	end if
	if Subject="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>邮件主题不能为空！</li>"
	end if
	if Content="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>邮件内容不能为空！</li>"
	end if
	if Priority="" then
		Priority=3
	end if
	
	InceptType=Clng(request("incepttype"))
	sql="select UserName,Email from [User] " 
	if InceptType=1 then		
		sql=sql & " where Email like '%@%'"		
	elseif InceptType=2 then
		InceptName=replace(replace(replace(replace(request("inceptname")," ",""),"'",""),chr(34),""),"|","','")
		sql=sql & " where UserName in ('" & InceptName & "') and Email like '%@%'"
	elseif InceptType=3 then
		InceptEmail=replace(replace(replace(replace(request("inceptemail")," ",""),"'",""),chr(34),""),"|","','")
		sql=sql & " where Email in ('" & InceptEmail & "')"
	end if
	
	
	if FoundErr=True then
		exit sub
	end if

	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
	if rs.bof and rs.eof then
	  if InceptType=3  then
	       if IsValidEmail(Email)=True then			   
				ErrMsg=SendMail(Email,mane,Subject,Content,Sendername,Senderemail,Priority)
				if ErrMsg<>"" then
					FoundErr=True					
				end if
				i=1			
		    end if
	  else	  
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>暂时没有用户注册！</li>"
	  end if	
	else	    
		response.write "<li>正在发送中，请等待 "
		do while not rs.eof		
			if IsValidEmail(rs("Email"))=True then			   
				ErrMsg=SendMail(rs("Email"),rs("UserName"),Subject,Content,Sendername,Senderemail,Priority)
				if ErrMsg<>"" then
					FoundErr=True
					exit sub
				end if
				i=i+1
				response.write "."
			else
				j=j+1				
			end if
			rs.movenext			
		loop
		response.write "<BR><li>成功发送邮件："&i&"封"
		if j>0 then response.write "<BR><li>未发送邮件："&j&"封（邮件地址错误）。" end if
	end if
	rs.close
	set rs=nothing
	call CloseConn()
end sub

sub ExportMail()
%>
<form method="POST" action="AdminMaillist.asp?Action=DoExport">
  <table width="90%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#000000" Class="border">
    <tr bgcolor="#FFFFFF" class="title"> 
      <td  height="28" colspan=2 align=center class="back_southidc"><b> 邮件列表批量导出到数据库</b></td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tdbg"> 
      <td width="24%" height="80" align="right" bgcolor="#A4B6D7">导出邮件列表到数据库：</td>
      <td width="76%" height="80"> 
        <input name="ExportType" type="hidden" id="ExportType" value="1">
        &nbsp;&nbsp;<font color=blue>导出</font>&nbsp;&nbsp; 
        <select name="UserType" id="UserType">
          <option value="0" selected>全部用户</option>          
        </select>
        &nbsp;<font color=blue>到</font>&nbsp;
        <input name="ExportFileName" type="text" id="ExportFileName" value="Emaillist.mdb" size="30" maxlength="200">
        <input name="Submit1" type="submit" id="Submit1" value="开始">
      </td>
    </tr>
  </table>
</form>
<form method="POST" action="AdminMaillist.asp?Action=DoExport">
  <table width="90%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#000000" Class="border">
    <tr bgcolor="#FFFFFF" class="title"> 
      <td height="28" colspan=2 align=center class="back_southidc"><b>邮件列表批量导出到文本</b></td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tdbg"> 
      <td width="24%" height="80" align="right" bgcolor="#A4B6D7">导出邮件列表到数据库：</td>
      <td width="76%" height="80"> 
        <input name="ExportType" type="hidden" id="ExportType" value="2">
        &nbsp;&nbsp;<font color=blue>导出</font>&nbsp;&nbsp; 
        <select name="UserType" id="UserType">
          <option value="0" selected>全部用户</option>        
        </select>
        &nbsp;<font color=blue>到</font>&nbsp;
		<input name="ExportFileName" type="text" id="ExportFileName" value="Emaillist.txt" size="30" maxlength="200">
        <input type="submit" name="Submit2" value="开始" <%If FSObjInstalled=false Then response.Write "disabled"%>> 
		<%
			If FSObjInstalled=false Then
				Response.Write "<font color=red>你的服务器不支持 FSO! 不能使用此功能。</font>"
			end if
		%>
      </td>
    </tr>
  </table>
</form>
<%
end sub

sub DoExportList()
   
	dim ExportType,UserType,ExportFileName,strResult,i
	ExportType=Clng(trim(Request("ExportType")))	
	ExportFileName=trim(request("ExportFileName"))
	if ExportFileName="" then
		FoundErr=True
		if ExportType=1 then
			ErrMsg=ErrMsg & "<br><li>请输入要导出的数据库文件名！</li>"
		else
			ErrMsg=ErrMsg & "<br><li>请输入要导出的文本文件名！</li>"
		end if
	else
		ExportFileName=replace(replace(ExportFileName,"'",""),chr(34),"")
	end if	
	
	set rs=server.createobject("adodb.recordset")	
	sql="select Email from [User] where Email like '%@%'"		
	rs.open sql,conn,1,1	
	i=0
	select case ExportType
	case 1
	    	    	
		dim tconn,tconnstr	
		Set tconn = Server.CreateObject("ADODB.Connection")		
		tconnstr="DBQ="+server.mappath(""&ExportFileName&"")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
		tconn.Open tconnstr
		do while not rs.eof
			tconn.execute("insert into [user] (email) values ('"&rs(0)&"')")			
			rs.movenext
			i=i+1
		loop
		tConn.close
		Set tconn = Nothing
		strResult="操作成功：共导出 "& i &" 个用户Email地址到数据库 "&tdb&"。<a href="&ExportFileName&">点击这里将数据库下载回本地</a>"
	case 2
		dim fso,filepath,writefile
	
		Set fso = CreateObject("Scripting.FileSystemObject")
		Application.lock
		filepath=Server.MapPath(""&ExportFileName&"")
		Set Writefile = fso.CreateTextFile(filepath,true)
		do while not rs.eof
			Writefile.WriteLine rs(0)
			rs.movenext
			i=i+1
		loop
		Writefile.close
		Application.unlock
		set fso=nothing
		strResult="操作成功：共导出 " & i & " 个用户Email地址到"&ExportFileName&"文件。<a href="&ttxt&">点击这里将文件下载回本地</a>)"
	end select
	rs.close
	set rs=nothing
%>
<table width="90%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#000000" class="border">
  <tr class="title"> 
    <td height="28" align=center bgcolor="#FFFFFF" class="back_southidc"><b>邮件列表批量导出反馈信息</b></td>
  </tr>
  <tr class="tdbg"> 
    <td height="100" align="center" bgcolor="#FFFFFF">
<%response.write strResult%></td>
  </tr>
</table>
<%
end sub
%>
</body>
</html>