<%@language=VBScript codepage=936 %>
<!--#include file="Conn.asp"-->
<!--#include file="Admin.asp"-->
<!-- #include file="Inc/Head.asp" -->
<%
dim Action,BigClassName,SmallClassName,rs,FoundErr,ErrMsg
Action=trim(Request("Action"))
BigClassName=trim(request("BigClassName"))
SmallClassName=trim(request("SmallClassName"))
if Action="Add" then
	if BigClassName="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>下载大类名不能为空！</li>"
	end if
	if SmallClassName="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>下载小类名不能为空！</li>"
	end if
	if FoundErr<>True then
		Set rs=Server.CreateObject("Adodb.RecordSet")
		rs.open "Select * From SmallClass_down Where BigClassName='" & BigClassName & "' AND SmallClassName='" & SmallClassName & "'",conn,1,3
		if not rs.EOF then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>“" & BigClassName & "”中已经存在下载小类“" & SmallClassName & "”！</li>"
		else
     		rs.addnew
			rs("BigClassName")=BigClassName
    	 	rs("SmallClassName")=SmallClassName
     		rs.update
	     	rs.Close
    	 	set rs=Nothing
     		call CloseConn()
			Response.Redirect "Down_ClassManage.asp"  
		end if
	end if
end if
if FoundErr=True then
	call WriteErrMsg()
else
%>
<script language="JavaScript" type="text/JavaScript">
function checkSmall()
{
  if (document.form2.BigClassName.value=="")
  {
    alert("请先添加大类名称！");
	document.form1.BigClassName.focus();
	return false;
  }

  if (document.form2.SmallClassName.value=="")
  {
    alert("小类名称不能为空！");
	document.form2.SmallClassName.focus();
	return false;
  }
}
</script>
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td align="center" valign="top"> <table width="80%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td align="center" valign="top"> <br>
            <strong>下 载 类 别 设 置</strong> <br> <form name="form2" method="post" action="Down_ClassAddSmall.asp" onsubmit="return checkSmall()">
              <table width="350" border="0" align="center" cellpadding="0" cellspacing="2" bgcolor="#ECF5FF" class="border">
                <tr bgcolor="#A4B6D7" class="title"> 
                  <td height="25" colspan="2" align="center"><strong>下载小类</strong></td>
                </tr>
                <tr class="tdbg"> 
                  <td width="126" height="22" bgcolor="#A4B6D7"><strong>所属大类：</strong></td>
                  <td width="218"> 
                    <select name="BigClassName">
                      <%
	dim rsBigClass
	set rsBigClass=server.CreateObject("adodb.recordset")
	rsBigClass.open "Select * From BigClass_down",conn,1,1
	if rsBigClass.bof and rsBigClass.bof then
		response.write "<option>请先添加下载大类</option>"
	else
		do while not rsBigClass.eof
			if rsBigClass("BigClassName")=BigClassName then
				response.write "<option value='"& rsBigClass("BigClassName") & "' selected>" & rsBigClass("BigClassName") & "</option>"
			else
				response.write "<option value='"& rsBigClass("BigClassName") & "'>" & rsBigClass("BigClassName") & "</option>"
			end if
			rsBigClass.movenext
		loop
	end if
	rsBigClass.close
	set rsBigClass=nothing
	%>
                    </select>
                  </td>
                </tr>
                <tr class="tdbg"> 
                  <td width="126" height="22" bgcolor="#A4B6D7"><strong>小类名称：</strong></td>
                  <td> 
                    <input name="SmallClassName" type="text" size="20" maxlength="30">
                  </td>
                </tr>
                <tr class="tdbg"> 
                  <td height="22" align="center" bgcolor="#A4B6D7">&nbsp; </td>
                  <td height="22" align="center"> 
                    <div align="left"> 
                      <input name="Action" type="hidden" id="Action3" value="Add">
                      <input name="Add" type="submit" value=" 添 加 ">
                    </div></td>
                </tr>
              </table>
            </form></td>
        </tr>
      </table>
      <%
end if
%> <br> </td>
  </tr>
</table>
<!-- #include file="Inc/Foot.asp" -->