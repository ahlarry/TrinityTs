<%@language=VBScript codepage=936 %>
<!--#include file="Conn.asp"-->
<!--#include file="Admin.asp"-->
<!--#include file="Inc/Function.asp"-->
<%
dim Action,BigClassName,SmallClassName,rs,FoundErr,ErrMsg
Action=trim(Request("Action"))
BigClassName=trim(request("BigClassName"))
SmallClassName=trim(request("SmallClassName"))
if Action="Add" then
	if BigClassName="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>��Ʒ����������Ϊ�գ�</li>"
	end if
	if SmallClassName="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>��ƷС��������Ϊ�գ�</li>"
	end if
	if FoundErr<>True then
		Set rs=Server.CreateObject("Adodb.RecordSet")
		rs.open "Select * From SmallClass Where BigClassName='" & BigClassName & "' AND SmallClassName='" & SmallClassName & "'",conn,1,3
		if not rs.EOF then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>��" & BigClassName & "�����Ѿ����ڲ�ƷС�ࡰ" & SmallClassName & "����</li>"
		else
     		rs.addnew
			rs("BigClassName")=BigClassName
    	 	rs("SmallClassName")=SmallClassName
     		rs.update
	     	rs.Close
    	 	set rs=Nothing
     		call CloseConn()
			Response.Redirect "ClassManage.asp"  
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
    alert("�������Ӵ������ƣ�");
	document.form2.BigClassName.focus();
	return false;
  }

  if (document.form2.SmallClassName.value=="")
  {
    alert("С�����Ʋ���Ϊ�գ�");
	document.form2.SmallClassName.focus();
	return false;
  }
}
</script>
<!-- #include file="Inc/Head.asp" -->
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td align="center" valign="top"> <br> <br>
      <strong>�� Ʒ �� �� �� ��</strong> <br> <br> <br> <form name="form2" method="post" action="ClassAddSmall.asp" onsubmit="return checkSmall()">
        <table width="350" border="0" align="center" cellpadding="0" cellspacing="2" bgcolor="#ECF5FF" class="border">
          <tr class="title"> 
            <td height="25" colspan="2" align="center" bgcolor="#A4B6D7"><strong>����С��</strong></td>
          </tr>
          <tr class="tdbg"> 
            <td width="126" height="22" bgcolor="#A4B6D7"><strong>�������ࣺ</strong></td>
            <td width="218"> 
              <select name="BigClassName">
                <%
	dim rsBigClass
	set rsBigClass=server.CreateObject("adodb.recordset")
	rsBigClass.open "Select * From BigClass",conn,1,1
	if rsBigClass.bof and rsBigClass.bof then
		response.write "<option>�������Ӳ�Ʒ����</option>"
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
            <td width="126" height="22" bgcolor="#A4B6D7"><strong>С�����ƣ�</strong></td>
            <td> 
              <input name="SmallClassName" type="text" size="20" maxlength="30">
            </td>
          </tr>
          <tr class="tdbg"> 
            <td height="22" align="center" bgcolor="#A4B6D7">&nbsp; </td>
            <td height="22" align="center"> 
              <div align="left"> 
                <input name="Action" type="hidden" id="Action3" value="Add">
                <input name="Add" type="submit" value=" �� �� ">
              </div></td>
          </tr>
        </table>
      </form></td>
  </tr>
</table>
<!-- #include file="Inc/Foot.asp" -->
<%
end if
call CloseConn()
%>