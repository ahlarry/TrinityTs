<%@language=VBScript codepage=936 %>
<!--#include file="Conn.asp"-->
<!--#include file="Admin.asp"-->
<!--#include file="Inc/Function.asp"-->
<%
dim SmallClassID,Action,BigClassName, SmallClassName, OldSmallClassName,rs,FoundErr,ErrMsg
SmallClassID=trim(Request("SmallClassID"))
Action=trim(Request("Action"))
BigClassName=trim(Request.form("BigClassName"))
SmallClassName=trim(Request.form("SmallClassName"))
OldSmallClassName=trim(request.form("OldSmallClassName"))
if SmallClassID="" then
	response.Redirect("ClassManage.asp")
end if
Set rs=Server.CreateObject("Adodb.RecordSet")
rs.Open "Select * from SmallClass where SmallClassID="&SmallClassID&"",conn,1,3
if rs.Bof or rs.EOF then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>�˲�ƷС�಻���ڣ�</li>"
else
	if Action="Modify" then
		if SmallClassName="" then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>��ƷС��������Ϊ�գ�</li>"
		end if
		if FoundErr<>True then
			rs("SmallClassName")=SmallClassName
     		rs.update
			rs.Close
    	 	set rs=Nothing
			if SmallClassName<>OldSmallClassName then
				conn.execute "Update Products set SmallClassName='" & SmallClassName & "' where BigClassName='" & BigClassName & "' and SmallClassName='" & OldSmallClassName & "'"
	     	end if	
			call CloseConn()
    	 	Response.Redirect "ClassManage.asp"
		end if
	end if
	if FoundErr=True then
		call WriteErrMsg()
	else
%>
<!-- #include file="Inc/Head.asp" -->
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td align="center" valign="top"> <br> <br>
      <strong>�� Ʒ �� �� �� ��</strong> <br> <br> <br> <form name="form1" method="post" action="ClassModifySmall.asp">
        <p>&nbsp;</p>
        <table width="350" border="0" align="center" cellpadding="0" cellspacing="2" bgcolor="#ECF5FF" class="border">
          <tr bgcolor="#A4B6D7" class="title"> 
            <td height="25" colspan="2" align="center"><strong>�޸�С������</strong></td>
          </tr>
          <tr class="tdbg"> 
            <td width="126" height="22" bgcolor="#A4B6D7"><strong>�������ࣺ</strong></td>
            <td width="204"><%=rs("BigClassName")%> <input name="SmallClassID" type="hidden" id="SmallClassID" value="<%=rs("SmallClassID")%>"> 
              <input name="OldSmallClassName" type="hidden" id="OldSmallClassName" value="<%=rs("SmallClassName")%>"> 
              <input name="BigClassName" type="hidden" id="BigClassName" value="<%=rs("BigClassName")%>">
            </td>
          </tr>
          <tr class="tdbg"> 
            <td height="22" bgcolor="#A4B6D7"><strong>С�����ƣ�</strong></td>
            <td> 
              <input name="SmallClassName" type="text" id="SmallClassName" value="<%=rs("SmallClassName")%>" size="20" maxlength="30">
            </td>
          </tr>
          <tr class="tdbg"> 
            <td height="22" align="center" bgcolor="#A4B6D7">&nbsp;</td>
            <td align="center"> 
              <div align="left"> 
                <input name="Action" type="hidden" id="Action4" value="Modify">
                <input name="Save" type="submit" id="Save" value=" �� �� ">
              </div></td>
          </tr>
        </table>
      </form></td>
  </tr>
</table>
<!-- #include file="Inc/Foot.asp" -->
<%
	end if
end if
rs.close
set rs=nothing
call CloseConn()
%>