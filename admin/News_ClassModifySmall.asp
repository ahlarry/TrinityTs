<%@language=VBScript codepage=936 %>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!-- #include file="Inc/Head.asp" -->
<%
dim SmallClassID,Action,BigClassName, SmallClassName, OldSmallClassName,rs,FoundErr,ErrMsg
SmallClassID=trim(Request("SmallClassID"))
Action=trim(Request("Action"))
BigClassName=trim(Request.form("BigClassName"))
SmallClassName=trim(Request.form("SmallClassName"))
OldSmallClassName=trim(request.form("OldSmallClassName"))
if SmallClassID="" then
	response.Redirect("News_ClassManage.asp")
end if
Set rs=Server.CreateObject("Adodb.RecordSet")
rs.Open "Select * from SmallClass_New where SmallClassID="&SmallClassID&"",conn,1,3
if rs.Bof or rs.EOF then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>������С�಻���ڣ�</li>"
else
	if Action="Modify" then
		if SmallClassName="" then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>����С��������Ϊ�գ�</li>"
		end if
		if FoundErr<>True then
			rs("SmallClassName")=SmallClassName
     		rs.update
			rs.Close
    	 	set rs=Nothing
			if SmallClassName<>OldSmallClassName then
				conn.execute "Update NEWS set SmallClassName='" & SmallClassName & "' where BigClassName='" & BigClassName & "' and SmallClassName='" & OldSmallClassName & "'"
	     	end if	
			call CloseConn()
    	 	Response.Redirect "News_ClassManage.asp"
		end if
	end if
	if FoundErr=True then
		call WriteErrMsg()
	else
%>
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td align="center" valign="top"><table width="80%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td align="center" valign="top"> <br> <strong>�� �� �� �� �� ��</strong> <br> 
            <form name="form1" method="post" action="News_ClassModifySmall.asp">
              <p>&nbsp;</p>
              <table width="350" border="0" align="center" cellpadding="0" cellspacing="2" class="border">
                <tr bgcolor="#A4B6D7" class="title"> 
                  <td height="25" colspan="2" align="center"><strong>�޸�С������</strong></td>
                </tr>
                <tr class="tdbg"> 
                  <td width="126" height="22" bgcolor="#C0C0C0"><strong>�������ࣺ</strong></td>
                  <td width="204" bgcolor="#E3E3E3"><%=rs("BigClassName")%> <input name="SmallClassID" type="hidden" id="SmallClassID" value="<%=rs("SmallClassID")%>"> 
                    <input name="OldSmallClassName" type="hidden" id="OldSmallClassName" value="<%=rs("SmallClassName")%>"> 
                    <input name="BigClassName" type="hidden" id="BigClassName" value="<%=rs("BigClassName")%>"></td>
                </tr>
                <tr class="tdbg"> 
                  <td height="22" bgcolor="#C0C0C0"><strong>С�����ƣ�</strong></td>
                  <td bgcolor="#E3E3E3"> <input name="SmallClassName" type="text" id="SmallClassName" value="<%=rs("SmallClassName")%>" size="20" maxlength="30"></td>
                </tr>
                <tr class="tdbg"> 
                  <td height="22" align="center" bgcolor="#C0C0C0">&nbsp;</td>
                  <td align="center" bgcolor="#E3E3E3"> <div align="left"> 
                      <input name="Action" type="hidden" id="Action4" value="Modify">
                      <input name="Save" type="submit" id="Save" value=" �� �� ">
                    </div></td>
                </tr>
              </table>
            </form></td>
        </tr>
      </table>
      <%
end if
end if
rs.close
set rs=nothing
%> </td>
  </tr>
</table>
<!-- #include file="Inc/Foot.asp" -->