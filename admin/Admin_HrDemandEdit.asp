<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<%
ID=trim(Request("ID"))
if Request.QueryString("mark")="southidc" then
HrName=Trim(Request("HrName"))
If HrName="" Then
response.write "��Ƹ������Ϊ��!!<a href=""javascript:history.go(-1)""><font color=""#FF0000"">�뷵������</font></a>"
response.end
end if
HrRequireNum=Trim(Request("HrRequireNum"))
If HrRequireNum="" Then
response.write "��Ƹ��������Ϊ��!!<a href=""javascript:history.go(-1)""><font color=""#FF0000"">�뷵������</font></a>"
response.end
end if
HrAddress=Trim(Request("HrAddress"))
If HrAddress="" Then
response.write "�����ص㲻��Ϊ��!!<a href=""javascript:history.go(-1)""><font color=""#FF0000"">�뷵������</font></a>"
response.end
end if
HrSalary=Trim(Request("HrSalary"))
HrDate=Trim(Request("HrDate"))
If HrDate="" Then
response.write "�������ڲ���Ϊ��!!<a href=""javascript:history.go(-1)""><font color=""#FF0000"">�뷵������</font></a>"
response.end
end if
HrValidDate=Trim(Request("HrValidDate"))
If HrValidDate="" Then
response.write "��Ч���޲���Ϊ��!!<a href=""javascript:history.go(-1)""><font color=""#FF0000"">�뷵������</font></a>"
response.end
end if
For i = 1 To Request.Form("Content").Count
  About = Content & Request.Form("Content")(i)
Next
set rs=server.createobject("adodb.recordset")
sql="select * from HrDemand where id="  & Clng(ID)
rs.open sql,conn,3,3
	rs("HrDetail")=About
	rs("HrName")=HrName
	rs("HrRequireNum")=HrRequireNum
	rs("HrAddress")=HrAddress
	rs("HrSalary")=HrSalary
	rs("HrDate")=HrDate
	rs("HrValidDate")=HrValidDate
	rs.update
	rs.close
	response.redirect "Admin_HrDemand.asp"
end if
sql="select * from HrDemand where id="  & Clng(ID)
Set rs_home= Server.CreateObject("ADODB.Recordset")
rs_home.open sql,conn,1,1
%>
<!-- #include file="Inc/Head.asp" -->
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td align="center" valign="top"><br> 
      <table width="650" border="0" cellpadding="2" cellspacing="1" class="table_southidc">
        <tr> 
          <td class="back_southidc" height="25"> <div align="center"><strong>�޸���Ƹ<br>
              </strong></div></td>
        </tr>
        <tr> 
          <form method="post" action="Admin_HrDemandEdit.asp?mark=southidc">
            <input type=hidden name=id value="<%=rs_home("id")%>">
            <td><div align="center"> 
                <table width="100%" border="0" cellspacing="1" cellpadding="0">
                  <tr bgcolor="#ECF5FF"> 
                    <td width="19%" height="25"> 
                      <div align="center">��Ƹ����</div></td>
                    <td width="81%"> 
                      <input name="HrName" type="text" id="HrName" value="<%=rs_home("HrName")%>" size="30" maxlength="30">
                    </td>
                  </tr>
                  <tr bgcolor="#ECF5FF"> 
                    <td height="22"> 
                      <div align="center">��Ƹ����</div></td>
                    <td bgcolor="#ECF5FF"> 
                      <input name="HrRequireNum" type="text" id="HrRequireNum" value="<%=rs_home("HrRequireNum")%>" size="5" maxlength="30">
                      ��</td>
                  </tr>
                  <tr bgcolor="#ECF5FF"> 
                    <td height="22"> 
                      <div align="center">�����ص�</div></td>
                    <td height="-4" bgcolor="#ECF5FF"> 
                      <input name="HrAddress" type="text" id="HrAddress" value="<%=rs_home("HrAddress")%>" size="50" maxlength="60">
                    </td>
                  </tr>
                  <tr bgcolor="#ECF5FF"> 
                    <td height="22"> 
                      <div align="center">���ʴ���</div></td>
                    <td height="-1" bgcolor="#ECF5FF"> 
                      <input name="HrSalary" type="text" id="Daiy" value="<%=rs_home("HrSalary")%>" size="20" maxlength="50">
                    </td>
                  </tr>
                  <tr bgcolor="#ECF5FF"> 
                    <td height="22"> 
                      <div align="center">��������</div></td>
                    <td height="-3"> <input name="HrDate" type="text" id="HrDate" value="<%=rs_home("HrDate")%>" size="12" maxlength="50"></td>
                  </tr>
                  <tr bgcolor="#ECF5FF"> 
                    <td height="22"> 
                      <div align="center">��Ч����</div></td>
                    <td height="0" bgcolor="#ECF5FF"> 
                      <input name="HrValidDate" type="text" id="HrValidDate" value="<%=rs_home("HrValidDate")%>" size="5" maxlength="30">
                      ��</td>
                  </tr>
                  <tr bgcolor="#ECF5FF"> 
                    <td height="22"> 
                      <div align="center">��ƸҪ��</div></td>
                    <td height="10">                    
					   <textarea name="Content"  style="display:none"><%=rs_home("HrDetail")%></textarea>
					  <iframe ID="eWebEditor1" src="Southidceditor/ewebeditor.asp?id=content&style=southidc" frameborder="0" scrolling="no" width="550" HEIGHT="400"></iframe> 
					  </td>
                  </tr>
                  <tr bgcolor="#ECF5FF"> 
                    <td height="25" colspan="2"> 
                      <div align="center"> 
                        <input name="submit" type="submit" value="ȷ��">
                        &nbsp; 
                        <input name="reset" type="reset" value="ȡ��">
                      </div></td>
                  </tr>
                </table>
              </div></td>
          </form>
        </tr>
      </table></td>
  </tr>
</table>
<!-- #include file="Inc/Foot.asp" -->