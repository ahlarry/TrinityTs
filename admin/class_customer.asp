<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<!--#include file="admin.asp" -->
<%
'����Աɾ�����-----------
if  Request("del_classid")<>"" then
   'conn.execute("delete from Info where InfoClassid="&Request("del_classid")) '��ɾ����Ϣ���� 
   conn.execute("delete from customer_class where id="&Request("del_classid")) 'ɾ������
end if
'ɾ������---------------

'�������begin----
if  Request("classname")<>"" then
   sql="select * from customer_class"
   set rscust=Server.Createobject("ADODB.Recordset")
   rscust.open sql,conn,1,3
   rscust.addnew
   'rscust("paixu")=cint(Request("paixu"))
   rscust("classname")=Request("classname")
   rscust.update
   set rscust=nothing
end if
'�������end----

'�������-----------
if  Request("action")<>"" then
   sql="select * from customer_class where id="&Request("classid")
   set rscust=Server.Createobject("ADODB.Recordset")
   rscust.open sql,conn,1,3
   'if Request("paixu")<>"" then
   ' rscust("paixu")=cint(Request("paixu"))
   'end if
   if Request("class_rename")<>"" then
   rscust("infoclass")=Request("class_rename")
   end if
   'if Request("image")<>"" then
   'rscust("image")=Request("image")
   'end if
   'if Request("ishome")<>"" then
   'rscust("ishome")=True
   'else
   'rscust("ishome")=False
   'end if
   rscust.update
   set rscust=nothing
end if
'�����������---------------

'ɾ������---------------
function del_img(path,filename)
  whichfile=server.mappath(path&filename)
  Set fs = CreateObject("Scripting.FileSystemObject")
  Set thisfile = fs.GetFile(whichfile)
  thisfile.Delete True
end function
%>
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="admin.css" rel="stylesheet" type="text/css">

<script language="JavaScript">
<!--
function cform(){
 if(!confirm("��ȷ��ɾ����? ��ע��ɾ�����޷��ָ�!"))
 return false;
}//-->
</script>
</head>

<body leftmargin="0" topmargin="0">
<table width="96%" border="0" align="center" class="px14">
  <tr> 
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td align="center" bgcolor="#3366CC"><font color="#FFFFFF">�� �� �� �� �� ��</font></td>
  </tr>
</table>
<br>
<table width="96%" border="0" align="center" class="px14">
  <tr valign="top" bgcolor="#EFEFEF"> 
    <td width="33%"><table width="100%" border="0">
        <tr> 
          <td align="center" bgcolor="#CCCCCC">ɾ�����</td>
        </tr>
        <tr> 
          <td height="100" valign="top"> <form name="form1" method="post" onsubmit="return cform()">
              <table width="100%" border="0">
                <tr> 
                  <td width="34%" align="right">ѡ�����</td>
                  <td width="66%"> <% 
	             sql="select * from customer_class"
	             set rscust=Server.Createobject("ADODB.Recordset")
                 rscust.open sql,conn,1,3%>
				  <select name="del_classid" class="input">
                      <%do while not rscust.eof  %>
                      <option value="<%= rscust("id") %>"><%= rscust("classname")  %></option>
                      <%  rscust.movenext 
					loop
					set rscust=nothing%>
                    </select></td>
                </tr>
                <tr> 
                  <td align="right">&nbsp;</td>
                  <td><input type="submit" name="Submit" value="ɾ�����"></td>
                </tr>
                <tr> 
                  <td align="right">&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
              </table>
            </form></td>
        </tr>
      </table></td>
    <td width="33%"><table width="100%" border="0">
        <tr> 
          <td align="center" bgcolor="#CCCCCC">�������</td>
        </tr>
        <tr> 
          <td height="100" valign="top"> <form name="form2" method="post" action="">
              <table width="100%" border="0">
                <tr> 
                  <td width="25%" align="right">����</td>
                  <td width="75%"><input name="classname" type="text" class="input" id="classname" maxlength="50"> 
                  </td>
                </tr>
                <tr> 
                  <td align="right">&nbsp;</td>
                  <td><input type="submit" name="Submit2" value="�������"></td>
                </tr>
              </table>
            </form></td>
        </tr>
      </table></td>
  </tr>
</table>
</body>
</html>
<% 
conn.close
set conn=nothing %>