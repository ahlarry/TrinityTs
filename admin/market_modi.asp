<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp"-->
<!--#include file="Admin.asp"-->
<!-- #include file="Inc/Head.asp" -->
<% set rs=server.createobject("adodb.recordset") 
rs.open "select * from market where id="&Request.QueryString("id"),conn,1,3%>

<div align="center">
  <table width="96%" border="0" align="center" class="px14">
    <tr> 
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td height="30" align="center" class="back_southidc">�޸���������</td>
    </tr>
  </table>
  <table width="96%" border="0">
    <tr> 
      <td valign="top"> 
<SCRIPT language=javascript>
function  save_onclick()
{
  document.form2.Content.value=editor.HtmlEdit.document.body.innerHTML; 
  var strTemp = document.form2.province.value;  
  if (strTemp.length == 0 )
  {
      alert("��ѡ�����ڵ�����");
      document.form2.province.focus();
      return false;
  }
  
  var strTemp = document.form2.title.value;  
  if (strTemp.length == 0 )
  {
      alert("�����������������ƣ�");
      document.form2.title.focus();
      return false;
  }

  var strTemp = document.form2.Content.value;
  if (strTemp.length == 0 )
  {
      alert("����д���ݣ�");
      return false;
  }
  return true;
}

</SCRIPT>
        <table width="100%" border="0" cellspacing="1" class="table_southidc">
          <form action="market_write.asp" method="post" name="form2" onsubmit="return save_onclick()">
            <tr bgcolor="#ECF5FF"> 
              <td width="18%" align="right">���ڵ���:</td>
              <td width="82%">
                <select name="province" class="input" id="province">
                  <option  value="">��ѡ�����ڵ���</option>
                  <%set rs1=server.createobject("adodb.recordset") 
				  rs1.open "select * from province",conn,1,3
       				do while not rs1.eof
					%>
                  <option value="<%=rs1("id")%>" <% If rs1("id")=rs("province") Then%>selected<% End If %>><%=rs1("province")%></option>
                  <%rs1.movenext
				   loop
				   rs1.close%>
                </select>
              <a href="product_catalog.asp"></a> </td>
            </tr>
            <tr bgcolor="#ECF5FF"> 
              <td align="right">��������:</td>
              <td><input name="title" type="text" class="input" id="title" value="<%=rs("title")%>" size="60" maxlength="100">
              (100�ַ�����)</td>
            </tr>
            <tr bgcolor="#ECF5FF"> 
              <td align="right">��ϵ��ʽ�����:<br>
			  <font color="#FF0000">�����밴Shift+Enter�� <br>
			  ����һ���밴Enter ��<br>
			  ֧��UBB��ǩ ��</font></td>
              <td valign="top"><textarea name="Content" style="display:none"></textarea>
	  <iframe ID="editor" src="Editor.asp?Action=Modify&ArticleID=<%=rs("id")%>&Tb=market" frameborder=1 scrolling=no width="620" height="405"></iframe></td>
            </tr>
            <tr bgcolor="#ECF5FF"> 
              <td align="right">
			  <input name="id" type="hidden" id="id" value="<%=rs("id")%>"></td>
              <td><input type="submit" name="Submit2" value="�޸�">���������� 
              <input type="button" name="Submit3" value="����" onClick="history.go(-1);"></td>
            </tr>
          </form>
        </table></td>
    </tr>
  </table>
</div>
<!-- #include file="Inc/Foot.asp" -->
