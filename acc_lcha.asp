<!--#include file="Inc/Sysconn.asp" -->
<style type="text/css">
<!--
.btn_2k3 {
	BORDER-RIGHT: #002D96 1px solid;
	PADDING-RIGHT: 2px;
	BORDER-TOP: #002D96 1px solid;
	PADDING-LEFT: 2px;
	FONT-SIZE: 12px;
FILTER: progid:DXImageTransform.Microsoft.Gradient(GradientType=0, StartColorStr=#FFFFFF, EndColorStr=#9DBCEA);
	BORDER-LEFT: #002D96 1px solid;
	CURSOR: hand;
	COLOR: black;
	PADDING-TOP: 2px;
	BORDER-BOTTOM: #002D96 1px solid
}
-->
</style>
<%
Dim iid
iid=Request("id")
set rs=server.createobject("adodb.recordset")
strSql="select * from [fenz_list] where ID=" & Clng(iid)
rs.open strSql,conn,1,1
%>
<HTML>
<HEAD>
<TITLE>挤模调试部考核系统</TITLE>

</head>
<BODY bgColor=#95db98 topmargin=15 leftmargin=15 >
<form name="form1" method="post" action="">
  <table width=100% border="0" cellpadding="0" cellspacing="2">
    <tr>
      <td><FIELDSET align=left>
        <LEGEND align=left>修改&nbsp;<%=Rs("zrr")%>&nbsp;实际得分</LEGEND>
        <TABLE border="0" cellpadding="0" cellspacing="0">
          <TR>
            <TD align="center">应得分值：
              <INPUT name="ydf" type="text" id=ydf  value=<%=Rs("ydfz")%> size=7 readonly="true">
            </td>
            <TD align="center">实际分值：
              <INPUT name="sjf" type="text" id=sjf ONKEYPRESS="event.returnValue=IsDigit();" value=<%=Rs("sjfz")%> size=7>
            </TD>
          </TR>
        </TABLE>
        </fieldset></td>
    </tr>
    <tr>
      <td height="1" colspan="2"><input type="hidden" name="id" value=<%=iid%>></td>
    </tr>
    <tr>
      <td align="center"><input name="cmdOK" class=btn_2k3 type="button" id="cmdOK" value="  确定  " onClick="OK();">
        <input name="cmdCancel" class=btn_2k3 type=button id="cmdCancel" onClick="window.close();" value='  取消  '></td>
    </tr>
  </table>
</form>
</body>
</html><%Rs.Close%>
<script language="JavaScript">
function OK(){
  var strurl=document.form1.sjf.value;
  if (strurl=="")
  {
  	alert("实际得分不能为空！");
	document.form1.sjf.focus();
	return false;
  }
  else
  {
    window.returnValue = strurl;
    window.close();
  }
}
function IsDigit()
{
  return ((event.keyCode >= 48) && (event.keyCode <= 57));
}
</script>
