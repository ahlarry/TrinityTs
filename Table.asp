<html>
<head>
<title>�������</title>
<meta name="Author" content="heweiqun">
<meta name="Contact" content="webmaster@wrtx.cn">
<meta name="Copyright" content="southidc.net">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
body, a, table, div, span, td, th, input, select{font:9pt;font-family: "����", Verdana, Arial, Helvetica, sans-serif;}
.text{border:1px solid #aaaaaa;}
.button{height:18;border:1px ridge #aaaaaa;background-color:aaaaaa;color:ffffff}

</style>
<script language="JavaScript">
function checkchange()
{
if  (checkbox.checked==true){
    width.disabled=false;
	}
else{	
   width.disabled=true;
   }
}	 


function table(){
var tablehtml="";
tablehtml=tablehtml+"<table ";
if (align=="left"||align=="center"||align=="right")
 tablehtml=tablehtml+"align="+align.value+" ";
if (border!="")
 tablehtml=tablehtml+"border="+border.value+" "
 else
 tablehtml=tablehtml+"border=1"+" "
if(cellpad.value!="")
  tablehtml=tablehtml+"cellpadding="+cellpad.value+" " 
 if(cellspace.value!="")
  tablehtml=tablehtml+"cellspacing="+cellspace.value+" "
  if(bordercolor.value!="")
  tablehtml=tablehtml+"bordercolor="+bordercolor.value+" " 
  if(bgcolor.value!="")
  tablehtml=tablehtml+"bgcolor="+bgcolor.value+" "
  if  (checkbox.checked==true)
    {
	if(width.value!="")
    { tablehtml=tablehtml+"width="+width.value
	     tablehtml=tablehtml+WidthUnit.value+""
	}	 
	} 	
	 tablehtml=tablehtml+">"  
for (i=1;i<=rows.value;i++)
{
tablehtml=tablehtml+"<tr>";
 for (j=1;j<=columns.value;j++)
   tablehtml=tablehtml+"<td></td>";
 }  
 tablehtml= tablehtml+"</table>"; 
window.returnValue = tablehtml;
 window.close();
 }
</script>

</head>
<body bgcolor=menu topmargin="5" leftmargin="0">

  
<table width="347" border="0" cellspacing="0" cellpadding="0" align="center" height="258">
  <tr> 
    <td height="20" width="84"> <b>��С</b> </td>
    <td height="20" colspan="4"> <hr size="2"> </td>
  </tr>
  <tr> 
    <td height="20" colspan="2" align="right"><font size="2">��������</font></td>
    <td height="20" width="81"> <input type="text" name="rows" size="8" maxlength="2"> 
    </td>
    <td width="87" height="0" align="right">������</td>
    <td height="0" width="94"> <input type="text" name="columns" size="8" maxlength="2"> 
    </td>
  </tr>
  <tr> 
    <td height="20" colspan="2"> <b><font size="2">����</font></b> </td>
    <td height="20" colspan="3"> <hr align="right" size="2"> </td>
  </tr>
  <tr> 
    <td height="20" colspan="2" align="right"><font size="2">�����뷽ʽ��</font></td>
    <td height="20" width="81"> <select name="align">
        <option selected>Default</option>
        <option value="left">Left</option>
        <option value="center">Center</option>
        <option value="right">Right</option>
      </select> </td>
    <td height="0" colspan="2"> &nbsp;&nbsp;&nbsp;&nbsp; <input name="checkbox" type="checkbox" onClick="checkchange();" value="1">
      ָ������ </td>
  </tr>
  <tr> 
    <td height="20" colspan="2" align="right">���߿��ϸ��</td>
    <td height="20" width="81"> <input type="text" name="border" size="8" maxlength="2"> 
    </td>
    <td height="0" align="right" nowrap><input name="width" type="text" value="100" size="5" maxlength="3" disabled></td>
    <td height="0" align="center" nowrap><select name="WidthUnit">
        <option>����</option>
        <option value="%" selected>�ٷֱ�</option>
      </select></td>
  </tr>
  <tr> 
    <td height="20" colspan="2" align="right">����Ԫ��߾ࣺ</td>
    <td height="20" width="81"> <input type="text" name="cellpad" size="8" maxlength="2"> 
    </td>
    <td height="0" align="right" nowrap>&nbsp;</td>
    <td height="0" align="center" nowrap>&nbsp;</td>
  </tr>
  <tr> 
    <td height="20" colspan="2" align="right">����Ԫ���ࣺ</td>
    <td height="20" width="81"> <input type="text" name="cellspace" size="8" maxlength="2"> 
    </td>
    <td height="0" width="87">&nbsp;</td>
    <td height="0" width="94">&nbsp;</td>
  </tr>
  <tr> 
    <td height="19" colspan="2"> <b><font size="2">��ɫ</font></b> </td>
    <td height="19" colspan="3"> <hr align="right" size="2"> </td>
  </tr>
  <tr> 
    <td height="20" colspan="2" align="right"><font size="2">���߿���ɫ��</font></td>
    <td height="20" width="81"> <input type="text" name="bordercolor" size="8" maxlength="7" value="#000000"> 
    </td>
    <td width="87" height="0" align="right"><font size="2">��Ԫ�񱳾���</font></td>
    <td height="0" width="94"> <input type="text" name="bgcolor" size="8" maxlength="7" value="#ffffff"> 
    </td>
  </tr>
  <tr> 
    <td height="21" colspan="5"> <div align="right"><br>
        <input class=button type=submit onClick='table();' value="ȷ����" name="submit">
        <input class=button type=button onClick='window.close();' value="ȡ����" name="button">
      </div></td>
  </tr>
</table>

<tr>
  <td>&nbsp; </td>
</tr>


</body>
</html>