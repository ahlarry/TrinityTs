<!--#include file="Inc/Sysconn.asp" -->
<!--#include file="inc/Skin_css.asp"-->
<%
Dim strzrr,strkhm,strlsh,strrwlr,iqiy,iqim,izhiy,izhim,iqitime,izhitime,dtstart,dtend
strzrr=Request("zrr")
Strrwlr=Request("rwlr")
Strkhm=Trim(Request("khm"))
strlsh=Trim(Request("lsh"))
iqiy=Trim(Request("qiy"))
iqim=Trim(Request("qim"))
izhiy=Trim(Request("zhiy"))
izhim=Trim(Request("zhim"))
iqitime="" : izhitime="" : dtstart="" : dtend=""
If iqiy<>"" and iqim<>"" then iqitime=cdate(iqiy&"��"&iqim&"��1��")
If izhiy<>"" and izhim<>"" then izhitime=cdate(izhiy&"��"&izhim&"��1��")
If iqitime<>"" and izhitime="" then
	dtstart=iqitime
	dtend=dateadd("m",1,dtstart)
	dtend=dateadd("d",-1,dtend)
End If
If iqitime="" and izhitime<>"" then
	dtstart=izhitime
	dtend=dateadd("m",1,dtstart)
	dtend=dateadd("d",-1,dtend)
End If
If iqitime<>"" and izhitime<>"" then
	If datediff("d",iqitime,izhitime)>=0 then
		dtstart=iqitime
		dtend=dateadd("m",1,izhitime)
		dtend=dateadd("d",-1,dtend)
	else
		dtstart=izhitime
		dtend=dateadd("m",1,iqitime)
		dtend=dateadd("d",-1,dtend)
	End If	
End If
%>
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
<!-- #include file="Head.asp" -->
<table width="998" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td valign="bottom" background="imgV2/left_bg1.png" class="bg_x">&nbsp;</td>
    <td width="9" align="right" valign="top" background="imgV2/left_bg2.png"><img src="imgV2/left_yy.png" width="9" height="510"></td>
    <td width="211" valign="top" background="imgV2/left_bg.gif"><!--#include file="login.asp" -->
      <!--#include file="Left_spc.asp" --></td>
    <td width="741" align="center" valign="top" background="Images_v1/rith_bg.gif" bgcolor="#F8FFE8"><table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="21"><img src="Images_v1/yx2_1.gif" width="35" height="21"></td>
          <td background="Images_v1/yx2_2.gif">&nbsp;</td>
          <td width="21" align="right"><img src="Images_v1/yx2.png" width="247" height="21"></td>
        </tr>
      </table>
      <table width="96%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td><img src="imgV2/spc_list.gif" border="0"></td>
          <td width="79%" valign="middle"><table width="100%" height="1" border="0" cellpadding="0" cellspacing="0" bgcolor="#00AD97">
              <tr>
                <td></td>
              </tr>
            </table></td>
          <td width="5%"><img src="imgV2/more2.gif" width="37" height="24" border="0"></td>
        </tr>
      </table>
<table border="0" cellpadding="2" cellspacing="0" width="100%">
  <form action="<%=Request.Servervariables("SCRIPT_NAME")%>" method="get" name="frm_searchinfo" id="frm_searchinfo" onsubmit='return true;'>
    <tr>
      <td rowspan="2">&nbsp;&nbsp;ɸѡ����:</td>
      <td><select name="qiy">
          <option value=""></option>
          <%for i=year(now)-2 to year(now)+2%>
          <option value="<%=i%>" <%If iqiy=cstr(i) then%>selected<%End If%>><%=i%></option>
          <%next%>
        </select>
        ��
        <select name="qim">
          <option value=""></option>
          <%for i=1 to 12%>
          <option value="<%=i%>" <%If iqim=cstr(i) then%>selected<%End If%>><%=i%></option>
          <%next%>
        </select>
        ��&nbsp;-&nbsp;
        <select name="zhiy">
          <option value=""></option>
          <%for i=year(now)-2 to year(now)+2%>
          <option value="<%=i%>" <%If izhiy=cstr(i) then%>selected<%End If%>><%=i%></option>
          <%next%>
        </select>
        ��
        <select name="zhim">
          <option value=""></option>
          <%for i=1 to 12%>
          <option value="<%=i%>" <%If izhim=cstr(i) then%>selected<%End If%>><%=i%></option>
          <%next%>
        </select>
        ��</td>
      <td>������:
        <input name="zrr" size="15" value="<%=strzrr%>"></td>
      <td rowspan="2" align="left"><input class=btn_2k3 type="submit" value="�� ��" /></td>
    </tr>
    <tr>
      <td>��������:
        <select name="rwlr">
          <option value="0" selected="selected">ȫ��</option>
          <option value="1" <%If strrwlr="1" Then%>selected="selected"<%End If%>>����</option>
          <option value="2" <%If strrwlr="2" Then%>selected="selected"<%End If%>>����</option>
          <option value="3" <%If strrwlr="3" Then%>selected="selected"<%End If%>>װ��</option>
        </select>
        &nbsp;&nbsp;�ͻ�����:
        <input name="khm" size="15" value="<%=strkhm%>"><input type="hidden" name="rwlx" value="<%=strrwlx%>">
 </td>
      <td>����: <input name="lsh" size="15" value="<%=strlsh%>"></td>
    </tr>
  </form>
</table>
      
      <% 
strSql=""
If strzrr <> "" Then
	strSql = "zrr = '"&strzrr&"' "
End If
If strlsh <> "" Then
	If strsql<>"" Then 
		strSql = strSql&"and  xldh like '%"&strlsh&"%' "
	else
		strSql="xldh like '%"&strlsh&"%' "
	End If
End If	
If strkhm <> "" Then
	If strsql<>"" Then 
		strSql = strSql&"and  rwnr like '%"&strkhm&"%' "
	else
		strSql="rwnr like '%"&strkhm&"%' "
	End If
End If	
iqitime="wcsj"
Select case strrwlr
	case "1"
		iqitime="rksj"
		If strSql="" then
			strSql = "rwlx='����' "
		else
			strSql=strSql&"and rwlx='����' "
		End If
	case "2"
		If strSql="" then
			strSql = "rwlx='����' "
		else
			strSql=strSql&"and rwlx='����' "
		End If	
	case "3"
		If strSql="" then
			strSql = "rwlx='װ��' "
		else
			strSql=strSql&"and rwlx='װ��' "
		End If	
End Select	

If dtstart<>"" then
		If strSql="" then
			strSql="datediff('d',"&iqitime&",'"&dtstart&"')<=0 and datediff('d',"&iqitime&",'"&dtend&"')>=0"
		else
			strSql=strSql&"and datediff('d',"&iqitime&",'"&dtstart&"')<=0 and datediff('d',"&iqitime&",'"&dtend&"')>=0"
		End If
End If
      
	set rs=server.createobject("adodb.recordset")
	If strSql="" then
		sqltex="select * from spc_mission Order BY wcsj desc,xldh desc"
	else
		sqltex="select * from spc_mission where "&strSql&" Order BY wcsj desc,xldh desc"
	End If
rs.open sqltex,conn,1,1
dim PerPage
PerPage=30
'����û������ʱ
If rs.eof and rs.bof then
response.write "<p align='center'><font color='#ff0000'>��û������</font></p>"
else
'ȡ��ҳ��,���ж��û�������Ƿ��������͵����ݣ��粻�ǽ��Ե�һҳ��ʾ
text="0123456789"
Rs.PageSize=PerPage
for i=1 to len(request("page"))
checkpage=instr(1,text,mid(request("page"),i,1))
if checkpage=0 then
exit for
end if
next

If checkpage<>0 then
  If NOT IsEmpty(request("page")) Then
   CurrentPage=Cint(request("page"))
   If CurrentPage < 1 Then CurrentPage = 1 
   If CurrentPage > Rs.PageCount Then CurrentPage = Rs.PageCount 
  Else
   CurrentPage= 1
  End If
   If not Rs.eof Then Rs.AbsolutePage = CurrentPage end if
Else
 CurrentPage=1
End if
  %>
      <table width="98%" border="2" bordercolor="#33FFCC">
        <caption>
        <span class="STYLE2">����ģ���������б�</span> <br />
        </caption>
        <tr>
          <th width="3%" class="STYLE4">ID</th>
          <th class="STYLE4">����</th>
          <th class="STYLE4">������</th>
          <th width="40%" class="STYLE4">��������</th>
          <th class="STYLE4">������</th>
          <th class="STYLE4">���ʱ��</th>
          <th class="STYLE4">��ֵ</th>
          <th class="STYLE4">����</th>
        </tr>
        <%if not rs.eof then
i=1
do while not rs.eof
%>
        <tr>
          <td class="STYLE4"><div align="center"><%=i%></div></td>
          <td class="STYLE4"><div align="center"><%=Rs("rwlx")%></div></td>
          <td class="STYLE4"><div align="center"><%=Rs("xldh")%></div></td>
          <td class="STYLE4"><div align="center"><%=Rs("rwnr")%></div></td>
          <td class="STYLE4"><div align="center"><%=Rs("zrr")%></div></td>
          <td class="STYLE4"><div align="center"><%=FormatDateTime(Rs("wcsj"),1)%></div></td>
          <td class="STYLE4"><div align="center"><%=Rs("fz")%></div></td>
          <td class="STYLE4"><label>
            <div align="center"> <a href="spc_change.asp?s_id=<%=Rs("ID")%>">����</a> | <a href="spc_indb.asp?action=del&id=<%=Rs("ID")%>" onClick="return confirm('ɾ���󽫲��ָܻ�!\n\nȷ��ɾ��  <%=Rs("xldh")%> ��������?');">ɾ��</a>
              </label>
            </div></td>
        </tr>
        <%
i=i+1
if i > Perpage then exit do
rs.movenext
loop
end if
%>
      </table>
      <table width="100%" border="0" align="center" cellpadding="3" cellspacing="0">
        <tr>
          <td><div align="right">
              <%
Response.write "ȫ��"
Response.write "��" & Cstr(Rs.RecordCount) & "������&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
Response.write "��" & Cstr(CurrentPage) &  "/" & Cstr(rs.pagecount) & "&nbsp;"
If currentpage > 1 Then
response.write "<a href='spc_list.asp?&page="+cstr(1)+"'>&nbsp;��ҳ&nbsp;</a>"
Response.write "<a href='spc_list.asp?page="+Cstr(currentpage-1)+"'>&nbsp;��һҳ&nbsp;</a>"
Else
Response.write "&nbsp;��һҳ&nbsp;"
End if
If currentpage < Rs.PageCount Then
Response.write "<a href='spc_list.asp?page="+Cstr(currentPage+1)+"'>&nbsp;��һҳ&nbsp;</a>"
Response.write "<a href='spc_list.asp?page="+Cstr(Rs.PageCount)+"'>&nbsp;βҳ&nbsp;</a>"
Else
Response.write ""
Response.write "&nbsp;��һҳ&nbsp;"
End if
Response.write "ת����"
response.write"<select name='sel_page' onChange='javascript:location=this.options[this.selectedIndex].value;'>"
    for i = 1 to Rs.PageCount
       if i = currentpage then 
          response.write"<option value='spc_list.asp?page="&i&"&id="&id&"' selected>"&i&"</option>"
       else
          response.write"<option value='spc_list.asp?page="&i&"&id="&id&"'>"&i&"</option>"
       end if
    next
response.write"</select>ҳ"
rs.close
%>
            </div></td>
        </tr>
        <tr>
          <td bgcolor="#00CC00"></td>
        </tr>
      </table>
      <%End If%>
    </td>
  </tr>
</table>
</table>
<!-- #include file="Inc/Foot.asp" -->
</BODY>
</HTML>