<!--#include file="Inc/Sysconn.asp" -->
<!--#include file="inc/Skin_CSS.asp"-->
<%
If session("UserName")="" then
	response.redirect "UserServer.asp"
End If

Dim strFeedBack,iqiy,iqim,dtstart,dtend,isearch
iqiy=Trim(Request("SearchY"))
iqim=Trim(Request("SearchM"))
'isearchΪ�¶Ⱥϸ�������Ⱥϸ��ʵ����ֱ�ʶ��Ϊ1ʱ��Ϊȫ��ϸ���
itime="" : dtstart="" : dtend=""
'If iqiy="" Then iqiy=year(now)
If iqiy<>"" and iqim<>"" then dtstart=cdate(iqiy&"��"&iqim&"��1��") : isearch=0
If iqiy<>"" and iqim="" then dtstart=cdate(iqiy&"��1��1��") : isearch=1
If iqiy="" and iqim<>"" then dtstart=cdate(year(now)&"��"&iqim&"��1��") : iqiy=year(now) : isearch=0
If iqiy="" and iqim="" then dtstart=cdate(year(now)&"��"&month(now)&"��1��") : iqiy=year(now) : isearch=0	'Ĭ��Ϊ����
If dtstart<>"" then
	If isearch=0 Then
		dtstart=dateadd("m",-6,dtstart)	'�ºϸ���Ϊ����ʱ��ǰ��6���³����ģ�ߺϸ����
		dtend=dateadd("m",1,dtstart)
		dtend=dateadd("d",-1,dtend)
	else
		dtend=dateadd("y",1,dtstart)
		dtend=dateadd("d",-1,dtend)
	End If 
End If

strFeedBack="SearchY="&iqiy&"&SearchM="&iqim
%>
<style type="text/css">
<!--
.STYLE2 {
	font-size: x-large;
	font-weight: bold;
	color: #000066;
}
.STYLE4 {
	color: #000066;
}
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
      <!--#include file="Left_Faq.asp" --></td>
    <td width="741" align="center" valign="top" background="Images_v1/rith_bg.gif"><table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="21"><img src="Images_v1/yx2_1.gif" width="35" height="21"></td>
          <td background="Images_v1/yx2_2.gif">&nbsp;</td>
          <td width="21" align="right"><img src="Images_v1/yx2.png" width="247" height="21"></td>
        </tr>
      </table>
      <table width="96%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td><img src="imgV2/faq_list.gif" border="0"></td>
          <td width="79%" valign="middle"><table width="100%" height="1" border="0" cellpadding="0" cellspacing="0" bgcolor="#00AD97">
              <tr>
                <td></td>
              </tr>
            </table></td>
          <td width="5%"><img src="imgV2/more2.gif" width="37" height="24" border="0"></td>
        </tr>
      </table>
      <%Call SearchInfo()%>
      <table width="98%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="1" colspan="2">&nbsp;</td>
        </tr>
        <tr>
          <td height="1" colspan="2" bgcolor="#00ffff"></td>
        </tr>
        <tr>
          <td height="1" colspan="2">&nbsp;</td>
        </tr>
      </table>
      <%     
If dtstart<>"" then
	strSql="datediff('d',cksj,'"&dtstart&"')<=0 and datediff('d',cksj,'"&dtend&"')>=0"
End If

Dim Mjts,GnMjts,HgTs,GnHgTs,Hgl,GnHgl,GwHgl,sqltex2,Rs2
Mjts=0 : GnMjts=0 : HgTs=0 : GnHgTs=0 : Hgl=0 : GnHgl=0 : GwHgl=0
set Rs2=server.createobject("adodb.recordset")
If strSql="" then
	sqltex2="select * from mission Order BY lsh,cksj desc"
else
	sqltex2="select * from mission where "&strSql&" Order BY lsh,cksj desc"
End If
Rs2.open sqltex2,conn,1,1
Mjts=Rs2.RecordCount			'ģ������
Rs2.Close
set Rs2=server.createobject("adodb.recordset")
If strSql="" then
	sqltex2="select * from mission where ysjg='���պϸ�' Order BY lsh,cksj desc"
else
	sqltex2="select * from mission where "&strSql&" and ysjg='���պϸ�' Order BY lsh,cksj desc"
End If
Rs2.open sqltex2,conn,1,1
HgTs=Rs2.RecordCount			'ģ�����պϸ�����
Rs2.Close
set Rs2=server.createobject("adodb.recordset")
If strSql="" then
	sqltex2="select * from mission where cjqy='����' Order BY lsh,cksj desc"
else
	sqltex2="select * from mission where "&strSql&" and cjqy='����' Order BY lsh,cksj desc"
End If
Rs2.open sqltex2,conn,1,1
GnMjts=Rs2.RecordCount			'����ģ������
Rs2.Close
set Rs2=server.createobject("adodb.recordset")
If strSql="" then
	sqltex2="select * from mission where ysjg='���պϸ�' and cjqy='����' Order BY lsh,cksj desc"
else
	sqltex2="select * from mission where "&strSql&" and ysjg='���պϸ�' and cjqy='����' Order BY lsh,cksj desc"
End If
Rs2.open sqltex2,conn,1,1
GnHgTs=Rs2.RecordCount			'����ģ�����պϸ�����
Rs2.Close
'''����ϸ���
Hgl=Round(100*HgTs/Mjts,1)&"%"							'ģ�ߺϸ���
GnHgl=Round(100*GnHgTs/GnMjts,1)&"%"					'����ģ�ߺϸ���
GwHgl=Round(100*(HgTs-GnHgTs)/(Mjts-GnMjts),1)&"%"		'����ģ�ߺϸ���
set rs=server.createobject("adodb.recordset")
If strSql="" then
	sqltex="select * from mission where ysjg='���պϸ�' Order BY lsh,cksj desc"
else
	sqltex="select * from mission where "&strSql&" and ysjg='���պϸ�' Order BY lsh,cksj desc"
End If
rs.open sqltex,conn,1,1
dim PerPage
PerPage=30
'����û������ʱ
If rs.eof and rs.bof then
response.write ("<p align='center'><font color='#ff0000'>����û��ģ�߳��⣡</font></p>")
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
      <table width="98%" border="1" cellspacing="0" bordercolor="#33FFCC">
        <caption>
        <span class="STYLE2"><%=iqiy%>��<%=iqim%>�¼���ģ�߳����Ժϸ���</span><br />
        <span class="STYLE4">
        <div align="right">����ʱ��:<%=dtstart%>/<%=dtend%></div>
        </span>
        </caption>
        <tr bgcolor="#d6f5a4">
          <th colspan="6" scope="col">ȫ��ģ��</th>
        </tr>
        <tr>
          <td width="15%">�ܹ�����ģ������</td>
          <td width="20%" align="center"><%=Mjts%></td>
          <td width="15%">�ܹ��ϸ�ģ������</td>
          <td width="20%" align="center"><%=HgTs%></td>
          <td width="10%">�ܺϸ���</td>
          <td width="20%" align="center"><strong><%=Hgl%></strong></td>
        </tr>
      </table>
      <table width="98%" border="1" cellspacing="0" bordercolor="#33FFCC">
        <tr bgcolor="#d6f5a4">
          <th colspan="6" scope="col">����ģ��</th>
          <th colspan="6" scope="col">����ģ��</th>
        </tr>
        <tr>
          <td width="8%">����ģ��</td>
          <td width="9%" align="center"><%=GnMjts%></td>
          <td width="8%">�ϸ�ģ��</td>
          <td width="9%" align="center"><%=GnHgTs%></td>
          <td width="6%">�ϸ���</td>
          <td width="9%" align="center"><strong><%=GnHgl%></strong></td>
          <td width="8%">����ĥ��</td>
          <td width="9%" align="center"><%=Mjts-GnMjts%></td>
          <td width="8%">�ϸ�ĥ��</td>
          <td width="9%" align="center"><%=HgTs-GnHgTs%></td>
          <td width="6%">�ϸ���</td>
          <td align="center"><strong><%=GwHgl%></strong></td>
        </tr>
      </table>
      <table width="98%" border="1" cellspacing="0"  bordercolor="#33FFCC">
        <tr bgcolor="#d6f5a4">
          <th width="3%" class="STYLE4">ID</th>
          <th class="STYLE4">��ˮ��</th>
          <th class="STYLE4">�ͻ�����</th>
          <th class="STYLE4">����ʱ��</th>
          <th class="STYLE4">��������</th>
        </tr>
        <%if not rs.eof then
i=1
do while not rs.eof
%>
        <tr>
          <td class="STYLE4"><div align="center"><%=i%></div></td>
          <td class="STYLE4"><div align="center"><a href="Faq_dis.asp?s_lsh=<%=rs("lsh")%>"><%=Rs("lsh")%></a></div></td>
          <td class="STYLE4"><div align="center"><%=Rs("khmc")%>&nbsp;</div></td>
          <td class="STYLE4"><div align="center"><%=FormatDateTime(Rs("cksj"),1)%>&nbsp;</div></td>
          <td class="STYLE4"><div align="center"><%=Rs("ysjg")%>&nbsp;</div></td>
        </tr>
        <%
i=i+1
if i > Perpage then exit do
rs.movenext
loop
end if
%>
      </table>
      <%
      Call showpage("Faq.asp?"&strFeedBack,Rs.RecordCount,PerPage,True,True,"��")
      End If%>
    </td>
  </tr>
</table>
<!-- #include file="Inc/Foot.asp" -->
<map name="Map">
  <area shape="rect" coords="18,14,69,24" href="News.asp">
</map>
</BODY>
</HTML><%
Function SearchInfo()
%>
<table border="0" cellpadding="2" cellspacing="0" width="100%">
  <form action="<%=Request.Servervariables("SCRIPT_NAME")%>" method="get" name="frm_searchinfo" id="frm_searchinfo" onsubmit='return true;'>
    <tr>
      <td>&nbsp;&nbsp;ʱ��:</td>
      <td><select name="SearchY">
          <option value=""></option>
          <%for i=year(now)-2 to year(now)+2%>
          <option value="<%=i%>" <%If iqiy=cstr(i) then%>selected<%End If%>><%=i%></option>
          <%next%>
        </select>
        ��
        <select name="SearchM">
          <option value=""></option>
          <%for i=1 to 12%>
          <option value="<%=i%>" <%If iqim=cstr(i) then%>selected<%End If%>><%=i%></option>
          <%next%>
        </select>
        ��</td>
      <td align="left"><input class=btn_2k3 type="submit" value="�� ��" /></td>
    </tr>
  </form>
</table>
<%
End Function
%>
