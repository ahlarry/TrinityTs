<!--#include file="Inc/Sysconn.asp" -->
<!--#include file="inc/Skin_CSS.asp"-->
<!--#include file="Inc/user_dbinf.asp"-->
<SCRIPT language=Javascript src="Inc/jsfunc.js"></SCRIPT>
<%
If session("UserName")="" then
	response.redirect "UserServer.asp"
End If

function cutstr(tempstr,tempwid)
if len(tempstr)>tempwid then
cutstr=left(tempstr,tempwid)&"..."
else
cutstr=tempstr
end if
end function

'������������
Dim strlsh, strddh, strkhm, strcz, StrRwlx, strFeedBack
strlsh=Trim(Request("lsh"))
If instr(strlsh,"��")>0 Then 	'�������ȫ�Ƿ���ת��Ϊ���
	strlsh=Replace(strlsh,"��",",")
End If
strddh=Trim(Request("ddh"))
strkhm=Trim(Request("khm"))
strcz=Trim(Request("cz"))
StrRwlx=Trim(Request("rwlx"))
strFeedBack="rwlx="&StrRwlx&"&lsh="&strlsh&"&ddh="&strddh&"&khm="&strkhm&"&cz="&strcz
If StrRwlx="" Then StrRwlx="1" End If
%>
<style type="text/css">
<!--
.STYLE1 {
	color: #EEEEEE;
}
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
  <tbody>
    <tr>
      <td valign="bottom" background="imgV2/left_bg1.png" class="bg_x">&nbsp;</td>
      <td width="9" align="right" valign="top" background="imgV2/left_bg2.png"><img src="imgV2/left_yy.png" width="9" height="510" /></td>
      <td width="211" valign="top" background="imgV2/left_bg.gif"><!--#include file="login.asp" -->
        <!--#include file="Left_pld.asp" --></td>
      <td width="741" align="center" valign="top" background="Images_v1/rith_bg.gif" bgcolor="#F8FFE8"><table width="100%" border="0" cellpadding="0" cellspacing="0">
          <tbody>
            <tr>
              <td width="21"><img src="Images_v1/yx2_1.gif" width="35" height="21" /></td>
              <td background="Images_v1/yx2_2.gif">&nbsp;</td>
              <td width="21" align="right"><img src="Images_v1/yx2.png" width="247" height="21" /></td>
            </tr>
          </tbody>
        </table>
        <table width="98%" border="0" cellspacing="0" cellpadding="0">
          <tbody>
            <tr>
              <td width="40%" height="18">���ƻ�����&gt;&gt; ����������� </td>
              <td width="71%"><div align="right">
                  <label>
                  <input name="rwlx" type="radio" value="1" onClick="location.href='pld_assign.asp?rwlx='+ this.value" <%If strrwlx="1" Then%>="" checked="checked" <%End If%>="" />
                  ���ڵ���
                  <input name="rwlx" type="radio" value="0" onClick="location.href='pld_assign.asp?rwlx='+ this.value" <%If strrwlx="0" Then%>="" checked="checked" <%End If%>="" />
                  �������</label>
                </div></td>
            </tr>
            <tr>
              <td height="1" colspan="2" bgcolor="#D6DCE3"></td>
            </tr>
            <tr>
              <td height="1" colspan="2">&nbsp;</td>
            </tr>
          </tbody>
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
        <table width="98%" border="2" bordercolor="#33FFCC">
          <%If Strlsh="" Then
			Call Mlist()
		else
			Call MAssign()%>
          <table width="98%" border="0" cellspacing="0" cellpadding="0">
            <tbody>
              <tr>
                <td height="1" colspan="2" bgcolor="#D6DCE3"></td>
              </tr>
              <tr>
                <td height="1" colspan="2">&nbsp;</td>
              </tr>
            </tbody>
          </table>
          <%Call Mlist()
		End If %>
        </table></td>
    </tr>
  </tbody>
</table>
<!-- #include file="Inc/Foot.asp" -->
<%
Function SearchInfo()
%>
<table border="0" cellpadding="2" cellspacing="0" width="100%">
  <form action="<%=Request.Servervariables("SCRIPT_NAME")%>" method="get" name="frm_searchinfo" id="frm_searchinfo" onsubmit='return true;'>
    <tr>
      <td>&nbsp;&nbsp;����ɸѡ����:</td>
      <td>������:
        <input name="ddh" value="<%=strddh%>"></td>
      <td>��ˮ��:
        <input name="lsh" size="10" value="<%=strlsh%>"></td>
      <td>�ͻ���:
        <input name="khm" size="10" value="<%=strkhm%>"></td>
      <td>��������:
        <select name="cz">
          <option value="" selected="selected">ȫ��</option>
          <option value="0" <%If strcz="0" Then%>selected="selected"<%End If%>>����</option>
          <option value="1" <%If strcz="1" Then%>selected="selected"<%End If%>>����</option>
          <option value="2" <%If strcz="2" Then%>selected="selected"<%End If%>>���</option>
          <option value="3" <%If strcz="3" Then%>selected="selected"<%End If%>>����</option>
        </select></td>
      <td><input type="hidden" name="rwlx" value="<%=strrwlx%>">
        <input class=btn_2k3 type="submit" value="�� ��" /></td>
    </tr>
  </form>
</table>
<%
End Function

Function Mlist()
strSql=""
	If strLsh <> "" Then
		strSql = "and lsh like '%"&strLsh&"%' "
	End If
	If strddh <> "" Then
		strSql = "and  ddh like '%"&strddh&"%' " & strSql
	End If
	If strkhm <> "" Then
		strSql = "and  khmc like '%"&strkhm&"%' " & strSql
	End If
	Select case strcz
		case "0"
			strSql = "and isNull(sjjssj) and rwlx<>'�ǵ�' and rwlx<>'TB' " & strSql
		case "1"
			strSql = "and not(isNull(sjjssj)) and sjtsyl=0 " & strSql
		case "2"
			strSql = "and isNull(rksj) and (sjtsyl<>0 or rwlx='�ǵ�' or rwlx='TB') " & strSql
		case "3"
			strSql = "and not(isNull(rksj)) and isNull(cksj)" & strSql
	End Select

set rs=server.createobject("adodb.recordset")
If StrRwlx="0" Then
	sqltex="select * from mission where not(isNull(cksj)) and isNull(cwjssj) Order by lsh"
else
	sqltex="select * from mission where isNull(cksj) "&strsql&" Order by jhjssj,lsh"
End If
rs.open sqltex,conn,1,1
dim PerPage
If Strlsh="" Then
	PerPage=20
else
	PerPage=10
End If
'����û������ʱ
If rs.eof and rs.bof then
response.write "<p align='center'><font color='#ff0000'>û���ҵ��������</font></p>"
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
<table width="98%" border="1" cellpadding="0" cellspacing="0"  bordercolor="#33FFCC">
  <caption>
  <span class="STYLE2">δ������������б�</span><br />
  </caption>
  <tbody>
    <tr>
      <th width="3%" class="STYLE4">ID</th>
      <th class="STYLE4">������</th>
      <th class="STYLE4">��ˮ��</th>
      <th class="STYLE4">�ͻ�����</th>
      <%If StrRwlx="1" Then %>
      <th class="STYLE4">��������</th>
      <th class="STYLE4">�����</th>
      <th class="STYLE4">ʵ������</th>
      <th class="STYLE4">���Դ���</th>
      <th class="STYLE4">���Խ���</th>
      <th class="STYLE4">���ʱ��</th>
      <%else%>
      <th class="STYLE4">����ʱ��</th>
      <th class="STYLE4">��ʼ����</th>
      <th width="30%" class="STYLE4">������Ϣ</th>
      <%End If%>
      <th class="STYLE4">����</th>
    </tr>
    <%if not rs.eof then
i=1
do while not rs.eof
%>
  <form name="frm_searchlsh" action='<%=Request.ServerVariables("SCRIPT_NAME") & "?rwlx=" &Strrwlx& "&Lsh=" & Rs("lsh")%>' method="post" onSubmit="return searchlsh_true();">
    <tr>
      <td class="STYLE4"><div align="center"> <%=i%></div></td>
      <td class="STYLE4"><div align="center"> <%=Rs("ddh")%>&nbsp;</div></td>
      <td class="STYLE4"><div align="center"> <a href='intpld_dis.asp?s_lsh=<%=rs("lsh")%>'><%=Rs("lsh")%>&nbsp;</a></div></td>
      <td class="STYLE4"><div align="center"> <%=Rs("khmc")%>&nbsp;</div></td>
      <%If StrRwlx="1" Then %>
      <td class="STYLE4"><div align="center"> <%=Rs("rwlx")%>&nbsp;</div></td>
      <td class="STYLE4"><div align="center"> <%=Rs("edtsyl")%>&nbsp;</div></td>
      <td class="STYLE4"><div align="center"> <%=Rs("sjtsyl")%>&nbsp;</div></td>
      <td class="STYLE4"><div align="center"> <%=Rs("sjtscs")%>&nbsp;</div></td>
      <td class="STYLE4"><div align="center"> <%=Rs("sjjssj")%>&nbsp;</div></td>
      <td class="STYLE4"><div align="center"> <%=Rs("rksj")%>&nbsp;</div></td>
      <%else%>
      <td class="STYLE4"><div align="center"> <%=Rs("cksj")%>&nbsp;</div></td>
      <td class="STYLE4"><div align="center"> <%=Rs("cwkssj")%>&nbsp;</div></td>
      <td class="STYLE4"><div align="center" alt="<%=Rs("bz")%>"> <%=cutstr(Rs("bz"),10)%>&nbsp;</div></td>
      <%End If%>
      <td class="STYLE4"><label>
        <div align="center">
          <%if StrRwlx="1" Then%>
          <input type="submit" name="Submit2" <%If isNull(Rs("sjjssj")) Then%>value="����"<%elseif isNull(Rs("rksj")) Then%>value="���"<%else%>value="����"<%End If%> />
          <%else%>
          <input type="submit" name="Submit2" value="����" />
          <%End If%>
          <input type="hidden" name="AssLsh" value='<%=Rs("lsh")%>' />
        </div>
        </label>
        <div align="center"> </div></td>
    </tr>
  </form>
  <%
i=i+1
if i > Perpage then exit do
rs.movenext
loop
end if
%>
  </tbody>

</table>
<%
      Call showpage("pld_assign.asp?"&strFeedBack,Rs.RecordCount,PerPage,True,True,"��")
      %>
<%End If
  End Function

  Function MAssign()
  %>
<%
  If StrRwlx="1" Then
		Call Assed_ins()
		strSql="select * from [mission] where lsh='"&strlsh&"' and isNull(cksj)"
		set Rs=server.createobject("adodb.recordset")
		Rs.open strSql,Conn,1,1
		If Not(Rs.eof Or Rs.bof) Then
			If IsNull(Rs("sjjssj")) and Rs("rwlx")<>"�ǵ�" and Rs("rwlx")<>"TB" Then
				Call InsAssign()
			else
				if Rs("sjtsyl")=0  and Rs("rwlx")<>"�ǵ�" and Rs("rwlx")<>"TB" Then
					Call YlAssign()
				else
					if isnull(Rs("rksj")) Then
						Call RkAssign()
					else
						If IsNull(Rs("cksj")) Then
						 	Call CkAssign()
						End If
					End If
				End If
			End if
		else
			Response.write("<font color='red'>���񲻴��ڻ�����ɣ�</font>")
		End If
		Rs.Close
	else
		Dim TmpArr,strjdu
		TmpArr=split(strlsh,",")	'������к�ͬʱ����
		strjdu="" : errmsg=""
		For i=0 to ubound(TmpArr)
			Call Assed_ext(TmpArr(i))
			strSql="select * from [mission] where lsh='"&TmpArr(i)&"' and isNull(cwjssj)"
			set Rs=server.createobject("adodb.recordset")
			Rs.open strSql,Conn,1,1
			If Rs.eof Or Rs.bof Then
				errmsg=errmsg&"��ˮ�� <strong>"&TmpArr(i)&"</strong> ���񲻴��ڻ������!<br>"
			else
				If strjdu="" Then
					strjdu=Rs("ysjg")
				else
					strjdu=strjdu&","&Rs("ysjg")
				End If
			End If
		Next
		Dim Tmpjdu,j
'		errmsg=errmsg&strjdu
		strjdu=split(strjdu,",")
		For i=0 to ubound(strjdu)
			For j=1 to ubound(strjdu)
				if strjdu(i)<>strjdu(j) then
					errmsg=errmsg&"������Ȳ�һ�£��޷��ϳ�һ��¼��!<br>"
					exit for
				End If
			Next
			if instr(errmsg,"���Ȳ�һ��")>0 then exit for
		Next
		If errmsg<>"" Then
			Call WriteErrMsg()
		else
			If strjdu(0)="���ڵ���" Then
				Call ExtAssign(strlsh)
			else
				Call ExtFzindb(strlsh)
			End If
		End If
	End If
End Function

Function Assed_ins()	'�ѷ��䳧������
	Dim m
	strSql="select * from [ins_tsxx] where lsh='"&strlsh&"'"
	set Rs=server.createobject("adodb.recordset")
	Rs.open strSql,Conn,1,1
	If Not(Rs.eof Or Rs.bof) Then
%>
<table width="98%" border="2" bordercolor="#33FFCC">
  <tr>
    <th scope="col">ID</th>
    <th scope="col">����Ա</th>
    <th scope="col">��������</th>
    <th scope="col">������ʽ</th>
  </tr>
  <%For m=1 to Rs.RecordCount%>
  <tr>
    <td><div align="center"><%=m%></div></td>
    <td><div align="center"><%=Rs("tsy")%>&nbsp;</div></td>
    <td><div align="center"><%=Rs("tsnr")%>&nbsp;</div></td>
    <td><div align="center"><%=Rs("cc")%>&nbsp;</div></td>
  </tr>
  <%Rs.movenext
  next%>
</table>
<br>
<%End If
    Rs.Close
  End Function

Function Assed_ext(ilsh)	'�ѷ��䳧������
	strSql="select * from [ext_tsxx] where lsh='"&ilsh&"'"
	set Rs=server.createobject("adodb.recordset")
	Rs.open strSql,Conn,1,1
	If Not(Rs.eof Or Rs.bof) Then
%>
<table width="98%" border="2" bordercolor="#33FFCC">
  <tr>
    <th scope="col">����</th>
    <th scope="col">��ʼʱ��</th>
    <th scope="col">����ʱ��</th>
    <th scope="col">����</th>
    <th scope="col">����</th>
    <th scope="col">��ʼʱ��</th>
    <th scope="col">����ʱ��</th>
    <th scope="col">����</th>
  </tr>
  <%do while Not(Rs.eof Or Rs.bof)%>
  <tr>
    <td><div align="center"><%=Rs("tsy")%></div></td>
    <td><div align="center"><%=Rs("kssj")%></div></td>
    <td><div align="center"><%=Rs("jssj")%></div></td>
    <td><div align="center"><%=Rs("cc")%></div></td>
    <%Rs.movenext
    If Not(Rs.eof Or Rs.bof) Then%>
    <td><div align="center"><%=Rs("tsy")%></div></td>
    <td><div align="center"><%=Rs("kssj")%></div></td>
    <td><div align="center"><%=Rs("jssj")%></div></td>
    <td><div align="center"><%=Rs("cc")%></div></td>
    <%Rs.movenext
    else%>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <%End If%>
  </tr>
  <%loop%>
</table>
<br>
<%End If
  Rs.Close
  End Function

  Function InsAssign()	'�����������%>
<form name="Form_Assing" id="Form_Assing" method="post" action="pld_assindb.asp?action=Massign">
  <table width="70%" border="2" bordercolor="#33FFCC">
    <tr>
      <td><div align="center"><strong>��ѡ����Ա</strong></div></td>
      <td rowspan="2"><div align="center">
          <INPUT language="javascript" name="btn_select_addall" onClick="fun_select_addall(document.Form_Assing)" style="COLOR: #000066; FONT-FAMILY: Webdings; FONT-SIZE: 12pt; FONT-WEIGHT: normal; HEIGHT: 28px; WIDTH: 27px" title="�������" type=button value=: DESIGNTIMESP="17713">
          <br>
          <br>
          <INPUT language="Javascript" name="btn_select_addany" onClick="fun_select_addany(document.Form_Assing)" style="COLOR: #000066; FONT-FAMILY: Webdings; FONT-SIZE: 12pt; FONT-WEIGHT: normal; HEIGHT: 28px; WIDTH: 27px" title="�����ѡ" type=button value="8">
          <br>
          <br>
          <INPUT language="javascript" name="btn_select_dltany" onClick="fun_select_dltany(document.Form_Assing)" style="COLOR: #000066; FONT-FAMILY: Webdings; FONT-SIZE: 12pt; FONT-WEIGHT: normal; HEIGHT: 28px; WIDTH: 27px" title ="ɾ����ѡ" type=button value="7">
          <br>
          <br>
          <INPUT language="javascript" name="btn_select_dltall" onClick="fun_select_dltall(document.Form_Assing)" style="COLOR: #000066; FONT-FAMILY: Webdings; FONT-SIZE: 12pt; FONT-WEIGHT: normal; HEIGHT: 28px; WIDTH: 27px" title ="ɾ������" type=button value="9">
        </div>
        <div align="center"></div></td>
      <td><div align="center"><strong>��ѡ����Ա</strong></div></td>
    </tr>
    <tr>
      <td><div align="center">
          <select name="dxtsy" style="height:160px;width:100px" multiple="multiple" id="dxtsy">
            <%for i = 0 to ubound(c_allzy)%>
            <option value="<%=c_allzy(i)%>"><%=c_allzy(i)%></option>
            <%next%>
          </select>
        </div></td>
      <td><div align="center">
          <select name="yxtsy" style="height:160px;width:100px" multiple="multiple" id="yxtsy" >
            <%
          	strSql="select * from [ins_tsxx] where lsh='"&strlsh&"' order by Id desc"
			set Rs=server.createobject("adodb.recordset")
			Rs.open strSql,Conn,1,1
			If Not(Rs.eof Or Rs.bof) Then
				StrTsy=split(Rs("tsy"),", ")
		        for i = 0 to ubound(StrTsy)%>
            <option value="<%=StrTsy(i)%>"><%=StrTsy(i)%></option>
            <%next
			End If
			Rs.Close
          %>
          </select>
        </div></td>
    </tr>
    <tr>
      <td colspan="3"><div align="center">��������:
          <select name="tsnr">
            <option value="ȫ��">ȫ��</option>
            <option value="ģͷ">ģͷ</option>
            <option value="����">����</option>
          </select>
          &nbsp;&nbsp;�������:
          <input type="text" name="cycs" size="5" value="1" onKeyPress="javascript:validationNumber(this, 'u_float', 100, '');">
          &nbsp;&nbsp;���շ�ʽ:
          <select name="ccfs">
            <option value="���ڵ���" selected>���ڵ���</option>
            <option value="�ֵ�">�ֵ�</option>
            <option value="��Ʒ�ϸ�">��Ʒ�ϸ�</option>
            <option value="һ��ϸ�">һ��ϸ�</option>
            <option value="����ϸ�">����ϸ�</option>
            <option value="����ϸ�">����ϸ�</option>
            <option value="֪ͨ����">֪ͨ����</option>
          </select>
        </div></td>
    </tr>
  </table>
  <br>
  <label>
  <div align="center">
    <input type="Submit" name="Submit3" onClick="SelTsy(document.Form_Assing);" value="��ȷ ����" />
    <input type="hidden" name="TsLsh" value='<%=strlsh%>' />
    <input type="hidden" name="Tsrwlx" value='<%=strrwlx%>' />
  </div>
  </label>
</form>
<%End Function

Function YlAssign()		'����
%>
<form name="Form_YL" id="Form_YL" method="post" action="pld_assindb.asp?action=YL">
  <table width="100%" border="2" bordercolor="#33FFCC">
    <tr>
      <td><div align="center">��ģ��������
          <input type="text" name="smyl" />
          &nbsp;&nbsp;
          <script language=javascript>
  		var myDate=new dateSelector();
  		myDate.year;
 		myDate.inputName='yxtsyjs';
  		myDate.display();
		</script>
          &nbsp;&nbsp;
          <input type="Submit" name="Submit4" value="��ȷ ����" />
          <input type="hidden" name="TsLsh" value='<%=strlsh%>' />
          <input type="hidden" name="Tsrwlx" value='<%=strrwlx%>' />
        </div></td>
    </tr>
  </table>
</form>
<%
End Function

Function RkAssign()		'���
%>
<form name="Form_Ruku" id="Form_Ruku" method="post" action="pld_assindb.asp?action=Ruku">
  <table width="100%" border="2" bordercolor="#33FFCC">
    <tr>
      <td><div align="center">���ʱ�䣺
          <script language=javascript>
  		var myDate=new dateSelector();
  		myDate.year;
 		myDate.inputName='rksj';
  		myDate.display();
		</script>
          &nbsp;&nbsp;
          <input type="Submit" name="Submit4" value="��ȷ ����" />
          <input type="hidden" name="TsLsh" value='<%=strlsh%>' />
          <input type="hidden" name="Tsrwlx" value='<%=strrwlx%>' />
        </div></td>
    </tr>
  </table>
</form>
<%
End Function

Function CkAssign()		'����
%>
<form name="Form_Chuuku" id="Form_Chuuku" method="post" action="pld_assindb.asp?action=Chuuku">
  <table width="100%" border="2" bordercolor="#33FFCC">
    <tr>
      <td><div align="center">����ʱ�䣺
          <script language=javascript>
  		var myDate=new dateSelector();
  		myDate.year;
 		myDate.inputName='cksj';
  		myDate.display();
		</script>
          &nbsp;&nbsp;
          <input type="Submit" name="Submit4" value="��ȷ ����" />
          <input type="hidden" name="TsLsh" value='<%=strlsh%>' />
          <input type="hidden" name="Tsrwlx" value='<%=strrwlx%>' />
        </div></td>
    </tr>
  </table>
</form>
<%
End Function

Function ExtAssign(lsh)	'�����������
Dim TmpArr,LshArr,Tmpstr,j,k
TmpArr=split(lsh,",")
For i=0 to ubound(TmpArr)
    strSql="select * from [mission] where lsh='"&TmpArr(i)&"'"
	set Rs=server.createobject("adodb.recordset")
	Rs.open strSql,Conn,1,1
	if not Rs.eof Then
		Select Case Rs("dmdj")
			Case "A��":Tmpstr="1"
			Case "B��":Tmpstr="2"
			Case "C��":Tmpstr="3"
			Case "D��":Tmpstr="4"
			Case "E��":Tmpstr="5"
			Case "F��":Tmpstr="6"
			Case "G��":Tmpstr="7"
			Case "H��":Tmpstr="8"
			Case Else:Tmpstr="9"
		End Select
		TmpArr(i)=TmpArr(i)&"||"&Tmpstr&Left(Rs("mjdj"),1)
	End If
	Rs.Close
Next
%>
<form name="Form_ExtA" id="Form_ExtA" method="post" action="pld_assindb.asp?action=Massign">
  <table width='80%' border="0" cellspacing="0" bordercolor="#33FFCC">
    <tr>
      <td colspan="6" align="center"><span class="STYLE2">����<%=lsh%>���������������</span></td>
    </tr>
    <%
    Tmpstr=ubound(TmpArr)/3
    if int(Tmpstr)<>Tmpstr Then
  		Tmpstr=int(Tmpstr+1)
	end if
	i=0
	For j=0 to Tmpstr
    %>
    <tr>
        <%k=0
       For k=i to i+2
        	If k <= ubound(TmpArr) Then
      			LshArr=split(TmpArr(k),"||")
     			Response.Write("<td align='center'><font color='#FF0000'>"&LshArr(0)&":"&LshArr(1)&"&nbsp;ģ��ϵ��:<input name='mjxs"&k&"' type='text' size='10'></font></td>")
     		End If
       next
        %>
    </tr>
    <%i=i+3
    next%>
    <tr>
      <td align="left"><font color='#FF0000'>[Ctrl+Enter ����һ��]</font></td>
      <td>&nbsp;</td>
      <td align="right"><font color='#FF0000'>[Shift+Enter ɾ��һ��]</font> </td>
    </tr>
  </table>
  <table id="ExtA_TB" width="80%" border="0" cellspacing="0" bordercolor="#33FFCC">
    <tr id='ExtA_Tr'>
      <td>������:
        <select name="zrr" id="zrr">
          <%for i = 0 to ubound(c_allzy)%>
          <option value="<%=c_allzy(i)%>"><%=c_allzy(i)%></option>
          <%next%>
        </select>
      </td>
      <td align="center">��ʼ����:
        <input name="tskssj" type="text" onFocus="CalendarWebControl.show(this,false,this.value);" size="10"></td>
      <td align="center"> ��������:
        <input name="tsjssj" type="text" onFocus="CalendarWebControl.show(this,false,this.value);" size="10"></td>
      <td><input value="���" type="button" onClick="Addrow()">
        <input value="ɾ��" type="button" onClick="Delrow()">
      </td>
    </tr>
  </table>
  <table width="70%" border="1" cellspacing="0" bordercolor="#33FFCC">
    <tr>
      <td><span class="STYLE4">��ע:</span>
        <input type="hidden" id="TsNum" name="TsNum"/></td>
      <td><textarea name="beiz" cols="65" rows="5"></textarea></td>
    </tr>
    <tr>
      <td colspan="2" align="center">���գ�
        <select name="ccfs">
          <option value="���ڵ���" selected>���ڵ���</option>
          <option value="���պϸ�">���պϸ�</option>
          <option value="֪ͨ����">֪ͨ����</option>
          <option value="�ܾ�����">�ܾ�����</option>
        </select>
        &nbsp;&nbsp;
        <input type="Submit" name="Submit5" value="��ȷ ����" />
        <input type="hidden" name="TsLsh" value='<%=lsh%>' /></td>
    </tr>
  </table>
</form>
<%
End Function

Function ExtFzindb(slsh)	'���������ֵ����(���δ�ϵͳ)
%>
<form name="Form_ExtF" method="post" action="pld_assindb.asp?action=ExtF" onSubmit='return CheckXs(<%=i%>);'>
  <table id="Ext_FzTb" width='90%' border="0" cellspacing="0" bordercolor="#33FFCC">
    <caption>
    <span class="STYLE2">����<%=slsh%> ����ϵ��</span><br />
    </caption>
    <%	'-----------------------------------������Է�ֵ
    Dim strtsfz,ilsh,iid,iyfz,ixs,ifz,ispan
    strtsfz=0 : ilsh=slsh : strlsh=split(slsh,",")
    For i=0 to ubound(strlsh)
		strSql="select * from [mission] where lsh='"&strlsh(i)&"'"
		set Rs=server.createobject("adodb.recordset")
		Rs.open strSql,Conn,1,1
		If Not(Rs.eof Or Rs.bof) Then
			strtsfz=strtsfz + Rs("cwtsfz")*Rs("djxs")		'������Է�ֵ
		End If
		Rs.Close
	next
    strSql="select * from [ext_tsfz] where lsh='"&ilsh&"'"
	set Rs=server.createobject("adodb.recordset")
	Rs.open strSql,Conn,1,1
	i=0
	Do while not Rs.eof
	iid="id"&i : iyfz="yfz"&i : ixs="xs"&i : ifz="fz"&i : ispan="span_fz"&i
	%>
    <tr>
      <td>����Ա:<%=Rs("zrr")%>
        <input type="hidden" name=<%=iid%> id=<%=iid%> value=<%=Rs("id")%> /></td>
      <td>��ֵ:
        <input type="text" name=<%=iyfz%> id=<%=iyfz%> size="8" value=<%=Rs("fz")%> readonly></td>
      <td>ϵ��:
        <input type="text" name=<%=ixs%> id=<%=ixs%> size="8" onKeyDown="valNum(event);" onpaste="valClip(event);" onChange="ChangeXs(this.value,<%=strtsfz%>,<%=i%>);">
        &nbsp;X&nbsp;<%=strtsfz%></td>
      <td><span id=<%=ispan%>>���յ÷�:<%=Rs("fz")%></span>
        <input type="hidden" id=<%=ifz%> name=<%=ifz%> value=<%=Rs("fz")%> /></td>
    </tr>
    <%
    Rs.movenext
	i=i+1
	Loop
	Rs.Close
%>
    <tr>
      <td height="1" colspan="4">&nbsp;</td>
    </tr>
    <tr>
      <td colspan="4" align="center"><input type="hidden" id="ExtFNB" name="ExtFNB" value=<%=i-1%> />
        <input type="hidden" name="TsLsh" value='<%=ilsh%>' />
        <input type="Submit" name="Submit6" value="��ȷ ����" /></td>
    </tr>
  </table>
</form>
<%End Function%>
<script language="javascript">
//���ڵ���Աѡ��
function fun_select_addany(theform){
    var i;
    for (i=0;i<theform.dxtsy.length;i++){
        if (theform.dxtsy.options[i].selected == true){
           theform.yxtsy.options[theform.yxtsy.length]=new Option(theform.dxtsy.options[i].text);
        }
    }
    sort_select(theform);
}

function fun_select_addall(theform){
    var i;
    for (i=0;i<theform.dxtsy.length;i++){
        theform.yxtsy.options[theform.yxtsy.length]=new Option(theform.dxtsy.options[i].text);
    }
    sort_select(theform);
}

function fun_select_dltany(theform){
   var i;
   for (i=0;i<theform.yxtsy.length;i++){
       if (theform.yxtsy.options[i].selected == true){
          theform.yxtsy.options[i] =new Option("");
          theform.yxtsy.options.remove(i);
          i--;
       }
   }
   sort_select(theform);
}

function fun_select_dltall(theform){
    var i;
    for (i=0;i<theform.yxtsy.length;i++){
    }
    theform.yxtsy.length=0;
    sort_select(theform);
}

function sort_select(theform){
    var i;
    var right_array = new Array();
    var o = new Object();
    for (i=0;i<theform.yxtsy.length;i++){
        right_array[i] = new Array(theform.yxtsy.options[i].text);
    }
	for (var i=0,j=0; i<right_array.length; i++)
	{
		if (typeof o[right_array[i]] == 'undefined')
		{
			o[right_array[i]] = j++;
		}
	}
	right_array.length = 0;
	for (var key in o)
	{
	right_array[o[key]] = key;
	}
    theform.yxtsy.length=0;
    for (i=0;i<right_array.length;i++){
       		theform.yxtsy.options[theform.yxtsy.length]=new Option(right_array[i]);
    }
    right_array = new Array();
}

function SelTsy(theform)
{
    for (i=0;i<theform.yxtsy.length;i++){
        theform.yxtsy.options[i].selected = true;
    }
}
//���Ӻ�ɾ����
i=1;
function Addrow(){
   var newTR = ExtA_Tr.cloneNode(true);
   newTR.id="a"+(++i)
   ExtA_Tr.parentNode.insertAdjacentElement("beforeEnd",newTR);
   RowReset();
   document.Form_ExtA.TsNum.value=ExtA_TB.rows.length;
}

function RowReset(){
   var RowCount=ExtA_TB.rows.length-1
   var ReName=RowCount
   for (var i=0;i<ExtA_TB.rows[RowCount].cells.length;i++){

      var str=ExtA_TB.rows[RowCount].cells[i].innerHTML
      str=(str.replace(/name=[\d]*/i,"name="+ReName));
      ExtA_TB.rows[RowCount].cells[i].innerHTML=str
   }
}

function Delrow(){
	if (ExtA_TB.rows.length >1){
		ExtA_TB.deleteRow(ExtA_TB.rows.length-1);
	}
	document.Form_ExtA.TsNum.value=ExtA_TB.rows.length;
}

function View(tb){
  for (var r=0;r<tb.rows.length;r++){
     for (var c=0;c<tb.rows[r].cells.length;c++){
        alert(tb.rows[r].cells[c].innerHTML)
     }
  }
}

function ShortcutKey(){
	if((event.ctrlKey)&&(window.event.keyCode==13)) //"Ctrl" + "�س�"���һ��
   		Addrow()
	if((event.shiftKey)&&(window.event.keyCode==13)) //"Shift" + "�س�"ɾ��һ��
		Delrow()
}

document.onkeydown=ShortcutKey;
//����������ε���ϵ��
function ChangeXs(value,tsfz,n){
	var yfz=eval("document.all.yfz"+ n +".value");
	if(!value)(value=0)
	var zzdf=(parseFloat(yfz) + parseFloat(value)*tsfz).toFixed(2)
	eval("document.all.span_fz"+ n +".innerHTML ='���յ÷�:"+ zzdf +"'");
	eval("document.all.fz"+ n +".value = '"+ zzdf +"'");
}
//����ϵ���ܺͱ���Ϊ0����ȷ���ֲܷ���ı�
function CheckXs(arg){
	var strsum=0
	for (i=0;i<=arg;i++){
		var xs=eval("document.all.xs"+ i +".value");
		if(!xs)(xs=0)
		strsum = strsum + parseFloat(xs);
	}
	if (strsum!=0){
		alert("ϵ���ܺͱ���Ϊ0!")
		return false;
	}
	else
	{
		return true;
	}
}
</script>
