<!--#include file="Inc/mission.asp" -->
<!--#include file="Inc/Sysconn.asp" -->
<!--#include file="inc/Skin_CSS.asp"-->
<SCRIPT language=Javascript src="Inc/jsfunc.js"></SCRIPT>
<%
If session("UserName")="" then
	response.redirect "UserServer.asp"
End If

Dim StrLsh,StrRwlx
StrRwlx=Trim(Request("rwlx"))
StrLsh=Trim(Request("s_lsh"))
If StrRwlx="" Then StrRwlx=1

    Dim c_cjqy, c_dmdj
    c_cjqy="" : c_dmdj=""
	'ȡ������ȼ�
	strSql="select distinct cjqy from sys_jldj"
	set rs=server.createobject("adodb.recordset")
	rs.open strSql,conn,1,1
	do while not rs.eof
		if c_cjqy <> "" then
			c_cjqy = c_cjqy & "|||" & rs("cjqy")
		else
			c_cjqy = rs("cjqy")
		end if
		rs.movenext
	loop
	rs.close
	strSql="select distinct dmdj from sys_jldj"
	set rs=server.createobject("adodb.recordset")
	rs.open strSql,conn,1,1
	do while not rs.eof
		if c_dmdj <> "" then
			c_dmdj = c_dmdj & "|||" & rs("dmdj")
		else
			c_dmdj = rs("dmdj")
		end if
		rs.movenext
	loop
	rs.close
	c_cjqy = split(c_cjqy, "|||")
	c_dmdj = split(c_dmdj, "|||")

%>
<style type="text/css">
<!--
.STYLE1 {
	color: #EEEEEE
}
.STYLE2 {
	font-size: x-large;
	font-weight: bold;
	color: #000066;
}
.STYLE4 {
	color: #000066;
}
-->
</style>
<!-- #include file="Head.asp" -->
<table width="998" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td valign="bottom" background="imgV2/left_bg1.png" class="bg_x">&nbsp;</td>
    <td width="9" align="right" valign="top" background="imgV2/left_bg2.png"><img src="imgV2/left_yy.png" width="9" height="510"></td>
    <td width="211" valign="top" background="imgV2/left_bg.gif"><!--#include file="login.asp" -->
      <!--#include file="Left_pld.asp" --></td>
    <td width="741" align="center" valign="top" background="Images_v1/rith_bg.gif" bgcolor="#F8FFE8"><table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="21"><img src="Images_v1/yx2_1.gif" width="35" height="21"></td>
          <td background="Images_v1/yx2_2.gif">&nbsp;</td>
          <td width="21" align="right"><img src="Images_v1/yx2.png" width="247" height="21"></td>
        </tr>
      </table>
      <table width="98%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="29%" height="18">���ƻ�����&gt;&gt;����������</td>
          <td width="71%"><div align="right">&nbsp; </div></td>
        </tr>
        <tr>
          <td height="1" colspan="2" bgcolor="#D6DCE3"></td>
        </tr>
        <tr>
          <td height="1" colspan="2">&nbsp;</td>
        </tr>
      </table>
      <Table border=0 cellpadding=2 cellspacing=0 width="98%">
        <Form name=frm_searchlsh action="<%=Request.Servervariables("SCRIPT_NAME")%>" method="get" onsubmit='return true();'>
          <tr>
            <td>&nbsp;&nbsp;������ˮ��:
              <input type=text name="s_lsh" size=15 value="<%=Trim(Request("s_lsh"))%>">
              <input type="submit" value=" �� �� "></td>
          </tr>
        </Form>
      </Table>
      <Table border=0 cellpadding=2 cellspacing=0 width="98%">
        <tr>
          <td height="1" colspan="2" bgcolor="#D6DCE3"></td>
        </tr>
        <tr>
          <td height="1" colspan="2">&nbsp;</td>
        </tr>
      </Table>
      <%Call mtask_change(StrLsh)%></td>
  </tr>
</table>
<!-- #include file="Inc/Foot.asp" -->
<map name="Map">
  <area shape="rect" coords="18,14,69,24" href="News.asp">
</map>
</BODY>
</HTML><%Function mtask_change(x)
	If x="" Then Response.Write("������Ҫ���ĵ����������ˮ��!<br>") : Exit Function
	strSql="select * from [mission] where [lsh] like '%"&x&"%'"
	set Rs=server.createobject("adodb.recordset")
	Rs.open strSql,Conn,1,3
	If Rs.Eof Or Rs.Bof Then
		errmsg="��ˮ��Ϊ �� " & x & " �� �������鲻����!"
		Call WriteErrMsg()
	Else
		If Rs.RecordCount>1 Then
			response.redirect "intpld_list.asp?lsh="&x
		else
			If Not(IsNull(Rs("rksj"))) Then
					errmsg="��ˮ��Ϊ �� " & x & " �� ���������Ѿ����!"
					Call WriteErrMsg()
			Else
				Call xchange(Rs)
			End if
		End If
	End If
End Function

Function xchange(Rs)
%>
<form id='pldadd' name='pldadd'  action=pld_indb.asp?action=change method=post onSubmit='return CheckM();'>
  <table width="98%" border="2" bordercolor="#33FFCC">
    <caption>
    <span class="STYLE2">������ˮ��Ϊ<%=Rs("lsh")%>��������</span><br />
    </caption>
          <tr>
            <th colspan="6" scope="col"><div align="left" class="STYLE4">ģ����Ϣ</div></th>
          </tr>
          <tr>
            <td width="100"><span class="STYLE4">��ˮ��</span></td>
            <td><input name="lsh" type="text" id="lsh" size="15" value=<%=Rs("lsh")%> /></td>
            <td width="100"><span class="STYLE4">������</span></td>
            <td><span class="STYLE4">
              <label>
                <input name="ddh" type="text" id="ddh" size="15" value=<%=Rs("ddh")%> />
              </label>
              </span></td>
            <td width="100"><span class="STYLE4">�ͻ�����</span></td>
            <td><input name="khmc" type="text" id="khmc" size="15" value=<%=Rs("khmc")%> /></td>
          </tr>
          <tr>
            <td><span class="STYLE4">������</span></td>
            <td><select name="scxh" id="scxh">
                <option value="0" <%if Rs("scx")="0" Then%>selected<%End If%>>���65</option>
                <option value="1" <%if Rs("scx")="1" Then%>selected<%End If%>>Τ��114</option>
                <option value="2" <%if Rs("scx")="2" Then%>selected<%End If%>>AROGOS93</option>
                <option value="3" <%if Rs("scx")="3" Then%>selected<%End If%>>��������58 </option>
                <option value="4" <%if Rs("scx")="4" Then%>selected<%End If%>>��������112</option>
                <option value="5" <%if Rs("scx")="5" Then%>selected<%End If%>>����˹���90</option>
              </select></td>
            <td><span class="STYLE4">ǣ���ٶ�</span></td>
            <td><p class="STYLE4">
                <input name="qysd" type="text" id="qysd" size="10" value=<%=Rs("qysd")%> onKeyPress="javascript:validationNumber(this, 'u_float', 100, '');" onChange="CheckYl();">
                m/min <br />
              </p></td>
            <td><span class="STYLE4">����</span></td>
            <td><input name="mz" type="text" id="mz" size="10" value=<%=Rs("mz")%>  onKeyPress="javascript:validationNumber(this, 'u_float', 100, '');" onChange="CheckYl();">
              Kg/m</td>
          </tr>
          <tr>
            <td><span class="STYLE4">��������</span></td>
            <td><select name="rwnr" id="rwnr" onChange="CheckYl();">
                <option value="ȫ��" <%if Rs("rwlx")="ȫ��" Then%>selected<%End If%>>ȫ��</option>
                <option value="��ͷ" <%if Rs("rwlx")="��ͷ" Then%>selected<%End If%>>��ͷ</option>
                <option value="����" <%if Rs("rwlx")="����" Then%>selected<%End If%>>����</option>
                <option value="��Ʒ" <%if Rs("rwlx")="��Ʒ" Then%>selected<%End If%>>��Ʒ</option>
                <option value="�ֵ�" <%if Rs("rwlx")="�ֵ�" Then%>selected<%End If%>>�ֵ�</option>
                <option value="�ǵ�" <%if Rs("rwlx")="�ǵ�" Then%>selected<%End If%>>�ǵ�</option>
                <option value="����" <%if Rs("rwlx")="����" Then%>selected<%End If%>>����</option>
                <option value="TB" <%if Rs("rwlx")="TB" Then%>selected<%End If%>>TB</option>
                <option value="TA" <%if Rs("rwlx")="TA" Then%>selected<%End If%>>TA</option>
              </select></td>
            <td><span class="STYLE4">�ͻ����</span></td>
            <td><span class="STYLE4">
                <select name="cjqy" id="cjqy" onChange="CheckYl();">
       	         <%for i = 0 to ubound(c_cjqy)%>
         	       <option value='<%=c_cjqy(i)%>' <%if Rs("cjqy")=c_cjqy(i) Then%>selected<%End If%>><%=c_cjqy(i)%></option>
                <%next%>
                </select>
              </span></td>
            <td><span class="STYLE4">����ȼ�</span></td>
            <td><span class="STYLE4">
              <select name="dmdj" id="dmdj" onChange="CheckYl();">
                <%for i = 0 to ubound(c_dmdj)%>
                	<option value='<%=c_dmdj(i)%>' <%if Rs("dmdj")=c_dmdj(i) Then%>selected<%End If%>><%=c_dmdj(i)%></option>
                <%next%>
              </select>
              </span></td>
          </tr>
          <tr>
            <td><span class="STYLE4">ģ������</span></td>
            <td><span class="STYLE4">
              <select name="mjlx" id="mjlx" onChange="CheckYl();">
                  <option value="��ͨ" <%if Rs("mjlx")="��ͨ" Then%>selected<%End If%>>��ͨ</option>
                  <option value="���湲��" <%if Rs("mjlx")="���湲��" Then%>selected<%End If%>>���湲��</option>
                  <option value="��������" <%if Rs("mjlx")="��������" Then%>selected<%End If%>>��������</option>
                  <option value="ȫ��������" <%if Rs("mjlx")="ȫ��������" Then%>selected<%End If%>>ȫ��������</option>
                  <option value="���渡�񹲼�" <%if Rs("mjlx")="���渡�񹲼�" Then%>selected<%End If%>>���渡�񹲼�</option>
                  <option value="1ģͷ+1����˫ǻģ��" <%if Rs("mjlx")="1ģͷ+1����˫ǻģ��" Then%>selected<%End If%>>1ģͷ+1����˫ǻģ��</option>
                  <option value="2ģͷ+2����˫ǻģ��" <%if Rs("mjlx")="2ģͷ+2����˫ǻģ��" Then%>selected<%End If%>>2ģͷ+2����˫ǻģ��</option>
                  <option value="����+�����干�ö���" <%if Rs("mjlx")="����+�����干�ö���" Then%>selected<%End If%>>����+�����干�ö���</option>
                  <option value="����+����ģͷ���ö���" <%if Rs("mjlx")="����+����ģͷ���ö���" Then%>selected<%End If%>>����+����ģͷ���ö���</option>
              </select>
              </span></td>
            <td><span class="STYLE4">ŷʽ����ǻ�ҩ�5</span></td>
            <td><span class="STYLE4">
              <label>
                <input name="qs" type="checkbox" id="qs" onClick="CheckYl();" value="1" <%if Rs("qs")=1 Then%>checked<%End If%>>
                ��</label>
              </span></td>
            <td><span class="STYLE4">���ں��2.8mm</span></td>
            <td><span class="STYLE4">
              <label>
                <input name="bh" type="checkbox" id="bh" onClick="CheckYl();" value="1" <%if Rs("bh")=1 Then%>checked<%End If%>>
                ��</label>
              </span></td>
          </tr>
          <tr id="cnwdh">
            <td><span class="STYLE4">����Դ�/����</span></td>
            <td><input name="tscs" type="text" id="tscs" size="10" readonly="true" value=Rs("edtscs") /></td>
            <td><span class="STYLE4">�Ͳļ���</span></td>
            <td><span class="STYLE4">
              <label>
                <input name="xcjy" type="checkbox" onClick="CheckYl();" value="1" <%if Rs("jy")=1 Then%>checked<%End If%>>
                ��</label>
              </span></td>
            <td><span class="STYLE4">�ٶ� �� 3m/min</span></td>
            <td><span class="STYLE4">
              <label>
                <input name="sd" type="checkbox" id="sd" onClick="CheckYl();" value="1" <%if Rs("sd")=1 Then%>checked<%End If%>>
                ��</label>
              </span></td>
          </tr>
          <tr>
    		  <td><span class="STYLE4">���������</span></td>
     		  <td><input name="tsyl" type="text" id="tsyl" size="10" value=<%=Rs("edtsyl")%>>
        &nbsp;Kg</td>
      <td width="15%" class="STYLE4">�������</td>
      <td class="STYLE4"><label>
          <select name="smlqk" id="smlqk" onChange="Checksml()">
            <option value="ԭ���ҷ���">ԭ���ҷ���</option>
            <option value="ԭ���Ҷ�����">ԭ���Ҷ�����</option>
            <option value="�����ϳ���">�����ϳ���</option>
          </select>
          <input id="smltxt" name="smltxt" type="text" id="smltxt" size="20" style="Display:none;" value=<%=Rs("smlqk")%>>
        </label></td>
      <td width="15%" class="STYLE4">���������</td>
      <td class="STYLE4"><p>
          <select name="gjlqk" id="gjlqk" onChange="Checkgjl()">
            <option value="ԭ���ҷ���">ԭ���ҷ���</option>
            <option value="ԭ���Ҷ�����">ԭ���Ҷ�����</option>
            <option value="�����ϳ���">�����ϳ���</option>
          </select>
          <input id="gjltxt" name="gjltxt" type="text" id="gjltxt" size="20" style="Display:none;" value=<%=Rs("gjlqk")%>>
        </p></td>
          </tr>
          <tr>
            <td><span class="STYLE4">���Է�ֵ:</span></td>
            <td colspan="5" class="STYLE4">&nbsp; <span id=fzxs class="STYLE4"></span>
              <input type="hidden" name="Cntsfz" value=0 />
              <input type="hidden" name="Cwtsfz" value=0 />
              <input type="hidden" name="Scxfz" value=0 />
              <input type="hidden" name="Zxfz" value=0 />
          </tr>
  </table>
  <table width="98%" border="2" cellpadding="0" bordercolor="#33FFCC">
    <tr>
      <th colspan="7" scope="col"><div align="left" class="STYLE4">������Ϣ</div></th>
    </tr>
    <tr>
      <td width="86"><span class="STYLE4">����Ա:</span></td>
      <td width="90"><span class="STYLE4">&nbsp;</span></td>
      <td colspan="5"><span class="STYLE4">
        <label>��Ʒ�ϸ�
          <input name="cphg" type="checkbox" disabled id="cphg" value="checkbox" />
          һ������
          <input name="checkbox2" type="checkbox" disabled value="checkbox" />
          ��������
          <input name="checkbox3" type="checkbox" disabled value="checkbox" />
          �ϴ�����֪ͨ����
          <input name="checkbox4" type="checkbox" disabled value="checkbox" />
        </label>
        </span></td>
    </tr>
    <tr>
      <td height="100"><span class="STYLE4">��ע:</span></td>
      <td height="100" colspan="6"><span class="STYLE4">&nbsp;</span></td>
    </tr>
    <tr>
      <td colspan="3"><span class="STYLE4"></span></td>
      <td width="60"><span class="STYLE4">���:</span></td>
      <td width="220"><span class="STYLE4"></span>&nbsp;</td>
      <td width="60"><span class="STYLE4">��׼:</span></td>
      <td width="220"><span class="STYLE4"></span>&nbsp;</td>
    </tr>
  </table>
  <label><br>
    <input type="submit" name="Submit2" value="��ȷ ����">
  </label>
</form>
<%End Function%>
<script language=javascript>
//��ʼ���ƻ�����ʱ��
// document.all.jhjs.value=document.all.Hijhjs.value;
 //��ʼ����ֵ
 CheckJl();
 //��ʼ����ģ��
 switch(document.all.smltxt.value){
   case "ԭ���ҷ���":
   		document.all.smlqk.selectedIndex=0;
   		document.all.smltxt.value="";
   		break;
   case "ԭ���Ҷ�����":
   		document.all.smlqk.selectedIndex=1;
   		document.all.smltxt.value="";
   		break;
   default:
   		document.all.smlqk.selectedIndex=2;
   		document.all.smltxt.style.display="";
}
 //��ʼ��������
 switch(document.all.gjltxt.value){
   case "ԭ���ҷ���":
   		document.all.gjlqk.selectedIndex=0;
   		document.all.gjltxt.value="";
   		break;
   case "ԭ���Ҷ�����":
   		document.all.gjlqk.selectedIndex=1;
   		document.all.gjltxt.value="";
   		break;
   default:
   		document.all.gjlqk.selectedIndex=2;
   		document.all.gjltxt.style.display="";
}

function CheckYl()	//�����ٶȺ�����ȷ�������
{
	CheckJl();
	var strdxcs=document.pldadd.tscs.value.split(":");
	strdxcs=strdxcs[1];
	var TmpYl
	if (document.pldadd.mz.value>0 && document.pldadd.qysd.value>0)
	{
		//���������=����*�ٶ�*60(����)*����Դ���
	TmpYl=document.pldadd.mz.value*document.pldadd.qysd.value*60*strdxcs;
	//��С˫ǻ�Ե������ϵ�Ӱ��
	if (document.pldadd.mjlx.value.indexOf("2����")>=0)
	{
		TmpYl=TmpYl*1.7;
	}
	if (document.pldadd.mjlx.value.indexOf("1����")>=0)
	{
		TmpYl=TmpYl*1.5;
	}
	document.pldadd.tsyl.value=Math.round(TmpYl);
	}
	else
	{
	document.pldadd.tsyl.value="";
	}
}

function Checksml() //��д�����ϳ���
{
	if (document.all.smlqk.value=="�����ϳ���")
 {
  	document.all.smltxt.style.display="";
  }
	else
 {
  document.all.smltxt.style.display="none";
  document.all.smltxt.value="";
  }
}
function Checkgjl() //��д�����ϳ���
{
	if (document.all.gjlqk.value=="�����ϳ���")
 {
  	document.all.gjltxt.style.display="";
  }
	else
 {
  document.all.gjltxt.style.display="none";
  document.all.gjltxt.value="";
  }
}

function CheckJl() //ģ�ߵ��Խ����ȼ����㺯��
{
	var Dm = new Array();
	var Jl = new Array();
	var key;
<%
	strSql="select * from sys_jldj order by ID"
	set rs=server.createobject("adodb.recordset")
	rs.open strSql,conn,1,1
	i=0
	do while not rs.eof
%>
Dm[<%=i%>]="<%=rs("cjqy") & "," & rs("dmdj")%>";
Jl[<%=i%>]="<%=rs("mtcs") & "," & rs("dxcs") & "," & rs("cntsjl") & "," & rs("scxjl") & "," & rs("cwtsjl") & "," & rs("cwtsts") & "," & rs("zxjl")%>";
<%
	i = i + 1
	rs.movenext
	loop
	rs.close
%>
	var JcCs="<%=TmpTscs%>";
		JcCs=JcCs.split(":");
	var Mtcs_inc1 = 0;	//���Դ�������2(ģ�����ͶԵ��Դ�����Ӱ��)
	var Mtcs_inc2 = 0;	//���Դ�������2(�Ͳļ����Ե��Դ�����Ӱ��)

	var Tsfz_inc1 = 0;	//��ֵ����2(ģ�����ͶԳ��ڷ�ֵ��Ӱ��+)
	var Tsfz_inc2 = 1;	//��ֵ����1(�ٶȶԳ��ڷ�ֵ��Ӱ��*)
	var Tsfz_inc3 = 0;	//��ֵ����3(ǻ�ҶԷ�ֵ��Ӱ��+)
	var Tsfz_inc4 = 0;	//��ֵ����4(�ں�Է�ֵ��Ӱ��+)
	var Tsfz_inc5 = 1;	//��ֵ����4(�������ݶԷ�ֵ��Ӱ��*)

	switch (document.pldadd.mjlx.selectedIndex){	//ģ�����ͶԵ��Դ�������ֵ��װ���ֵ��Ӱ��
		case 0 : Mtcs_inc1=0;Tsfz_inc1=0;break;
		case 1 : Mtcs_inc1=0;Tsfz_inc1=0.1;break;
		case 2 : Mtcs_inc1=0;Tsfz_inc1=0.1;break;
		case 3 : Mtcs_inc1=0;Tsfz_inc1=0.3;break;
		case 4 : Mtcs_inc1=0;Tsfz_inc1=0.2;break;
		case 5 : Mtcs_inc1=0;Tsfz_inc1=0.4;break;
		case 6 : Mtcs_inc1=0;Tsfz_inc1=0.7;break;
		case 7 : Mtcs_inc1=0;Tsfz_inc1=0.2;break;
		case 8 : Mtcs_inc1=0;Tsfz_inc1=0.4;break;
	}
	if(document.all.sd.checked==false){		//�ٶȶԵ��Դ�������ֵ��Ӱ��
		Tsfz_inc2=0.84;
	}
	if(document.all.qs.checked==true){		//ǻ��>5�Ե��Դ�������ֵ��Ӱ��
		Tsfz_inc3=0.1;
	}
	if(document.all.bh.checked==true){		//�ں�Ե��Դ�������ֵ��Ӱ��
		Tsfz_inc4=0.1;
	}
	if (document.pldadd.xcjy.checked==true){	//�Ͳļ����Ե��Դ�����Ӱ��
		Mtcs_inc2=2;
	}
	var Ftxs=1;//�ǵ�ϵ���������ж��Ƿ���Ҫ�ڳ�����Է�ֵ���ټӳ��ڷ�ֵ��һ�롣
	var Mtcs=Mtcs_inc1 + Mtcs_inc2;//���Դ�������

	for (key in Dm){
		if (Dm[key]==document.pldadd.cjqy.value + "," + document.pldadd.dmdj.value){//ģ�ߵȼ�������ȼ�
			var arrTmp = Jl[key].split(",");
			switch (document.pldadd.rwnr.value){//�������ͶԵ��Դ����ͷ�ֵ��Ӱ��
				case "ȫ��" :
					document.pldadd.tscs.value=Number(JcCs[0]) + Mtcs + ":" + (Number(JcCs[1]) + Mtcs);break;
					case "��ͷ" :
					document.pldadd.tscs.value=Number(JcCs[0]) + Mtcs + ":0";
					arrTmp[2]=Math.round(arrTmp[2]*0.3);
					arrTmp[3]=Math.round(arrTmp[3]*0.3);
					arrTmp[6]=10;		//װ���ֵ
					Tsfz_inc5=0.3;
					break;
				case "����" :
					document.pldadd.tscs.value="0:" + (Number(JcCs[1]) + Mtcs);
					arrTmp[2]=Math.round(arrTmp[2]*0.7);
					arrTmp[3]=Math.round(arrTmp[3]*0.7);
					arrTmp[6]=0;		//װ���ֵ
					Tsfz_inc5=0.7;
					break;
				case "��Ʒ" :
					document.pldadd.tscs.value="0:0";
					arrTmp[2]=Math.round(arrTmp[2]*1.7);
					arrTmp[3]=Math.round(arrTmp[3]*1.7);
					arrTmp[4]=Math.round(arrTmp[4]*1.7);
					break;
				case "�ֵ�" :
					document.pldadd.tscs.value=Number(JcCs[0]) + Mtcs_inc1 + ":" + (Number(JcCs[1]) + Mtcs);
					arrTmp[2]=Math.round(arrTmp[2]*0.6);
					arrTmp[3]=Math.round(arrTmp[3]*0.6);
					arrTmp[4]=Math.round(arrTmp[4]*0.6);
			     	arrTmp[6]=Math.round(arrTmp[6]*0.6);	//װ���ֵ
					break;
				case "�ǵ�" :
					document.pldadd.tscs.value="0:0";
					Ftxs=0;
					arrTmp[3]=Math.round(arrTmp[3]*0);
					arrTmp[6]=35;		//װ���ֵ
					break;
				case "����" :
					document.pldadd.tscs.value=Number(JcCs[0]) + Mtcs + ":" + (Number(JcCs[1]) + Mtcs);
					arrTmp[2]=Math.round(arrTmp[2]*0.3);
					arrTmp[3]=Math.round(arrTmp[3]*0.3);
					arrTmp[4]=Math.round(arrTmp[4]*0.3);
					break;
				case "TB" :
					document.pldadd.tscs.value="0:0";
					arrTmp[2]=Math.round(arrTmp[2]*0);
					arrTmp[3]=Math.round(arrTmp[3]*0);
					arrTmp[6]=20;		//װ���ֵ
					break;
				case "TA" :
					document.pldadd.tscs.value=Number(JcCs[0]) + Mtcs + ":" + (Number(JcCs[1]) + Mtcs);
					if (document.pldadd.dmdj.selectedIndex<5){	//װ���ֵ
				 		arrTmp[6]=140;
					 }
					else{
				 		arrTmp[6]=90;
					}
					break;
			}
			var StrCntsfz=Math.round(arrTmp[2]*Tsfz_inc2*Tsfz_inc5*(1+Tsfz_inc1+Tsfz_inc3+Tsfz_inc4));
			var StrScxfz=1;
			var StrCwtsfz=1;
			var StrCntsfz=StrCntsfz*Ftxs;
			var StrZxfz=1;
			fzxs.innerHTML="���ڵ���:" + StrCntsfz + ";������:" + StrScxfz + ";װ��:" + StrZxfz + ";�������:" + StrCwtsfz;
			document.pldadd.Cntsfz.value=StrCntsfz;
			document.pldadd.Scxfz.value=StrScxfz;
			document.pldadd.Cwtsfz.value=StrCwtsfz;
			document.pldadd.Zxfz.value=StrZxfz;
		}
	}
}
</script>
