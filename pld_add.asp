<!--#include file="Inc/mission.asp" -->
<!--#include file="Inc/Sysconn.asp" -->
<!--#include file="inc/Skin_CSS.asp"-->
<!--#include file="dmlx.asp"-->
<SCRIPT language=Javascript src="Inc/jsfunc.js"></SCRIPT>
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
<%
if  session("UserName")="" then
	response.redirect "UserServer.asp"
end if

%>
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
      <form id='pldadd' name='pldadd'  action=pld_indb.asp?action=add method=post onSubmit='return CheckM();'>
        <table width="98%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="49%" height="18">���ƻ�����&gt;&gt;
              ���������
            <td width="71%"><div align="right">&nbsp; </div></td>
          </tr>
          <tr>
            <td height="1" colspan="2" bgcolor="#D6DCE3"></td>
          </tr>
          <tr>
            <td height="1" colspan="2">&nbsp;</td>
          </tr>
        </table>
        <table id="rwslx" width="98%" border="1" cellspacing="0" bordercolor="#33FFCC">
          <caption>
          <span id=span_rwlx class="STYLE2">����ģ�߳�����ģ����������</span> <br />
          </caption>
          <tr>
            <th colspan="6" scope="col"><div align="left" class="STYLE4">ģ����Ϣ</div></th>
          </tr>
          <tr>
            <td width="100"><span class="STYLE4">������</span></td>
            <td><span class="STYLE4">
              <label>
                <input name="ddh" type="text" id="ddh" size="15" />
              </label>
              </span></td>
            <td width="100"><span class="STYLE4">��ˮ��</span></td>
            <td><input name="lsh" type="text" id="lsh" size="15" /></td>
            <td width="100"><span class="STYLE4">�ͻ�����</span></td>
            <td><input name="khmc" type="text" id="khmc" size="15" /></td>
          </tr>
          <tr>
            <td><span class="STYLE4">������</span></td>
            <td><input name="scxh" type="text" id="scxh" size="10" /></td>
            <td><span class="STYLE4">ǣ���ٶ�</span></td>
            <td><p class="STYLE4">
                <input name="qysd" type="text" id="qysd" size="10" onKeyPress="javascript:validationNumber(this, 'u_float', 100, '');" onChange="CheckYl();">
                m/min <br />
              </p></td>
            <td><span class="STYLE4">����</span></td>
            <td><input name="mz" type="text" id="mz" size="10" onKeyPress="javascript:validationNumber(this, 'u_float', 100, '');" onChange="CheckYl();">
              Kg/m</td>
          </tr>
          <tr>
            <td><span class="STYLE4">�����ȼ�</span></td>
            <td><span class="STYLE4">
              <label>
                <select name="jsdj" id="jsdj" onChange="CheckYl();">
                  <option value="A��">A��</option>
                  <option value="B��">B��</option>
                  <option value="C��">C��</option>
                </select>
              </label>
              </span></td>
            <td><span class="STYLE4">ģ�ߵȼ�</span></td>
            <td><span class="STYLE4">
              <label>
                <select name="mjdj" id="mjdj" onChange="CheckYl();">
                  <option value="A��">A��</option>
                  <option value="B��">B��</option>
                  <option value="C��">C��</option>
                </select>
              </label>
              </span></td>
            <!--<td onMouseover="GradientClose();" onMouseMove="MenuClick();"><span class="STYLE4">����ȼ�</span></td>
-->
            <td><span class="STYLE4">����ȼ�</span></td>
            <td><span class="STYLE4">
              <select name="dmdj" id="dmdj" onChange="CheckYl();">
                <%for i = 0 to ubound(c_dmdj)%>
                <option value='<%=c_dmdj(i)%>'><%=c_dmdj(i)%></option>
                <%next%>
              </select>
              </span></td>
          </tr>
          <tr>
            <td><span class="STYLE4">��������</span></td>
            <td><select name="rwnr" id="rwnr" onChange="CheckYl();">
                <option value="ȫ��">ȫ��</option>
                <option value="��ͷ">��ͷ</option>
                <option value="����">����</option>
                <option value="��Ʒ">��Ʒ</option>
                <option value="�ֵ�">�ֵ�</option>
                <option value="�ǵ�">�ǵ�</option>
                <option value="TB">TB</option>
                <option value="TA">TA</option>
              </select></td>
            <td><span class="STYLE4">ģ������</span></td>
            <td><span class="STYLE4">
              <select name="mjlx" id="mjlx" onChange="CheckYl();">
                <%
                Dim c_mjlx
                c_mjlx=""
				'ȡ��ģ������
				strSql="select * from mjlx"
				set rs=server.createobject("adodb.recordset")
				rs.open strSql,conn,1,1
				do while not rs.eof
				if c_mjlx <> "" then
					c_mjlx = c_mjlx & "|||" & rs("lx")
				else
					c_mjlx = rs("lx")
				end if
				rs.movenext
				loop
				rs.close
				c_mjlx = split(c_mjlx, "|||")
                for i = 0 to ubound(c_mjlx)%>
                <option value='<%=c_mjlx(i)%>'><%=c_mjlx(i)%></option>
                <%next%>
              </select>
              </span></td>
            <td><span class="STYLE4">���ں�2.7mm</span></td>
            <td><span class="STYLE4">
              <label>
                <input name="bh" type="checkbox" id="bh" onClick="CheckYl();" value="1" />
                ��</label>
              </span></td>
          </tr>
          <tr id="cnwdh">
            <td><span class="STYLE4">�Ͳļ���</span></td>
            <td><span class="STYLE4">
              <label>
                <input name="xcjy" type="radio" onClick="CheckYl();" value="1" />
                ��</label>
              <label>
                <input name="xcjy" type="radio" onClick="CheckYl();" value="0" checked="checked" />
                ��</label>
              </span></td>
            <td><span class="STYLE4">���������</span></td>
            <td><input name="tsyl" type="text" id="tsyl" size="10">
              &nbsp;Kg</td>
            <td><span class="STYLE4">����Դ�/����</span></td>
            <td><input name="tscs" type="text" id="tscs" size="10" readonly="true" /></td>
          </tr>
          <tr>
            <td class="STYLE4">�������:</td>
            <td class="STYLE4"><select name="smlqk" id="smlqk" onChange="Checksml()">
                <option value="ԭ���ҷ���">ԭ���ҷ���</option>
                <option value="ԭ���Ҷ�����">ԭ���Ҷ�����</option>
                <option value="�����ϳ���">�����ϳ���</option>
              </select>
              <input id="smltxt" name="smltxt" type="text" size="20" style="Display:none;"></td>
            <td class="STYLE4">���������:</td>
            <td class="STYLE4"><select name="gjlqk" id="gjlqk" onChange="Checkgjl()">
                <option value="ԭ���ҷ���">ԭ���ҷ���</option>
                <option value="ԭ���Ҷ�����">ԭ���Ҷ�����</option>
                <option value="�����ϳ���">�����ϳ���</option>
              </select>
              <input id="gjltxt" name="gjltxt" type="text" id="gjltxt" size="20" style="Display:none;"></td>
            <td class="STYLE4">�ͻ�����:</td>
            <td class="STYLE4"><select name="cjqy" id="cjqy">
                <option value="����">����</option>
                <option value="����" selected>����</option>
              </select></td>
          </tr>
          <tr>
            <td><span class="STYLE4">ǻ��</span></td>
            <td><span class="STYLE4">
              <label>
                <input name="qs" type="checkbox" id="qs" onClick="CheckYl();" value="1" />
                > 4</label>
              </span></td>
            <td><span class="STYLE4">���Է�ֵ:</span></td>
            <td colspan="3" class="STYLE4">&nbsp; <span id=fzxs class="STYLE4"></span>
              <input type="hidden" name="Cntsfz" value=0 />
              <input type="hidden" name="Cwtsfz" value=0 />
              <input type="hidden" name="Scxfz" value=0 />
              <input type="hidden" name="Zxfz" value=0 /></td>
          </tr>
        </table>
        <table width="98%" border="1" cellspacing="0" bordercolor="#33FFCC">
          <tr>
            <th colspan="8" scope="col"><div align="left" class="STYLE4">��Ա��Ϣ</div></th>
          </tr>
          <tr>
            <td align="center"><span class="STYLE4">ID</span>
              </th>
            <td align="center"><span class="STYLE4">����Ա</span>
              </th>
            <td align="center"><span class="STYLE4">��������</span>
              </th>
            <td align="center"><span class="STYLE4">������ʽ</span>
              </th>
          </tr>
          <tr>
            <td><div align="center"><span class="STYLE4">&nbsp;</span></div></td>
            <td><div align="center"><span class="STYLE4">&nbsp;</span></div></td>
            <td><div align="center"><span class="STYLE4">&nbsp;</span></div></td>
            <td><div align="center"><span class="STYLE4">&nbsp;</span></div></td>
          </tr>
        </table>
        <table width="98%" border="1" cellspacing="0" bordercolor="#33FFCC">
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
        <label> <br>
          <input type="submit" name="Submit2" value="��ȷ ����">
        </label>
      </form></td>
  </tr>
</table>
<!-- #include file="Inc/Foot.asp" -->
<map name="Map">
  <area shape="rect" coords="18,14,69,24" href="News.asp">
</map>
</BODY>
</HTML><script language=javascript>
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
	if (document.pldadd.mjlx.value.indexOf("��˫ǻ")>=0)
	{
		TmpYl=TmpYl*1.7;
	}
	if (document.pldadd.mjlx.value.indexOf("С˫ǻ")>=0)
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
	var Cnxs = new Array();
	var Cwxs = new Array();
	var key;
<%
	strSql="select * from sys_jldj order by ID"
	set rs=server.createobject("adodb.recordset")
	rs.open strSql,conn,1,1
	i=0
	do while not rs.eof
%>
Dm[<%=i%>]="<%=rs("dmdj") & "," & rs("jsdj")%>";
Jl[<%=i%>]="<%=rs("mtcs") & "," & rs("dxcs") & "," & rs("cntsjl") & "," & rs("scxjl") & "," & rs("cwtsjl") & "," & rs("cwtsts") & "," & rs("zxjl")%>";
<%
	i = i + 1
	rs.movenext
	loop
	rs.close

	strSql="select * from sys_mjdj order by ID"
	set rs=server.createobject("adodb.recordset")
	rs.open strSql,conn,1,1
	i=0
	do while not rs.eof
%>
Cnxs[<%=i%>]="<%=rs("Cnxs")%>";
Cwxs[<%=i%>]="<%=rs("Cwxs")%>";
<%
		i = i + 1
		rs.movenext
	loop
	rs.close
%>

	var Mtcs_inc1 = 0	//ģͷ��������1(ģ�ߵȼ���ģͷ���Դ�����Ӱ��)
	var Mtcs_inc2 = 0	//ģͷ��������2(ģ�����Ͷ�ģͷ���Դ�����Ӱ��)
	var Mtcs_inc3 = 0	//ģͷ��������3(�ں��ģͷ���Դ�����Ӱ��)
	var Mtcs_inc4 = 0	//ģͷ��������4(������ģͷ���Դ�����Ӱ��)

	var Dxcs_inc1 = 0	//���ʹ�������1(ģ�ߵȼ��Զ��͵��Դ�����Ӱ��)
	var Dxcs_inc2 = 0	//���ʹ�������2(ģ�����ͶԶ��͵��Դ�����Ӱ��)
	var Dxcs_inc3 = 0	//���ʹ�������3(�ں�Զ��͵��Դ�����Ӱ��)
	var Dxcs_inc4 = 0	//���ʹ�������4(�����Զ��͵��Դ�����Ӱ��)
	var Dxcs_inc5 = 0	//���ʹ�������5(ǻ�Ҵ���4��ģͷ���Դ�����Ӱ��)

	var Tsfz_inc1 = 0	//��ֵ����1(ģ�ߵȼ��Գ��ڷ�ֵ��Ӱ��)
	var Tsfz_inc2 = 0	//��ֵ����2(ģ�ߵȼ��Գ����ֵ��Ӱ��)
	var Tsfz_inc3 = 0	//��ֵ����3(ģ�����ͶԷ�ֵ��Ӱ��)
	var Tsfz_inc4 = 0	//��ֵ����4(�ں�Է�ֵ��Ӱ��)
	var Tsfz_inc5 = 0	//��ֵ����5(ģ�ߵȼ��������Ķ�װ���ֵ��Ӱ��)
	var Tsfz_inc6 = 1	//��ֵ����6(ģ�����ͼ�ǻ����װ���ֵ��Ӱ��)
	var Tsfz_inc7 = 1	//��ֵ����7(�������ݶԳ����ֵ��Ӱ��)
	var Tsfz_inc8 = 0	//��ֵ����8(ǻ�Ҵ���4�Է�ֵ��Ӱ��)

	if (document.pldadd.dmdj.selectedIndex<4 && document.pldadd.mjlx.value.indexOf("˫ǻ")>0){
		Tsfz_inc6=1.732;
	}

	switch (document.pldadd.mjdj.selectedIndex){//ģ�ߵȼ��Ե��Դ��������Է�ֵ��װ���ֵ��Ӱ��
		case 0 : Mtcs_inc1=0;Dxcs_inc1=1;Tsfz_inc1=Cnxs[0];Tsfz_inc2=Cwxs[0];
				 if (document.pldadd.dmdj.selectedIndex<5){
				 	Tsfz_inc5=35;
				 }
				 else{
				 	Tsfz_inc5=25;
				}
				break;
		case 1 : Mtcs_inc1=0;Dxcs_inc1=0;Tsfz_inc1=Cnxs[1];Tsfz_inc2=Cwxs[1];break;
		case 2 : Mtcs_inc1=0;Dxcs_inc1=0;Tsfz_inc1=Cnxs[2];Tsfz_inc2=Cwxs[2];break;
	}
	if (document.pldadd.dmdj.selectedIndex==8){
	 	Tsfz_inc1=1;
	 	Tsfz_inc2=1;
	 	Tsfz_inc5=0;
	 }
	//���ģװ���ֵ�̶�Ϊ1000���������ģ�ߵȼ��ı仯


	switch (document.pldadd.mjlx.selectedIndex){//ģ�����ͶԵ��Դ�������ֵ��װ���ֵ��Ӱ��
		case 0 : Mtcs_inc2=0;Dxcs_inc2=0;Tsfz_inc3=0;break;
		case 1 : Mtcs_inc2=0;Dxcs_inc2=2;Tsfz_inc3=0.5;Tsfz_inc6=2;break;
		case 2 : Mtcs_inc2=1;Dxcs_inc2=1;Tsfz_inc3=0.7;Tsfz_inc6=2;break;
		case 3 : Mtcs_inc2=0;Dxcs_inc2=1;Tsfz_inc3=0.3;break;
		case 4 : Mtcs_inc2=0;Dxcs_inc2=1;Tsfz_inc3=0.25;break;
		case 5 : Mtcs_inc2=0;Dxcs_inc2=1;Tsfz_inc3=0.1;break;
		case 6 : Mtcs_inc2=1;Dxcs_inc2=1;Tsfz_inc3=0.2;break;
		case 7 : Mtcs_inc2=0;Dxcs_inc2=1;Tsfz_inc3=0.15;break;
		case 8 : Mtcs_inc2=0;Dxcs_inc2=0;Tsfz_inc3=0.05;break;
		case 9 : Mtcs_inc2=1;Dxcs_inc2=1;Tsfz_inc3=0.4;break;
	}

	if(document.all.bh.checked==true){		//�ں�Ե��Դ�������ֵ��Ӱ��
		Mtcs_inc3=0;Dxcs_inc3=1;Tsfz_inc4=0.1;
	}

	if (document.pldadd.xcjy[0].checked){	//�Ͳļ����Ե��Դ�����Ӱ��
		Mtcs_inc4=1;Dxcs_inc4=1;
	}

	if(document.all.qs.checked==true){		//ǻ��>4�Ե��Դ�������ֵ��Ӱ��
		Dxcs_inc5=1;Tsfz_inc8=0.1;
	}

	var Mtcs=Mtcs_inc1 + Mtcs_inc2 + Mtcs_inc3 + Mtcs_inc4;	//ģͷ���Դ�������
	var Dxcs=Dxcs_inc1 + Dxcs_inc2 + Dxcs_inc3 + Dxcs_inc4 + Dxcs_inc5;	//���͵��Դ�������
	var Ftxs=1;//�ǵ�ϵ���������ж��Ƿ���Ҫ�ڳ�����Է�ֵ���ټӳ��ڷ�ֵ��һ�롣
	for (key in Dm){
		if (Dm[key]==document.pldadd.dmdj.value + "," + document.pldadd.jsdj.value){//����ȼ��������ȼ�
			var arrTmp = Jl[key].split(",");

			switch (document.pldadd.rwnr.value){//�������ͶԵ��Դ����ͷ�ֵ��Ӱ��
				case "ȫ��" :
				document.pldadd.tscs.value=Number(arrTmp[0]) + Mtcs + ":" + (Number(arrTmp[1]) + Dxcs);break;
				case "��ͷ" :
				document.pldadd.tscs.value=Number(arrTmp[0]) + Mtcs + ":0";
				arrTmp[2]=Math.round(arrTmp[2]*0.3);
				arrTmp[3]=Math.round(arrTmp[3]*0.3);
				arrTmp[6]=10;Tsfz_inc5=0;Tsfz_inc6=1;		//װ���ֵ
				Tsfz_inc7=0.3;
				break;
				case "����" :
				document.pldadd.tscs.value="0:" + (Number(arrTmp[1]) + Dxcs);
				arrTmp[2]=Math.round(arrTmp[2]*0.7);
				arrTmp[3]=Math.round(arrTmp[3]*0.7);
				arrTmp[6]=0;Tsfz_inc5=0;Tsfz_inc6=1;		//װ���ֵ
				Tsfz_inc7=0.7;
				break;
				case "��Ʒ" :
					document.pldadd.tscs.value="0:0";
					arrTmp[2]=Math.round(arrTmp[2]*1.7);
					arrTmp[3]=Math.round(arrTmp[3]*1.7);
					arrTmp[4]=Math.round(arrTmp[4]*1.7);
					break;
				case "�ֵ�" :
					document.pldadd.tscs.value=Number(arrTmp[0]) + Mtcs + ":" + (Number(arrTmp[1]) + Dxcs);
					arrTmp[2]=Math.round(arrTmp[2]*0.6);
					arrTmp[3]=Math.round(arrTmp[3]*0.6);
					arrTmp[4]=Math.round(arrTmp[4]*0.6);
					arrTmp[6]=Math.round(arrTmp[6]*0.6);								//װ���ֵ
					break;
				case "�ǵ�" :
					document.pldadd.tscs.value="0:0";
//					arrTmp[2]=Math.round(arrTmp[2]*0);
					Ftxs=0;
					arrTmp[3]=Math.round(arrTmp[3]*0);
//					arrTmp[4]=Math.round(arrTmp[4]*0.7);
					arrTmp[6]=35;Tsfz_inc5=0;Tsfz_inc6=1;		//װ���ֵ
					break;
				case "TB" :
					document.pldadd.tscs.value="0:0";
					arrTmp[2]=Math.round(arrTmp[2]*0);
					arrTmp[3]=Math.round(arrTmp[3]*0);
					arrTmp[6]=20;Tsfz_inc5=0;Tsfz_inc6=1;		//װ���ֵ
					break;
				case "TA" :
					document.pldadd.tscs.value=Number(arrTmp[0]) + Mtcs + ":" + (Number(arrTmp[1]) + Dxcs);
					if (document.pldadd.dmdj.selectedIndex<5){	//װ���ֵ
				 		arrTmp[6]=140;
					 }
					else{
				 		arrTmp[6]=90;
					}
					Tsfz_inc5=0;Tsfz_inc6=1;
					break;
			}

//			if (document.pldadd.rwlx[1].checked){	//������ԶԵ��Դ�����Ӱ��
//				document.pldadd.tscs.value=arrTmp[5]
//			}

			var StrCntsfz=Math.round(arrTmp[2]*Tsfz_inc1*(1+Tsfz_inc3+Tsfz_inc4+Tsfz_inc8));
			var StrScxfz=Math.round(arrTmp[3]*Tsfz_inc1*(1+Tsfz_inc3+Tsfz_inc4+Tsfz_inc8));
//			var StrCwtsfz=Math.round(arrTmp[4]*Tsfz_inc2*(1+Tsfz_inc3+Tsfz_inc4+Tsfz_inc8));
			var StrCwtsfz=Math.round((arrTmp[4]*1*(1+Tsfz_inc3+Tsfz_inc4+Tsfz_inc8)-StrCntsfz*0.5*(Ftxs-1))*Tsfz_inc7);
			var StrCntsfz=StrCntsfz*Ftxs;
			var StrZxfz=Math.round(arrTmp[6]*Tsfz_inc6+Tsfz_inc5);
			fzxs.innerHTML="���ڵ���:" + StrCntsfz + ";������:" + StrScxfz + ";װ��:" + StrZxfz + ";�������:" + StrCwtsfz;
			document.pldadd.Cntsfz.value=StrCntsfz;
			document.pldadd.Scxfz.value=StrScxfz;
			document.pldadd.Cwtsfz.value=StrCwtsfz;
			document.pldadd.Zxfz.value=StrZxfz;
		}
	}
}

CheckJl();
</script>
