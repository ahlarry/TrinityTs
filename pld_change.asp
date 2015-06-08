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
	'取出断面等级
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
          <td width="29%" height="18">　计划任务&gt;&gt;更改任务书</td>
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
            <td>&nbsp;&nbsp;输入流水号:
              <input type=text name="s_lsh" size=15 value="<%=Trim(Request("s_lsh"))%>">
              <input type="submit" value=" 查 找 "></td>
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
	If x="" Then Response.Write("请输入要更改的任务书的流水号!<br>") : Exit Function
	strSql="select * from [mission] where [lsh] like '%"&x&"%'"
	set Rs=server.createobject("adodb.recordset")
	Rs.open strSql,Conn,1,3
	If Rs.Eof Or Rs.Bof Then
		errmsg="流水号为 【 " & x & " 】 的任务书不存在!"
		Call WriteErrMsg()
	Else
		If Rs.RecordCount>1 Then
			response.redirect "intpld_list.asp?lsh="&x
		else
			If Not(IsNull(Rs("rksj"))) Then
					errmsg="流水号为 【 " & x & " 】 的任务书已经完成!"
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
    <span class="STYLE2">更改流水号为<%=Rs("lsh")%>的任务书</span><br />
    </caption>
          <tr>
            <th colspan="6" scope="col"><div align="left" class="STYLE4">模具信息</div></th>
          </tr>
          <tr>
            <td width="100"><span class="STYLE4">流水号</span></td>
            <td><input name="lsh" type="text" id="lsh" size="15" value=<%=Rs("lsh")%> /></td>
            <td width="100"><span class="STYLE4">订单号</span></td>
            <td><span class="STYLE4">
              <label>
                <input name="ddh" type="text" id="ddh" size="15" value=<%=Rs("ddh")%> />
              </label>
              </span></td>
            <td width="100"><span class="STYLE4">客户名称</span></td>
            <td><input name="khmc" type="text" id="khmc" size="15" value=<%=Rs("khmc")%> /></td>
          </tr>
          <tr>
            <td><span class="STYLE4">生产线</span></td>
            <td><select name="scxh" id="scxh">
                <option value="0" <%if Rs("scx")="0" Then%>selected<%End If%>>金湖65</option>
                <option value="1" <%if Rs("scx")="1" Then%>selected<%End If%>>韦柏114</option>
                <option value="2" <%if Rs("scx")="2" Then%>selected<%End If%>>AROGOS93</option>
                <option value="3" <%if Rs("scx")="3" Then%>selected<%End If%>>辛辛那提58 </option>
                <option value="4" <%if Rs("scx")="4" Then%>selected<%End If%>>辛辛那提112</option>
                <option value="5" <%if Rs("scx")="5" Then%>selected<%End If%>>克劳斯玛菲90</option>
              </select></td>
            <td><span class="STYLE4">牵引速度</span></td>
            <td><p class="STYLE4">
                <input name="qysd" type="text" id="qysd" size="10" value=<%=Rs("qysd")%> onKeyPress="javascript:validationNumber(this, 'u_float', 100, '');" onChange="CheckYl();">
                m/min <br />
              </p></td>
            <td><span class="STYLE4">米重</span></td>
            <td><input name="mz" type="text" id="mz" size="10" value=<%=Rs("mz")%>  onKeyPress="javascript:validationNumber(this, 'u_float', 100, '');" onChange="CheckYl();">
              Kg/m</td>
          </tr>
          <tr>
            <td><span class="STYLE4">任务类型</span></td>
            <td><select name="rwnr" id="rwnr" onChange="CheckYl();">
                <option value="全套" <%if Rs("rwlx")="全套" Then%>selected<%End If%>>全套</option>
                <option value="机头" <%if Rs("rwlx")="机头" Then%>selected<%End If%>>机头</option>
                <option value="定型" <%if Rs("rwlx")="定型" Then%>selected<%End If%>>定型</option>
                <option value="新品" <%if Rs("rwlx")="新品" Then%>selected<%End If%>>新品</option>
                <option value="粗调" <%if Rs("rwlx")="粗调" Then%>selected<%End If%>>粗调</option>
                <option value="非调" <%if Rs("rwlx")="非调" Then%>selected<%End If%>>非调</option>
                <option value="修理" <%if Rs("rwlx")="修理" Then%>selected<%End If%>>修理</option>
                <option value="TB" <%if Rs("rwlx")="TB" Then%>selected<%End If%>>TB</option>
                <option value="TA" <%if Rs("rwlx")="TA" Then%>selected<%End If%>>TA</option>
              </select></td>
            <td><span class="STYLE4">客户类别</span></td>
            <td><span class="STYLE4">
                <select name="cjqy" id="cjqy" onChange="CheckYl();">
       	         <%for i = 0 to ubound(c_cjqy)%>
         	       <option value='<%=c_cjqy(i)%>' <%if Rs("cjqy")=c_cjqy(i) Then%>selected<%End If%>><%=c_cjqy(i)%></option>
                <%next%>
                </select>
              </span></td>
            <td><span class="STYLE4">断面等级</span></td>
            <td><span class="STYLE4">
              <select name="dmdj" id="dmdj" onChange="CheckYl();">
                <%for i = 0 to ubound(c_dmdj)%>
                	<option value='<%=c_dmdj(i)%>' <%if Rs("dmdj")=c_dmdj(i) Then%>selected<%End If%>><%=c_dmdj(i)%></option>
                <%next%>
              </select>
              </span></td>
          </tr>
          <tr>
            <td><span class="STYLE4">模具类型</span></td>
            <td><span class="STYLE4">
              <select name="mjlx" id="mjlx" onChange="CheckYl();">
                  <option value="普通" <%if Rs("mjlx")="普通" Then%>selected<%End If%>>普通</option>
                  <option value="表面共挤" <%if Rs("mjlx")="表面共挤" Then%>selected<%End If%>>表面共挤</option>
                  <option value="胶条共挤" <%if Rs("mjlx")="胶条共挤" Then%>selected<%End If%>>胶条共挤</option>
                  <option value="全包覆共挤" <%if Rs("mjlx")="全包覆共挤" Then%>selected<%End If%>>全包覆共挤</option>
                  <option value="表面浮雕共挤" <%if Rs("mjlx")="表面浮雕共挤" Then%>selected<%End If%>>表面浮雕共挤</option>
                  <option value="1模头+1定型双腔模具" <%if Rs("mjlx")="1模头+1定型双腔模具" Then%>selected<%End If%>>1模头+1定型双腔模具</option>
                  <option value="2模头+2定型双腔模具" <%if Rs("mjlx")="2模头+2定型双腔模具" Then%>selected<%End If%>>2模头+2定型双腔模具</option>
                  <option value="白料+共挤板共用定型" <%if Rs("mjlx")="白料+共挤板共用定型" Then%>selected<%End If%>>白料+共挤板共用定型</option>
                  <option value="白料+共挤模头共用定型" <%if Rs("mjlx")="白料+共挤模头共用定型" Then%>selected<%End If%>>白料+共挤模头共用定型</option>
              </select>
              </span></td>
            <td><span class="STYLE4">欧式主材腔室﹥5</span></td>
            <td><span class="STYLE4">
              <label>
                <input name="qs" type="checkbox" id="qs" onClick="CheckYl();" value="1" <%if Rs("qs")=1 Then%>checked<%End If%>>
                是</label>
              </span></td>
            <td><span class="STYLE4">主壁厚≥2.8mm</span></td>
            <td><span class="STYLE4">
              <label>
                <input name="bh" type="checkbox" id="bh" onClick="CheckYl();" value="1" <%if Rs("bh")=1 Then%>checked<%End If%>>
                是</label>
              </span></td>
          </tr>
          <tr id="cnwdh">
            <td><span class="STYLE4">额定调试次/天数</span></td>
            <td><input name="tscs" type="text" id="tscs" size="10" readonly="true" value=Rs("edtscs") /></td>
            <td><span class="STYLE4">型材寄样</span></td>
            <td><span class="STYLE4">
              <label>
                <input name="xcjy" type="checkbox" onClick="CheckYl();" value="1" <%if Rs("jy")=1 Then%>checked<%End If%>>
                是</label>
              </span></td>
            <td><span class="STYLE4">速度 ≥ 3m/min</span></td>
            <td><span class="STYLE4">
              <label>
                <input name="sd" type="checkbox" id="sd" onClick="CheckYl();" value="1" <%if Rs("sd")=1 Then%>checked<%End If%>>
                是</label>
              </span></td>
          </tr>
          <tr>
    		  <td><span class="STYLE4">额定调试用料</span></td>
     		  <td><input name="tsyl" type="text" id="tsyl" size="10" value=<%=Rs("edtsyl")%>>
        &nbsp;Kg</td>
      <td width="15%" class="STYLE4">粉料情况</td>
      <td class="STYLE4"><label>
          <select name="smlqk" id="smlqk" onChange="Checksml()">
            <option value="原厂家粉料">原厂家粉料</option>
            <option value="原厂家二次料">原厂家二次料</option>
            <option value="代用料厂家">代用料厂家</option>
          </select>
          <input id="smltxt" name="smltxt" type="text" id="smltxt" size="20" style="Display:none;" value=<%=Rs("smlqk")%>>
        </label></td>
      <td width="15%" class="STYLE4">共挤料情况</td>
      <td class="STYLE4"><p>
          <select name="gjlqk" id="gjlqk" onChange="Checkgjl()">
            <option value="原厂家粉料">原厂家粉料</option>
            <option value="原厂家二次料">原厂家二次料</option>
            <option value="代用料厂家">代用料厂家</option>
          </select>
          <input id="gjltxt" name="gjltxt" type="text" id="gjltxt" size="20" style="Display:none;" value=<%=Rs("gjlqk")%>>
        </p></td>
          </tr>
          <tr>
            <td><span class="STYLE4">调试分值:</span></td>
            <td colspan="5" class="STYLE4">&nbsp; <span id=fzxs class="STYLE4"></span>
              <input type="hidden" name="Cntsfz" value=0 />
              <input type="hidden" name="Cwtsfz" value=0 />
              <input type="hidden" name="Scxfz" value=0 />
              <input type="hidden" name="Zxfz" value=0 />
          </tr>
  </table>
  <table width="98%" border="2" cellpadding="0" bordercolor="#33FFCC">
    <tr>
      <th colspan="7" scope="col"><div align="left" class="STYLE4">质量信息</div></th>
    </tr>
    <tr>
      <td width="86"><span class="STYLE4">检验员:</span></td>
      <td width="90"><span class="STYLE4">&nbsp;</span></td>
      <td colspan="5"><span class="STYLE4">
        <label>产品合格
          <input name="cphg" type="checkbox" disabled id="cphg" value="checkbox" />
          一级评审
          <input name="checkbox2" type="checkbox" disabled value="checkbox" />
          二级评审
          <input name="checkbox3" type="checkbox" disabled value="checkbox" />
          较大问题通知放行
          <input name="checkbox4" type="checkbox" disabled value="checkbox" />
        </label>
        </span></td>
    </tr>
    <tr>
      <td height="100"><span class="STYLE4">备注:</span></td>
      <td height="100" colspan="6"><span class="STYLE4">&nbsp;</span></td>
    </tr>
    <tr>
      <td colspan="3"><span class="STYLE4"></span></td>
      <td width="60"><span class="STYLE4">审核:</span></td>
      <td width="220"><span class="STYLE4"></span>&nbsp;</td>
      <td width="60"><span class="STYLE4">批准:</span></td>
      <td width="220"><span class="STYLE4"></span>&nbsp;</td>
    </tr>
  </table>
  <label><br>
    <input type="submit" name="Submit2" value="·确 定·">
  </label>
</form>
<%End Function%>
<script language=javascript>
//初始化计划结束时间
// document.all.jhjs.value=document.all.Hijhjs.value;
 //初始化分值
 CheckJl();
 //初始化试模料
 switch(document.all.smltxt.value){
   case "原厂家粉料":
   		document.all.smlqk.selectedIndex=0;
   		document.all.smltxt.value="";
   		break;
   case "原厂家二次料":
   		document.all.smlqk.selectedIndex=1;
   		document.all.smltxt.value="";
   		break;
   default:
   		document.all.smlqk.selectedIndex=2;
   		document.all.smltxt.style.display="";
}
 //初始化共挤料
 switch(document.all.gjltxt.value){
   case "原厂家粉料":
   		document.all.gjlqk.selectedIndex=0;
   		document.all.gjltxt.value="";
   		break;
   case "原厂家二次料":
   		document.all.gjlqk.selectedIndex=1;
   		document.all.gjltxt.value="";
   		break;
   default:
   		document.all.gjlqk.selectedIndex=2;
   		document.all.gjltxt.style.display="";
}

function CheckYl()	//根据速度和米重确定额定用料
{
	CheckJl();
	var strdxcs=document.pldadd.tscs.value.split(":");
	strdxcs=strdxcs[1];
	var TmpYl
	if (document.pldadd.mz.value>0 && document.pldadd.qysd.value>0)
	{
		//额定调试用料=米重*速度*60(分钟)*额定调试次数
	TmpYl=document.pldadd.mz.value*document.pldadd.qysd.value*60*strdxcs;
	//大、小双腔对调试用料的影响
	if (document.pldadd.mjlx.value.indexOf("2定型")>=0)
	{
		TmpYl=TmpYl*1.7;
	}
	if (document.pldadd.mjlx.value.indexOf("1定型")>=0)
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

function Checksml() //填写待用料厂家
{
	if (document.all.smlqk.value=="代用料厂家")
 {
  	document.all.smltxt.style.display="";
  }
	else
 {
  document.all.smltxt.style.display="none";
  document.all.smltxt.value="";
  }
}
function Checkgjl() //填写共挤料厂家
{
	if (document.all.gjlqk.value=="代用料厂家")
 {
  	document.all.gjltxt.style.display="";
  }
	else
 {
  document.all.gjltxt.style.display="none";
  document.all.gjltxt.value="";
  }
}

function CheckJl() //模具调试奖励等级计算函数
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
	var Mtcs_inc1 = 0;	//调试次数增量2(模具类型对调试次数的影响)
	var Mtcs_inc2 = 0;	//调试次数增量2(型材寄样对调试次数的影响)

	var Tsfz_inc1 = 0;	//分值增量2(模具类型对厂内分值的影响+)
	var Tsfz_inc2 = 1;	//分值增量1(速度对厂内分值的影响*)
	var Tsfz_inc3 = 0;	//分值增量3(腔室对分值的影响+)
	var Tsfz_inc4 = 0;	//分值增量4(壁厚对分值的影响+)
	var Tsfz_inc5 = 1;	//分值增量4(任务内容对分值的影响*)

	switch (document.pldadd.mjlx.selectedIndex){	//模具类型对调试次数及分值、装箱分值的影响
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
	if(document.all.sd.checked==false){		//速度对调试次数及分值的影响
		Tsfz_inc2=0.84;
	}
	if(document.all.qs.checked==true){		//腔室>5对调试次数及分值的影响
		Tsfz_inc3=0.1;
	}
	if(document.all.bh.checked==true){		//壁厚对调试次数及分值的影响
		Tsfz_inc4=0.1;
	}
	if (document.pldadd.xcjy.checked==true){	//型材寄样对调试次数的影响
		Mtcs_inc2=2;
	}
	var Ftxs=1;//非调系数，用以判断是否需要在厂外调试分值上再加厂内分值的一半。
	var Mtcs=Mtcs_inc1 + Mtcs_inc2;//调试次数变量

	for (key in Dm){
		if (Dm[key]==document.pldadd.cjqy.value + "," + document.pldadd.dmdj.value){//模具等级、断面等级
			var arrTmp = Jl[key].split(",");
			switch (document.pldadd.rwnr.value){//任务类型对调试次数和分值的影响
				case "全套" :
					document.pldadd.tscs.value=Number(JcCs[0]) + Mtcs + ":" + (Number(JcCs[1]) + Mtcs);break;
					case "机头" :
					document.pldadd.tscs.value=Number(JcCs[0]) + Mtcs + ":0";
					arrTmp[2]=Math.round(arrTmp[2]*0.3);
					arrTmp[3]=Math.round(arrTmp[3]*0.3);
					arrTmp[6]=10;		//装箱分值
					Tsfz_inc5=0.3;
					break;
				case "定型" :
					document.pldadd.tscs.value="0:" + (Number(JcCs[1]) + Mtcs);
					arrTmp[2]=Math.round(arrTmp[2]*0.7);
					arrTmp[3]=Math.round(arrTmp[3]*0.7);
					arrTmp[6]=0;		//装箱分值
					Tsfz_inc5=0.7;
					break;
				case "新品" :
					document.pldadd.tscs.value="0:0";
					arrTmp[2]=Math.round(arrTmp[2]*1.7);
					arrTmp[3]=Math.round(arrTmp[3]*1.7);
					arrTmp[4]=Math.round(arrTmp[4]*1.7);
					break;
				case "粗调" :
					document.pldadd.tscs.value=Number(JcCs[0]) + Mtcs_inc1 + ":" + (Number(JcCs[1]) + Mtcs);
					arrTmp[2]=Math.round(arrTmp[2]*0.6);
					arrTmp[3]=Math.round(arrTmp[3]*0.6);
					arrTmp[4]=Math.round(arrTmp[4]*0.6);
			     	arrTmp[6]=Math.round(arrTmp[6]*0.6);	//装箱分值
					break;
				case "非调" :
					document.pldadd.tscs.value="0:0";
					Ftxs=0;
					arrTmp[3]=Math.round(arrTmp[3]*0);
					arrTmp[6]=35;		//装箱分值
					break;
				case "修理" :
					document.pldadd.tscs.value=Number(JcCs[0]) + Mtcs + ":" + (Number(JcCs[1]) + Mtcs);
					arrTmp[2]=Math.round(arrTmp[2]*0.3);
					arrTmp[3]=Math.round(arrTmp[3]*0.3);
					arrTmp[4]=Math.round(arrTmp[4]*0.3);
					break;
				case "TB" :
					document.pldadd.tscs.value="0:0";
					arrTmp[2]=Math.round(arrTmp[2]*0);
					arrTmp[3]=Math.round(arrTmp[3]*0);
					arrTmp[6]=20;		//装箱分值
					break;
				case "TA" :
					document.pldadd.tscs.value=Number(JcCs[0]) + Mtcs + ":" + (Number(JcCs[1]) + Mtcs);
					if (document.pldadd.dmdj.selectedIndex<5){	//装箱分值
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
			fzxs.innerHTML="厂内调试:" + StrCntsfz + ";生产线:" + StrScxfz + ";装箱:" + StrZxfz + ";厂外调试:" + StrCwtsfz;
			document.pldadd.Cntsfz.value=StrCntsfz;
			document.pldadd.Scxfz.value=StrScxfz;
			document.pldadd.Cwtsfz.value=StrCwtsfz;
			document.pldadd.Zxfz.value=StrZxfz;
		}
	}
}
</script>
