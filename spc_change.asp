<!--#include file="Inc/Sysconn.asp" -->
<!--#include file="inc/Skin_CSS.asp"-->
<!--#include file="Inc/user_dbinf.asp"-->
<SCRIPT language=Javascript src="Inc/jsfunc.js"></SCRIPT>
<%
If session("UserName")="" then
	response.redirect "UserServer.asp"
End If

Dim Strid,Strxldh,ArrFz,TmpFz
Strid=Trim(Request("s_id"))
TmpFz=0
'Strxldh=Trim(Request("s_xldh"))
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
          <td><img src="imgV2/spc_cha.gif" border="0"></td>
          <td width="79%" valign="middle"><table width="100%" height="1" border="0" cellpadding="0" cellspacing="0" bgcolor="#00AD97">
              <tr>
                <td></td>
              </tr>
            </table></td>
          <td width="5%"><img src="imgV2/more2.gif" width="37" height="24" border="0"></td>
        </tr>
      </table>
      <!--       <Table border=0 cellpadding=2 cellspacing=0 width="98%">
        <Form name=frm_searchlsh action="<%=Request.Servervariables("SCRIPT_NAME")%>" method=post onsubmit='return searchlsh_true();'>
          <tr>
            <td>&nbsp;&nbsp;输入修理单号:
              <input tabindex=1 type=text name=s_xldh size=15 value="<%=Trim(Request("s_xldh"))%>">
              <input type="submit" value=" 查 找 ">
            </td>
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
 -->
      <%Call spc_change(Strid)%>
    </td>
  </tr>
</table>
</table>
<!-- #include file="Inc/Foot.asp" -->
<map name="Map">
  <area shape="rect" coords="18,14,69,24" href="News.asp">
</map>
</BODY>
</HTML><%Function spc_change(x)
	If x="" Then Response.Write("请进入任务列表选择要更改的零星任务!<br>") : Exit Function
	strSql="select * from [spc_mission] where ID=" & Clng(x)
	set Rs=server.createobject("adodb.recordset")
	Rs.open strSql,Conn,1,3
	If Rs.Eof Or Rs.Bof Then
		errmsg="修理单号为 【 " & x & " 】 的零星任务不存在!"
		Call WriteErrMsg()
	Else
		If Rs("rwlx")<>"调试" Then
			Call xchange(Rs)
		Else
			Call Tschange(Rs)
		End If 
	End If
	Rs.Close
End Function

Function xchange(Rs)
Dim TmpDH,TmpRwlx,TmpRwnr,TmpSj,TmpFz,TmpZrr,TmpSql,TmpRs
TmpDH=Rs("xldh") : TmpRwlx=Rs("rwlx") : TmpRwnr=Rs("rwnr") : TmpSj=Rs("wcsj") : TmpFz=0
TmpSql="select * from spc_mission where rwlx='" &TmpRwlx& "' and xldh='" &TmpDH& "' and datediff('d',wcsj,'"&TmpSj&"')=0"
set TmpRs=server.createobject("adodb.recordset")
TmpRs.open TmpSql,Conn,1,3
Do While Not TmpRs.eof
	TmpFz=TmpFz + TmpRs("fz")
	If TmpZrr="" then
		TmpZrr=TmpRs("zrr")
	else
		TmpZrr=TmpZrr &","& TmpRs("zrr")
	End If 
	TmpRs.movenext
Loop
TmpRs.Close

%>
<form id='spccha' name='spccha'  action=spc_indb.asp?action=change method=post>
  <table width="98%" border="2" bordercolor="#33FFCC">
    <caption>
    <span class="STYLE2">更改单号为<%=TmpDH%>的零星任务</span><br />
    </caption>
    <tr>
      <th class=th height=25>项目名称</th>
      <th class=th>项目内容</th>
    </tr>
    <tr>
      <td class=rtd width="20%">任务类型</td>
      <td class=ltd><input name="rwlx" size="10" readonly="true" value=<%=TmpRwlx%>>
        <input name="id" type="hidden" value=<%=Strid%>></td>
    </tr>
    <tr>
      <td class="rtd">修理单号</td>
      <td class="ltd"><input name="xldh" size="10" value=<%=TmpDH%>>
        <input name="ydh" type="hidden" value=<%=TmpDH%>></td>
    </tr>
    <tr>
      <td class="rtd">任务内容</td>
      <td class="ltd"><textarea name="rwnr" cols="75" rows="4"><%=TmpRwnr%></textarea></td>
    </tr>
    <tr>
      <td class=rtd>分值</td>
      <td class=ltd><input type=text name="fz" size=8 onKeyPress="javascript:validationNumber(this, 'u_float', 100000, '');" value=<%=TmpFz%>>
        分</td>
    </tr>
    <tr>
      <td class=rtd>结束时间</td>
      <td colspan="2" class=ltd><script language=javascript>
  		var myDate=new dateSelector();
  		myDate.year;
 		myDate.inputName='jssj';
  		myDate.display();
		</script>
        <input name="wcsj" type="hidden" value=<%=TmpSj%>>
        <input name="ysj" type="hidden" value=<%=TmpSj%>></td>
    </tr>
    <tr>
      <td class=rtd>责任人</td>
      <td class=ltd><input name="zrr" value=<%=TmpZrr%>>
        <font color="#ff0000">&nbsp;&nbsp;可多人，用逗号分隔</font></td>
    </tr>
    <tr>
      <td height="30" colspan=2 class=ctd>&nbsp;</td>
    </tr>
    <tr>
      <td class=ctd colspan=2><div align="center">
          <input type=submit value=" · 确 定 · ">
        </div></td>
    </tr>
  </table>
</form>
<%End Function

Function Tschange(Rs)
Dim TmpDH,TmpRwlx,TmpRwnr,TmpRwlb,TmpKssj,TmpWcsj,TmpFz,TmpZrr,TmpSql,TmpRs,j,TmpStr
TmpDH=Rs("xldh") : TmpRwlx=Rs("rwlx") : TmpRwnr=Rs("rwnr") : TmpWcsj=Rs("wcsj") : TmpFz=0
TmpSql="select * from spc_mission where rwlx='" &TmpRwlx& "' and xldh='" &TmpDH& "' and datediff('d',wcsj,'"&TmpWcsj&"')=0"
set TmpRs=server.createobject("adodb.recordset")
TmpRs.open TmpSql,Conn,1,3
Do While Not TmpRs.eof
	If TmpZrr="" then
		TmpZrr=TmpRs("zrr")
		TmpKssj=TmpRs("kssj")
		TmpWcsj=TmpRs("wcsj")
	else
		TmpZrr=TmpZrr &","& TmpRs("zrr")
		TmpKssj=TmpKssj &","& TmpRs("kssj")
		TmpWcsj=TmpWcsj &","& TmpRs("wcsj")
	End If 
	TmpFz=TmpFz + TmpRs("fz")
	TmpRs.movenext
Loop
TmpRs.Close
TmpZrr=split(TmpZrr,",")
TmpKssj=split(TmpKssj,",")
TmpWcsj=split(TmpWcsj,",")
TmpFz=Round(TmpFz,0)

TmpStr=split(TmpRwnr,",模具分")
TmpRwlb=Mid(TmpStr(0),6)
TmpStr=Mid(TmpRwnr,Instr(TmpRwnr,"。")+1)
%>
<form id='spccha' name='spccha'  action=spc_indb.asp?action=Tscha method=post>
  <table id="lxlsh" width='90%' border="0" cellspacing="0" bordercolor="#33FFCC">
    <caption>
    <span class="STYLE2">更改单号为<%=TmpDH%>的零星任务</span><br />
    </caption>
    <tr id='lshtr'>
      <td class=rtd width="23%">流水号:
        <input name="lsh" type="text" size="9" value=<%=TmpDH%>>
        <input name="id" type="hidden" value=<%=Strid%>></td>
      <td class=rtd width="23%">任务类型:
        <input name="rwlx" type="text" size="10" value=<%=TmpRwlx%> readonly="true"></td>
      <td class=rtd width="23%">模具类别:
        <input name="mjlb" type="text" size="10" value=<%=TmpRwlb%>></td>
      <td>任务分值:
        <input name="rwzf" type="text" size="10"  value=<%=TmpFz%>></td>
    </tr>
  </table>
  <table width='90%' border="0" cellspacing="0">
    <tr>
      <td height="1" colspan="4" bgcolor="#D6DCE3"></td>
    </tr>
  </table>
  <table id="lxtsy" width='90%' border="0" cellspacing="0" bordercolor="#33FFCC">
    <%For i=0 to ubound(TmpZrr)%>
    <tr id=<%If i=0 Then Response.Write("tsytr") Else Response.Write(i&"tsytr") End If%>>
      <td class=rtd width="23%">调试人:
        <select name=<%If i=0 Then Response.Write("tsy") Else Response.Write(i&"tsy") End If%>>
          <%for j = 0 to ubound(c_allzy)%>
          <option value="<%=c_allzy(j)%>" <%if c_allzy(j)=TmpZrr(i) Then%>Selected<%End If%>><%=c_allzy(j)%></option>
          <%next%>
        </select></td>
      <td class=rtd width="23%">开始调试:
        <input name=<%If i=0 Then Response.Write("tskssj") Else Response.Write(i&"tskssj") End If%> type="text" onFocus="CalendarWebControl.show(this,false,this.value);" size="10" value=<%=TmpKssj(i)%>></td>
      <td class=rtd width="23%">结束调试:
        <input name=<%If i=0 Then Response.Write("tsjssj") Else Response.Write(i&"tsjssj") End If%> type="text" onFocus="CalendarWebControl.show(this,false,this.value);" size="10" value=<%=TmpWcsj(i)%>></td>
      <td align="center"><input value="添加" type="button" onClick="AddTsy()">
        <input value="删除" type="button" onClick="DelTsy()"></td>
    </tr>
    <%Next%>
  </table>
  <table id="lxbeiz" width='90%' border="0" cellspacing="0" bordercolor="#33FFCC">
    <tr>
      <td align="right"><span class="STYLE4">备注:</span>
        <input type="hidden" id="TsyNum" name="TsyNum"/></td>
      <td><textarea name="beiz" cols="80" rows="5"><%=TmpStr%></textarea></td>
    </tr>
  </table>
  <table>
    <tr>
      <td height="30" class=ctd>&nbsp;</td>
    </tr>
    <tr>
      <td><div align="center">
          <input type=submit value=" · 确 定 · ">
        </div></td>
    </tr>
  </table>
</form>
<%End Function%>
<script language=javascript>
//初始化计划结束时间
 document.all.jssj.value=document.all.wcsj.value;
 
 //增加和删除行
var j=1;
function AddTsy(){
   var newTR = tsytr.cloneNode(true);
   newTR.id="a"+(++j)
   tsytr.parentNode.insertAdjacentElement("beforeEnd",newTR);
   ResetTsy();
   document.spccha.TsyNum.value=lxtsy.rows.length;
}

function ResetTsy(){
   var RowCount=lxtsy.rows.length-1
   var ReName=RowCount
   for (var i=0;i<lxtsy.rows[RowCount].cells.length;i++){
      
      var str=lxtsy.rows[RowCount].cells[i].innerHTML
      str=(str.replace(/name=[\d]*/i,"name="+ReName));
      lxtsy.rows[RowCount].cells[i].innerHTML=str 
   }
}

function DelTsy(){
	if (lxtsy.rows.length >1){
		lxtsy.deleteRow(lxtsy.rows.length-1);    
	}
	document.spccha.TsyNum.value=lxtsy.rows.length;
}

function View(){
  for (var r=0;r<lxtsy.rows.length;r++){
     for (var c=0;c<lxtsy.rows[r].cells.length;c++){
        alert(lxtsy.rows[r].cells[c].innerHTML)
     }
  }
}
</script>
