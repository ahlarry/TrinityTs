<!--startprint-->
<!--#include file="Inc/Sysconn.asp" -->
<!--#include file="inc/Skin_CSS.asp"-->
<!--#include file="Inc/function.asp"-->
<!--#include file="Inc/user_dbinf.asp"-->
<!--endprint-->
<%
If session("UserName")="" or session("Gzz")>1 then
	response.redirect "UserServer.asp"
End If
'定义搜索参数
Dim SerLsh
SerLsh=Trim(Request("s_lsh"))
SerLsh=Replace(SerLsh,"，",",")
'定义变量及变量赋值
Dim iyear, imonth, dtstart, dtend
iyear = request("selyear")
imonth = request("selmon")
If iyear = "" Then iyear = year(now)
If imonth = "" Then imonth = month(now)
dtstart=cdate(iyear&"年"&imonth&"月1日")
dtend=dateadd("m",1,dtstart)
dtend=dateadd("d",-1,dtend)
%>
<object ID="WebBrowser" WIDTH="0" HEIGHT="0"
CLASSID="CLSID:8856F961-340A-11D0-A96B-00C04FD705A2"></object>
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
.warp {
	color: #000066;
	word-break:break-all;
	width:250px;
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
      <table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td class="title_left" height="34" align="center" ><font color="#000000">分 值 统 计</font></td>
        </tr>
        <tr>
          <td align="center"><table width="86%" border="0" align="center" cellpadding="0" cellspacing="2">
              <tr>
                <TD width=34 height=24 align=center><IMG src="img/class1.gif" width=20 height=20></TD>
                <TD width="100%"><a href='miss_acc.asp'>任 务 分 值</a></TD>
              </tr>
              <tr>
                <TD width=34 height=24 align=center><IMG src="img/class1.gif" width=20 height=20></TD>
                <TD width="100%"><a href='assess_dis.asp'>考 评 分 值</a></TD>
              </tr>
              <tr>
                <TD width=34 height=24 align=center><IMG src="img/class1.gif" width=20 height=20></TD>
                <TD width="100%"><a href='acc_list.asp'>厂 内 分 值</a></TD>
              </tr>
              <tr>
                <TD width=34 height=24 align=center><IMG src="img/class1.gif" width=20 height=20></TD>
                <TD width="100%"><a href='extacc_list.asp'>厂 外 分 值</a></TD>
              </tr>
              <tr>
                <TD width=34 height=24 align=center><IMG src="img/class1.gif" width=20 height=20></TD>
                <TD width="100%"><a href='lsh_xsxg.asp'>重 置 系 数</a></TD>
              </tr>
            </table></td>
        </tr>
      </table>
      <!--#include file="Left_Index.asp" -->
    </td>
    <td width="741" align="center" valign="top" background="Images_v1/rith_bg.gif" bgcolor="#F8FFE8"><table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="21"><img src="Images_v1/yx2_1.gif" width="35" height="21" /></td>
          <td background="Images_v1/yx2_2.gif">&nbsp;</td>
          <td width="21" align="right"><img src="Images_v1/yx2.png" width="247" height="21" /></td>
        </tr>
      </table>
      <table width="98%" border="0" cellspacing="0" cellpadding="0">
        <form action=<%=request.servervariables("script_name")%> method=get>

        <tr>
          <td width="40%" height="18">　分值统计&gt;&gt; 分值总表 </td>
          <td width="71%"><div align="right">请选择:
              <select name="selyear" id="selyear" onchange='location.href("<%=request.servervariables("script_name")%>?selyear="+this.form.selyear.value+"&selmon="+this.form.selmon.value);'>
                <%for i=year(now)-1 to year(now)+1%>
                <option value="<%=i%>" <%if cint(iyear) = i Then%>selected<%End If%>><%=i%></option>
                <%next%>
              </select>
              年
              <select name="selmon" id="selmon" onchange='location.href("<%=request.servervariables("script_name")%>?selyear="+this.form.selyear.value+"&selmon="+this.form.selmon.value);'>
                <%for i=1 to 12%>
                <option value="<%=i%>" <%if cint(imonth)=i Then%>selected<%End If%>><%=i%></option>
                <%next%>
              </select>
              月 </div></td>
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
      <!--startprint-->
      <%Call mtstatDisplay()%>
      <!--endprint-->
    </td>
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
      <td>&nbsp;&nbsp;输入筛选条件:</td>
      <td>流水号:
        <input type="text" id="s_lsh" name="s_lsh"value=<%=SerLsh%>></td>
      <td><font color="#ff0000">多个流水号以，分隔,中间不能有空格！</font></td>
      <td><input class=btn_2k3 type="submit" value="检 索" /></td>
      <td align="left"><input name='preview' type='image' src='Imgv2/prev.png'   id='preview' value='打印预览' alt='打印预览' onClick="doPrint()">
      </td>
    </tr>
    <tr>
      <td colspan="5"><div align="center"><%=SerLsh%></div></td>
    </tr>
  </form>
</table>
<%
End Function

Function mtstatDisplay()
	%>
<table id="fzbhead" width="98%" border="0" cellspacing="0" bordercolor="#33FFCC">
  <caption>
  <span class="STYLE2"><%=iyear & "年" & imonth & "月 调试部厂外分值统计表"%></span>
  </br>

  </caption>
  <tr>
    <td align="right">统计数据:<%=dtstart%> / <%=dtend%></td>
  </tr>
</table>
<table id="fzbbody" width="98%" border="1" cellspacing="0" bordercolor="#33FFCC">
  <tr>
    <th class="STYLE4"  width="3%">ID</th>
    <th class="STYLE4" >责任人</th>
    <th class="STYLE4" >流水号</th>
    <th class="STYLE4" >任务类型</th>
    <th class="STYLE4" >开始时间</th>
    <th class="STYLE4" >结束时间</th>
    <th class="STYLE4" width="6%">系数</th>
    <th class="STYLE4" >得分</th>
    <th class="STYLE4" >总分</th>
  </tr>
  <%
   	Dim strlsh, strxs, strfz, j, LxStr, TmpStr, TmpArr, strzf, k,strkssj,strjssj,izf
    k=0 : izf=0 : strkssj="" : strjssj=""
  	If SerLsh<>"" Then
  		SerLsh=","&SerLsh&","
  	End If
	for i=0 to ubound(c_allzy)
  		strlsh="" : strxs="" : strfz="" : LxStr="" : TmpStr="" : strzf=0 : TmpArr=""
 		set rs=server.createobject("adodb.recordset")
'*****************************开始零星任务的调试部分
  		strSql="select * from [spc_mission] where datediff('d',rksj,'"&dtstart&"')<=0 and datediff('d',rksj,'"&dtend&"')>=0 and (rwlx='调试' or rwlx='修理') and zrr='"&c_allzy(i)&"' order by xldh"
 	 	rs.open strSql,conn,1,1
  		Do while not Rs.eof
  			If SerLsh<>"" Then
  				If Instr(SerLsh,Rs("xldh"))>0 Then
 					If LxStr="" Then
 	 					LxStr=Rs("xldh")&"$"&Rs("kssj")&"$"&Rs("wcsj")&"$0$"&Rs("fz")&"$厂外"&Rs("rwlx")
 	 				else
 	 					LxStr=LxStr&"||"&Rs("xldh")&"$"&Rs("kssj")&"$"&Rs("wcsj")&"$0$"&Rs("fz")&"$厂外"&Rs("rwlx")
 	 				End If
      				strzf=strzf +  Rs("fz")
      			End If
      		else
      			If LxStr="" Then
 	 				LxStr=Rs("xldh")&"$"&Rs("kssj")&"$"&Rs("wcsj")&"$0$"&Rs("fz")&"$厂外"&Rs("rwlx")
 	 			else
 	 				LxStr=LxStr&"||"&Rs("xldh")&"$"&Rs("kssj")&"$"&Rs("wcsj")&"$0$"&Rs("fz")&"$厂外"&Rs("rwlx")
 	 			End If
      			strzf=strzf +  Rs("fz")
      		End If
      		Rs.movenext
      	Loop
      	Rs.Close
'*****************************结束零星任务的调试部分
  		strSql="select * from [ext_tsfz] where datediff('d',rksj,'"&dtstart&"')<=0 and datediff('d',rksj,'"&dtend&"')>=0 and zrr='"&c_allzy(i)&"' order by lsh"
 	 	rs.open strSql,conn,1,1
  		Do while not Rs.eof
  			If SerLsh<>"" Then
  				If Instr(SerLsh,","&Rs("lsh")&",")>0 Then
          			If strlsh="" Then
          				strlsh=Rs("lsh")
          				strxs=Rs("xs")
          				strfz=Rs("fz")
          				If isNull(Rs("kssj")) Then
          					strkssj="/"
          				else
          					strkssj=Rs("kssj")
          				End If
          				If isNull(Rs("jssj")) Then
          					strjssj="/"
          				else
          					strjssj=Rs("jssj")
          				End If
          			else
              			strlsh=strlsh&"||"&Rs("lsh")
              			strxs=strxs&"||"&Rs("xs")
              			strfz=strfz&"||"&Rs("fz")
              			If isNull(Rs("kssj")) Then
              				strkssj=strkssj&"||/"
              			else
              				strkssj=strkssj&"||"&Rs("kssj")
              			End If
              			If isNull(Rs("jssj")) Then
              				strjssj=strjssj&"||/"
              			else
              				strjssj=strjssj&"||"&Rs("jssj")
              			End If
              		End If
              		strzf=strzf +  Rs("fz")
              	End If
            else
     	  		If strlsh="" Then
  					strlsh=Rs("lsh")
  					strxs=Rs("xs")
  					strfz=Rs("fz")
  					If isNull(Rs("kssj")) Then
  						strkssj="/"
  					else
          				strkssj=Rs("kssj")
          			End If
  					If isNull(Rs("jssj")) Then
  						strjssj="/"
  					else
 	 					strjssj=Rs("jssj")
 	 				End If
  				else
      				strlsh=strlsh&"||"&Rs("lsh")
      				strxs=strxs&"||"&Rs("xs")
      				strfz=strfz&"||"&Rs("fz")
              		If isNull(Rs("kssj")) Then
              			strkssj=strkssj&"||/"
              		else
              			strkssj=strkssj&"||"&Rs("kssj")
              		End If
              		If isNull(Rs("jssj")) Then
              			strjssj=strjssj&"||/"
              		else
              			strjssj=strjssj&"||"&Rs("jssj")
              		End If
      			End If
          		strzf=strzf +  Rs("fz")
			End If
      		Rs.movenext
      	Loop
      	Rs.Close

  		If strlsh<>"" Then
  			strlsh=split(strlsh,"||")
  			strxs=split(strxs,"||")
  			strfz=split(strfz,"||")
  			strkssj=split(strkssj,"||")
  			strjssj=split(strjssj,"||")

   			For j=0 to ubound(strlsh)
 	 				If TmpStr="" Then
 	 					TmpStr=strlsh(j)&"$"&strkssj(j)&"$"&strjssj(j)&"$"&strxs(j)&"$"&strfz(j)&"$厂外调试"
 	 				else
 	 					TmpStr=TmpStr&"||"&strlsh(j)&"$"&strkssj(j)&"$"&strjssj(j)&"$"&strxs(j)&"$"&strfz(j)&"$厂外调试"
 	 				End If
	 		Next
	  	End if

	  	If LxStr<>"" Then
	  		If TmpStr<>"" Then
	  			TmpStr=LxStr&"||"&TmpStr
	  		else
	  			TmpStr=LxStr
	  		End If
	  	End If
	  	If TmpStr<>"" Then
	  		k=k+1
	 		TmpStr=split(TmpStr,"||")
	 		TmpArr=split(TmpStr(0),"$")
			strzf=Round(strzf,2)
	%>
  <tr onmouseover='this.style.backgroundColor="#E6E9F2"' onmouseout='this.style.backgroundColor="#F8FFE8"'>
    <td rowspan=<%=ubound(TmpStr)+1%> class="STYLE4"><div align="center"><%=k%></div>
    <td rowspan=<%=ubound(TmpStr)+1%> class="STYLE4"><div align="center"><%=c_allzy(i)%></div></td>
    <td class="warp"><div align="center"><%=TmpArr(0)%></div></td>
    <td class="STYLE4"><div align="center"><%=TmpArr(5)%></div></td>
    <td class="STYLE4"><div align="center"><%=TmpArr(1)%>&nbsp;</div></td>
    <td class="STYLE4"><div align="center">
        <%If TmpArr(2)<>"" Then Response.Write(TmpArr(2))%>
        &nbsp;</div></td>
    <td class="STYLE4"><div align="center">
        <%If TmpArr(3)<>"" Then Response.Write(TmpArr(3))%>
        </span></div></td>
    <td class="STYLE4"><div align="center"><%=TmpArr(4)%></div></td>
    <td rowspan=<%=ubound(TmpStr)+1%> class="STYLE4"><div align="center"><%=strzf%>&nbsp;</div></td>
  </tr>
  <%

  			for j=1 to ubound(TmpStr)
  				TmpArr=split(TmpStr(j),"$")
  %>
  <tr onmouseover='this.style.backgroundColor="#E6E9F2"' onmouseout='this.style.backgroundColor="#F8FFE8"'>
    <td class="warp"><div align="center"><%=TmpArr(0)%></div></td>
    <td class="STYLE4"><div align="center"><%=TmpArr(5)%></div></td>
    <td class="STYLE4"><div align="center"><%=TmpArr(1)%>&nbsp;</div></td>
    <td class="STYLE4"><div align="center"><%=TmpArr(2)%>&nbsp;</div></td>
    <td class="STYLE4"><div align="center"><%=TmpArr(3)%></span></div></td>
    <td class="STYLE4"><div align="center"><%=TmpArr(4)%></div></td>
  </tr>
  <%			Next
	 	End If
	 	izf=izf+strzf
	Next	%>
  <tr onmouseover='this.style.backgroundColor="#E6E9F2"' onmouseout='this.style.backgroundColor="#F8FFE8"'>
    <td colspan="8" class="STYLE4"><div align="right">合计：</div></td>
    <td class="STYLE4"><div align="center"><strong><%=Round(izf,2)%></strong></div></td>
  </tr>
  <tr onmouseover='this.style.backgroundColor="#E6E9F2"' onmouseout='this.style.backgroundColor="#F8FFE8"'>
    <td colspan="3" class="STYLE4"><div align="right">拟制：</div></td>
    <td class="STYLE4">&nbsp;</td>
    <td class="STYLE4"><div align="right">审核：</div></td>
    <td class="STYLE4">&nbsp;</td>
    <td class="STYLE4"><div align="right">批准：</div></td>
    <td colspan="2" class="STYLE4">&nbsp;</td>
  </tr>
</table>
<%
End Function
%>
<script language="javascript">
function doPrint() { //打印特定部分的内容
bdhtml=window.document.body.innerHTML;
sprnstr="<!--startprint-->";
eprnstr="<!--endprint-->";
prnhtml=bdhtml.substr(bdhtml.indexOf(sprnstr)+17);
prnhtml=prnhtml.substring(0,prnhtml.indexOf(eprnstr));
window.document.body.innerHTML=prnhtml;
window.print();
}
function changesjf(arg){
var strsjf
strsjf=showModalDialog("acc_lcha.asp?id="+arg, "", "dialogWidth:250px; dialogHeight:120px; center:yes; help: no; scroll: no; status:no;");
if (strsjf!=null){
eval("document.all.span_sjf"+ arg +".innerHTML = '"+ strsjf +"'");
}
//window.opener.document.form.minipic.value=smileface;
}
</script>
