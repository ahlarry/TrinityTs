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
Dim strins, strspc, strydf
strins=Trim(Request("ins"))
strspc=Trim(Request("spc"))
strydf=Trim(Request("ydf"))
'定义变量及变量赋值
Dim iyear, imonth, dtstart, dtend
iyear = request("selyear")
imonth = request("selmon")
If iyear = "" Then iyear = year(now)
If imonth = "" Then imonth = month(now)
dtstart=cdate(iyear&"年"&imonth&"月1日")
dtend=dateadd("m",1,dtstart)
dtend=dateadd("d",-1,dtend)
'定义各项分值变量
Dim izrr, irwzf, iextrwzf, ispcrwzf, iTotalFz, iSjFz, ikpzf, igxsj
irwzf=0 : iextrwzf=0 : ispcrwzf=0 : ispcrwzf=0 : iTotalFz=0	: iSjFz=0 : ikpzf=100 : igxsj=""
%>
<object ID="WebBrowser" WIDTH="0" HEIGHT="0"
CLASSID="CLSID:8856F961-340A-11D0-A96B-00C04FD705A2"></object>
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
.btn3_mouseover {
	BORDER-RIGHT: #2C59AA 1px solid;
	PADDING-RIGHT: 2px;
	BORDER-TOP: #2C59AA 1px solid;
	PADDING-LEFT: 2px;
	FONT-SIZE: 12px;
FILTER: progid:DXImageTransform.Microsoft.Gradient(GradientType=0, StartColorStr=#ffffff, EndColorStr=#D7E7FA);
	BORDER-LEFT: #2C59AA 1px solid;
	CURSOR: hand;
	COLOR: black;
	PADDING-TOP: 2px;
	BORDER-BOTTOM: #2C59AA 1px solid
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
      <td>厂内分值:
        <select name="ins">
          <option value="0" selected="selected">全部</option>
          <option value="1" <%If strins="1" Then%>selected="selected"<%End If%>>非零</option>
          <option value="2" <%If strins="2" Then%>selected="selected"<%End If%>>为零</option>
        </select></td>
      <td>零星分值:
        <select name="spc">
          <option value="0" selected="selected">全部</option>
          <option value="1" <%If strspc="1" Then%>selected="selected"<%End If%>>非零</option>
          <option value="2" <%If strspc="2" Then%>selected="selected"<%End If%>>为零</option>
        </select></td>
      <td>应得分值:
        <select name="ydf">
          <option value="0" selected="selected">全部</option>
          <option value="1" <%If strexp="1" Then%>selected="selected"<%End If%>>非零</option>
          <option value="2" <%If strexp="2" Then%>selected="selected"<%End If%>>为零</option>
        </select></td>
      <td><input class=btn_2k3 type="submit" value="检 索" /></td>
    </tr>
  </form>
</table>
<%
End Function

Function mtstatDisplay()
	set rs=server.createobject("adodb.recordset")	'厂内任务分值统计
	strSql="select * from [fenz_list] where bz='"&iyear&imonth&"' order by id desc"
	Rs.open strSql,conn,1,1
		If not Rs.eof Then
			igxsj=Rs("gxsj")
		End If
	Rs.Close
	Dim TxtSql
	TxtSql=""
	Select case strins
		case "1"
			TxtSql = " and insfz<>0"
		case "2"
			TxtSql = " and insfz=0"
	End Select
	Select case strydf
		case "1"
			TxtSql = " and ydfz<>0" & TxtSql
		case "2"
			TxtSql = " and ydfz=0" & TxtSql
	End Select
	Select case strspc
		case "1"
			TxtSql = " and spcfz<>0" & TxtSql
		case "2"
			TxtSql = " and spcfz=0" & TxtSql
	End Select
	%>
<table id="fzbhead" width="98%" border="0" cellspacing="0" bordercolor="#33FFCC">
  <caption>
  <span class="STYLE2"><%=iyear & "年" & imonth & "月 调试部厂内分值统计表"%></span>
  </br>

  </caption>
  <tr>
    <td align="left"><input name='preview' type='image' src='Imgv2/prev.png'   id='preview' value='打印预览' alt='打印预览' onClick="doPrint()">
      统计数据:<%=dtstart%> - <%=dtend%></td>
    <td align="right"><button class=btn_2k3 onMouseOver="this.className='btn_2k3'"
onmouseout="this.className='btn3_mouseout'" onClick="RefreshNow()">立即更新</button>
      数据版本:<%=igxsj%> </td>
  </tr>
</table>
<table id="fzbbody" width="98%" border="1" cellspacing="0" bordercolor="#33FFCC">
  <tr>
    <th class="STYLE4" width="3%">ID</th>
    <th class="STYLE4" >责任人</th>
    <th class="STYLE4" >厂内任务</th>
    <th class="STYLE4" >零星任务</th>
    <th class="STYLE4" >应得分</th>
    <th class="STYLE4" >实际分</th>
    <th class="STYLE4" >考评分</th>
    <th class="STYLE4" >操作</th>
  </tr>
  <%
	set rs2=server.createobject("adodb.recordset")
	TmpSql="select * from [fenz_list] where bz='"&iyear&imonth&"'"
	rs2.open TmpSql,conn,1,1
	If Rs2.eof Then
		for i=0 to ubound(c_allzy)
			call intpld_mt(c_allzy(i))
		Next
	End If
	Rs2.Close
		i=1
		set rs2=server.createobject("adodb.recordset")
		TmpSql=TmpSql & TxtSql
		rs2.open TmpSql,conn,1,1
		Do while not Rs2.eof
			izrr=Rs2("zrr")
			irwzf=Rs2("insfz")
			ispcrwzf=Rs2("spcfz")
			iTotalFz=Rs2("ydfz")
			iSjFz=Rs2("sjfz")
			ikpzf=Rs2("kpfz")
	%>
  <tr onmouseover='this.style.backgroundColor="#E6E9F2"' onmouseout='this.style.backgroundColor="#F8FFE8"'>
    <td class="STYLE4"><div align="center"><%=i%></div>
    <td class="STYLE4"><div align="center"><%=izrr%></div></td>
    <td class="STYLE4"><div align="center"><%=irwzf%></div></td>
    <td class="STYLE4"><div align="center"><%=ispcrwzf%></div></td>
    <td class="STYLE4"><div align="center"><%=iTotalFz%></div></td>
    <td class="STYLE4"><div align="center"><span id=span_sjf<%=Rs2("id")%>><%=iSjFz%></span></div></td>
    <td class="STYLE4"><div align="center"><%=ikpzf%></div></td>
    <td class="STYLE4"><div align="center">
        <input type=button id=<%=Rs2("id")%> value="修改" onClick="changesjf(this.id)">
      </div></td>
  </tr>
  <%Rs2.movenext
i=i+1
loop%>
</table>
<%
End Function

function intpld_mt(struser)
	irwzf=0			'厂内任务总分
	ispcrwzf=0		'零星任务总分
	iTotalFz=0		'应得任务总分
	iSjFz=0			'实际任务总分
	ikpzf=100			'考评总分
	set rs=server.createobject("adodb.recordset")	'厂内任务分值统计
	strSql="select * from [ins_tsfz] where zrr='"&struser&"' and datediff('d',sj,'"&dtstart&"')<=0 and datediff('d',sj,'"&dtend&"')>=0 order by sj desc"
	rs.open strSql,conn,1,1
		Do While Not Rs.eof
			irwzf=irwzf+Round(Rs("fz"),1)
			Rs.movenext
		loop
	Rs.close

	set rs=server.createobject("adodb.recordset")	'零星任务分值统计
	strSql="select * from [spc_mission] where zrr='"&struser&"' and datediff('d',wcsj,'"&dtstart&"')<=0 and datediff('d',wcsj,'"&dtend&"')>=0 and rwlx<>'调试' and rwlx<>'修理' order by wcsj desc"
	rs.open strSql,conn,1,1
		Do While Not Rs.eof
			ispcrwzf=ispcrwzf+Rs("fz")
			Rs.movenext
		loop
	Rs.close

	iTotalFz=Round(irwzf + iextrwzf + ispcrwzf,1)	'任务总分统计
	iSjFz=iTotalFz									'实际总分统计

    set rs=server.createobject("adodb.recordset")	'考评分统计
	strSql="select * from kplb where zrr='"&struser&"' and datediff('d',kpsj,'"&dtstart&"')<=0 and datediff('d',kpsj,'"&dtend&"')>=0 Order BY kpnr desc"
	rs.open strSql,conn,1,1
		Do While Not Rs.eof
			ikpzf=ikpzf+Rs("fz")
			Rs.movenext
		loop
	Rs.close

    set rs=server.createobject("adodb.recordset")	'更新统计表
	strSql="select * from fenz_list where zrr='"&struser&"' and bz='"&iyear&imonth&"'"
	rs.open strSql,conn,1,3
		If not Rs.eof Then
			Rs("insfz")=irwzf
			Rs("spcfz")=ispcrwzf
			Rs("ydfz")=iTotalFz
			Rs("sjfz")=iTotalFz
			Rs("kpfz")=ikpzf
			Rs("gxsj")=now()
			Rs.update
		else
			Rs.AddNew
			Rs("zrr")=struser
			Rs("insfz")=irwzf
			Rs("spcfz")=ispcrwzf
			Rs("ydfz")=iTotalFz
			Rs("sjfz")=iTotalFz
			Rs("kpfz")=ikpzf
			Rs("gxsj")=now()
			Rs("bz")=iyear&imonth
			Rs.update
		End If
	Rs.close
end function
%>
<script language="javascript">
function RefreshNow(){ //立即更新统计表中数据
<%	for i=0 to ubound(c_allzy)
		call intpld_mt(c_allzy(i))
	next
%>
window.location.reload();
}
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
