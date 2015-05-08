<!--#include file="Inc/Sysconn.asp" -->
<!--#include file="inc/Skin_CSS.asp"-->
<!--#include file="Inc/function.asp"-->
<!--#include file="Inc/user_dbinf.asp"-->
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

Dim strFeedBack,StrRwlx,StrLsh,StrTsy
StrRwlx=Trim(Request("rwlx"))
StrLsh=Trim(Request("Asslsh"))
If StrRwlx="" Then StrRwlx="1" End If
strFeedBack="&rwlx="&StrRwlx&"&lsh="&StrLsh

'定义变量及变量赋值
Dim iyear, imonth, dtstart, dtend, struser, irwzf, iextrwzf, ispcrwzf, icount
icount=1
irwzf=0			'任务总分
iextrwzf=0		'厂外任务总分
ispcrwzf=0		'零星任务总分

iyear = request("selyear")
imonth = request("selmon")
struser = request("seltsy")
If iyear = "" Then iyear = year(now)
If imonth = "" Then imonth = month(now)
If struser = "" Then struser = c_allzy(0)

dtstart=cdate(iyear&"年"&imonth&"月1日")
dtend=dateadd("m",1,dtstart)
dtend=dateadd("d",-1,dtend)
%>
<style type="text/css">
<!--
.STYLE2 {
	font-size: x-large;
	font-weight: bold;
	color: #000066;
}
.warp {
	word-break:break-all;
	width:350px;
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
          <td width="40%" height="18">　分值统计&gt;&gt; 任务分值 </td>
          <td width="71%"><div align="right">请选择:
              <select name="selyear" id="selyear" onchange='location.href("<%=request.servervariables("script_name")%>?selyear="+this.form.selyear.value+"&selmon="+this.form.selmon.value+"&seltsy="+this.form.seltsy.value);'>
                <%for i=year(now)-1 to year(now)+1%>
                <option value="<%=i%>" <%if cint(iyear) = i Then%>selected<%End If%>><%=i%></option>
                <%next%>
              </select>
              年
              <select name="selmon" id="selmon" onchange='location.href("<%=request.servervariables("script_name")%>?selyear="+this.form.selyear.value+"&selmon="+this.form.selmon.value+"&seltsy="+this.form.seltsy.value);'>
                <%for i=1 to 12%>
                <option value="<%=i%>" <%if cint(imonth)=i Then%>selected<%End If%>><%=i%></option>
                <%next%>
              </select>
              月
              &nbsp;&nbsp;
              <select name="seltsy" id="seltsy" onchange='location.href("<%=request.servervariables("script_name")%>?selyear="+this.form.selyear.value+"&selmon="+this.form.selmon.value+"&seltsy="+this.form.seltsy.value);'>
                <%for i=0 to ubound(c_allzy)%>
                <option value="<%=c_allzy(i)%>" <%if struser=c_allzy(i) Then%>selected<%End If%>><%=c_allzy(i)%></option>
                <%next%>
              </select>
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
      <%Call mtstatDisplay()%>
    </td>
  </tr>
  </tbody>

</table>
<!-- #include file="Inc/Foot.asp" -->
<%
Function mtstatDisplay()
'		%>
<table width="98%" border="2" bordercolor="#33FFCC">
  <caption>
  <span class="STYLE2"><%=iyear & "年" & imonth & "月 " &struser& " 任务分值统计"%></span> <br />
  </caption>
  <tr>
    <th width="3%">ID</th>
    <th>流水号</th>
    <th>责任人</th>
    <th>参与次/天数</th>
    <th>完成日期</th>
    <th>任务分值</th>
  </tr>
  <%
			call intpld_mt()	'厂内任务
			call extpld_mt()	'厂外任务
			call spcpld_mt()	'零星任务
			call total_mt()		'总分
		%>
</table>
<%
End Function

function intpld_mt()		'厂内任务分值统计
	icount=1
	set rs=server.createobject("adodb.recordset")
	strSql="select * from [ins_tsfz] where zrr='"&struser&"' and datediff('d',sj,'"&dtstart&"')<=0 and datediff('d',sj,'"&dtend&"')>=0 order by sj desc"
	rs.open strSql,conn,1,1
		Do While Not Rs.eof
		%>
<tr>
  <td align="center"><%=icount%></td>
  <td align="center" class="warp"><a href="intpld_dis.asp?s_lsh=<%=Rs("lsh")%>"><%=Rs("lsh")%></a></td>
  <td align="center"><%=Rs("zrr")%></td>
  <td align="center"><%=Rs("cycs")%></td>
  <td align="center"><%=FormatDateTime(Rs("sj"),2)%></td>
  <td align="center"><%=Rs("fz")%></td>
</tr>
<%
			icount = icount + 1
			irwzf=irwzf+Round(Rs("fz"),1)
			Rs.movenext
		loop
		%>
<tr>
  <td align="right" colspan=5>厂内任务总分:</td>
  <td align="center"><b><%=irwzf%></b></td>
</tr>
<%
	Rs.close
end function

function extpld_mt()		'厂外任务分值统计
	icount=1
	set rs=server.createobject("adodb.recordset")
	strSql="select * from [ext_tsfz] where zrr='"&struser&"' and datediff('d',rksj,'"&dtstart&"')<=0 and datediff('d',rksj,'"&dtend&"')>=0 order by rksj desc"
	rs.open strSql,conn,1,1
		Do While Not Rs.eof
		%>
<tr>
  <td align="center"><%=icount%></td>
  <td align="center" class="warp"><a href="extpld_dis.asp?s_lsh=<%=Rs("lsh")%>"><%=Rs("lsh")%></a></td>
  <td align="center"><%=Rs("zrr")%></td>
  <td align="center">/</td>
  <td align="center"><%=FormatDateTime(Rs("rksj"),2)%></td>
  <td align="center"><%=Rs("fz")%></td>
</tr>
<%
			icount = icount + 1
			iextrwzf=iextrwzf+Rs("fz")
			Rs.movenext
		loop
	Rs.close

	set rs=server.createobject("adodb.recordset")
	strSql="select * from [spc_mission] where zrr='"&struser&"' and rwlx='调试' and datediff('d',rksj,'"&dtstart&"')<=0 and datediff('d',rksj,'"&dtend&"')>=0 order by rksj desc"
	rs.open strSql,conn,1,1
		Do While Not Rs.eof
		%>
<tr>
  <td align="center"><%=icount%></td>
  <td align="center" class="warp" alt="<%=replace(Rs("rwlx"),vbcrlf,"<br>")%>"><%=Rs("xldh")%>&nbsp;</td>
  <td align="center"><%=Rs("zrr")%>&nbsp;</td>
  <td align="center">&nbsp;</td>
  <td align="center"><%=FormatDateTime(Rs("wcsj"),2)%>&nbsp;</td>
  <td align="center"><%=Rs("fz")%>&nbsp;</td>
</tr>
<%
			icount = icount + 1
			iextrwzf=iextrwzf+Rs("fz")
			Rs.movenext
		loop
	Rs.close
		%>
<tr>
  <td align="right" colspan=5>厂外任务总分:</td>
  <td align="center"><b><%=Round(iextrwzf,1)%></b></td>
</tr>
<%
end function

function spcpld_mt()		'零星任务分值统计
	icount=1
	set rs=server.createobject("adodb.recordset")
	strSql="select * from [spc_mission] where zrr='"&struser&"' and rwlx<>'调试' and datediff('d',wcsj,'"&dtstart&"')<=0 and datediff('d',wcsj,'"&dtend&"')>=0 order by wcsj desc"
	rs.open strSql,conn,1,1
		Do While Not Rs.eof
		%>
<tr>
  <td align="center"><%=icount%></td>
  <td align="center" class="warp" alt="<%=replace(Rs("rwlx"),vbcrlf,"<br>")%>"><%=Rs("xldh")%>&nbsp;</td>
  <td align="center"><%=Rs("zrr")%>&nbsp;</td>
  <td align="center">&nbsp;</td>
  <td align="center"><%=FormatDateTime(Rs("wcsj"),2)%>&nbsp;</td>
  <td align="center"><%=Rs("fz")%>&nbsp;</td>
</tr>
<%
			icount = icount + 1
			ispcrwzf=ispcrwzf+Rs("fz")
			Rs.movenext
		loop
		%>
<tr>
  <td align="right" colspan=5>零星任务总分:</td>
  <td align="center"><b><%=ispcrwzf%></b></td>
</tr>
<%
	Rs.close
end function

function total_mt()		'总分统计
	Dim iTotalFz
		iTotalFz=Round(irwzf + iextrwzf + ispcrwzf,1)
%>
<tr>
  <td align="right" colspan=5>总分:</td>
  <td align="center"><b><%=iTotalFz%></b></td>
</tr>
<%
end function
%>
