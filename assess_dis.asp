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

Dim strFeedBack,StrTsy
strFeedBack="&rwlx="&StrRwlx&"&lsh="&StrLsh

'定义变量及变量赋值
Dim iyear, imonth, dtstart, dtend, struser, irwzf, iaddfz, ilxrwzf, icount
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
-->
</style>
<!-- #include file="Head.asp" -->
<table width="998" border="0" cellspacing="0" cellpadding="0">
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
            <td width="40%" height="18">　分值统计&gt;&gt; 考评分值 </td>
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
        </form>
        <tr>
          <td height="1" colspan="2" bgcolor="#D6DCE3"></td>
        </tr>
        <tr>
          <td height="1" colspan="2">&nbsp;</td>
        </tr>
      </table>
      <%Call AssDisplay()%></td>
  </tr>
</table>
<!-- #include file="Inc/Foot.asp" -->
<%
Function AssDisplay()
'		%>
<table width="98%" border="2" bordercolor="#33FFCC">
  <caption>
  <span class="STYLE2"><%=struser & " " & dtstart & "至" & dtend & " 考评分值统计"%></span> <br />
  </caption>
  <tr>
    <th width="3%" class="STYLE4"><div align="center">ID</div></th>
    <th class="STYLE4"><div align="center">考评项目</div></th>
    <th class="STYLE4"><div align="center">考评内容</div></th>
    <th class="STYLE4"><div align="center">次数</div></th>
    <th class="STYLE4"><div align="center">应得分</div></th>
  </tr>
  <%Dim i, izf
        i=1 : izf=0
        set rs=server.createobject("adodb.recordset")
		sqltex="select * from kplb where zrr='"&struser&"' and datediff('d',kpsj,'"&dtstart&"')<=0 and datediff('d',kpsj,'"&dtend&"')>=0 Order BY kpnr desc"
		rs.open sqltex,conn,1,1
		do while not rs.eof%>
  <tr>
    <td class="STYLE4"><div align="center"><%=i%></div></td>
    <td class="STYLE4"><div align="center"><%=Rs("kpxm")%></div></td>
    <td class="STYLE4"><div align="center"><%=Rs("kpnr")%></div></td>
    <td class="STYLE4"><div align="center"><%=i%></div></td>
    <td class="STYLE4"><div align="center"><%=Rs("fz")%></div></td>
  </tr>
  <%izf=izf + Rs("fz")
		Rs.movenext
		loop
		Rs.Close%>
  <tr>
    <td class="STYLE4" colspan="4"><div align="right">合计:</div></td>
    <td class="STYLE4"><div align="center"><b><%=izf%></b></div></td>
</table>
<%
End Function
%>
