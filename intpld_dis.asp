<!--#include file="Inc/Sysconn.asp" -->
<!--#include file="inc/Skin_CSS.asp"-->
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
      <%
      	Dim s_lsh, action
	s_lsh=Trim(Request("s_lsh"))
	action=Trim(Request("action"))
	strSql="select * from [mission] where lsh='"&s_lsh&"'"
		set rs=server.createobject("adodb.recordset")
		rs.open strSql,conn,1,1
		If Rs.eof Or Rs.bof Then
			errmsg="��ˮ�� " & strlsh & " �����鲻����!�������ˮ��!"
			Call WriteErrMsg()
			Response.End
		End If
		Dim strEdcs, strSjcs, strCsxs, strYlxs
		strEdcs=Split(Rs("edtscs"),":")
		strSjcs=Split(Rs("sjtscs"),":")
		strCsxs=(strEdcs(0)-strSjcs(0))*0.02+(strEdcs(1)-strSjcs(1))*0.05	'���Դ���ϵ��
		strYlxs=Fix((Rs("edtsyl")-Rs("sjtsyl"))/(Rs("edtsyl")*0.25))*0.05	'��������ϵ��
      %>
      <table width="98%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="29%" height="18">���ƻ�����&gt;&gt;
            �鿴������</td>
          <td width="71%">&nbsp;</td>
        </tr>
        <tr>
          <td height="1" colspan="2" bgcolor="#D6DCE3"></td>
        </tr>
        <tr>
          <td height="1" colspan="2">&nbsp;</td>
        </tr>
      </table>
      <table width="98%" border="1" cellspacing="0" bordercolor="#33FFCC">
        <caption>
        <span class="STYLE2">����ģ�߳�����������</span> <br />
        </caption>
        <tr>
          <th colspan="6" scope="col"><div align="left" class="STYLE4">ģ����Ϣ</div></th>
        </tr>
        <tr>
          <td width="100"><span class="STYLE4">������:</span></td>
          <td class="STYLE4"><%=Rs("ddh")%>&nbsp;</td>
          <td width="100"><span class="STYLE4">��ˮ��:</span></td>
          <td class="STYLE4"><%=Rs("lsh")%>&nbsp;</td>
          <td width="100"><span class="STYLE4">�ͻ�����:</span></td>
          <td class="STYLE4"><%=Rs("khmc")%>&nbsp;</td>
        </tr>
        <tr>
          <td><span class="STYLE4">�ͻ�����:</span></td>
          <td class="STYLE4"><%=Rs("cjqy")%>&nbsp;</td>
          <td><span class="STYLE4">ǣ���ٶ�:</span></td>
          <td class="STYLE4"><%=Rs("qysd")%>&nbsp;m/min</td>
          <td><span class="STYLE4">����:</span></td>
          <td class="STYLE4"><%=Rs("mz")%>&nbsp;Kg/m</td>
        </tr>
        <tr>
          <td><span class="STYLE4">�����ȼ�:</span></td>
          <td class="STYLE4"><%=Rs("jsdj")%>&nbsp;</td>
          <td><span class="STYLE4">ģ�ߵȼ�:</span></td>
          <td class="STYLE4"><%=Rs("mjdj")%>&nbsp;</td>
          <td><span class="STYLE4">����ȼ�:</span></td>
          <td class="STYLE4"><%=Rs("dmdj")%>&nbsp;</td>
        </tr>
        <tr>
          <td><span class="STYLE4">��������:</span></td>
          <td class="STYLE4"><%=Rs("rwlx")%>&nbsp;</td>
          <td><span class="STYLE4">ģ������:</span></td>
          <td class="STYLE4"><%=Rs("mjlx")%>&nbsp;</td>
          <td><span class="STYLE4">���ں�2.7mm:</span></td>
          <td class="STYLE4"><%If Rs("bh")="1" Then%>
            ��
            <%else%>
            ��
            <%End If%>
            &nbsp;</td>
        </tr>
        <tr>
          <td><span class="STYLE4">�Ͳļ���:</span></td>
          <td class="STYLE4"><%If Rs("jy")="True" Then%>
            ��
            <%else%>
            ��
            <%End If%>
            &nbsp;</td>
          <td><span class="STYLE4">���������:</span></td>
          <td class="STYLE4"><%=Rs("edtsyl")%>&nbsp;Kg</td>
          <td><span class="STYLE4">����Դ���:</span></td>
          <td class="STYLE4"><%=Rs("edtscs")%>&nbsp;</td>
        </tr>
        <tr>
          <td class="STYLE4">�������:</td>
          <td colspan="2" class="STYLE4"><%=Rs("smlqk")%>&nbsp;</td>
          <td class="STYLE4">���������:</td>
          <td colspan="2" class="STYLE4"><%=Rs("gjlqk")%>&nbsp;</td>
        </tr>
        <tr>
          <td><span class="STYLE4">���Է�ֵ:</span></td>
          <td class="STYLE4" colspan="5">���ڵ���:<%=Rs("cntsfz")%>;������:<%=Rs("scxfz")%>;װ��:<%=Rs("bgxfz")%>;�������:<%=Rs("cwtsfz")%>.&nbsp;</td>
        </tr>
      </table>
      <table width="98%" border="1" cellspacing="0" bordercolor="#33FFCC">
        <tr>
          <th colspan="4" scope="col"><div align="left" class="STYLE4">����ϵ��</div></th>
        </tr>
        <tr>
          <td width="100"><span class="STYLE4">�ļ�Ӧ��ϵ��:</span></td>
          <td width="245" class="STYLE4">\&nbsp;</td>
          <td width="145"><span class="STYLE4">ģ������ʱ��:</span></td>
          <td class="STYLE4"><%=Rs("qtsj")%>&nbsp;</td>
        </tr>
        <tr>
          <td><span class="STYLE4">���Խ���ʱ��:</span></td>
          <td class="STYLE4"><%=Rs("sjjssj")%>&nbsp;</td>
          <td><span class="STYLE4">��������ϵ��:</span></td>
          <td class="STYLE4"><%Select case Rs("ccfs")
          		Case "���ڵ���"
          			Response.Write("\")
          		Case "��Ʒ�ϸ�"
          			Response.Write("0")
          		Case "һ��ϸ�"
          			Response.Write("-0.1")
          		Case "����ϸ�"
          			Response.Write("-0.2")
          		Case "����ϸ�"
          			Response.Write("-0.3")
          		Case else
          			Response.Write("ǿ��0.6")
          end select%>
            &nbsp;</td>
        </tr>
        <tr>
          <td><span class="STYLE4">ʵ�ʵ��Դ���:</span></td>
          <td class="STYLE4"><%=Rs("sjtscs")%>&nbsp;</td>
          <td><span class="STYLE4">����Ӧ��ϵ��:</span></td>
          <td class="STYLE4"><%=strCsxs%>&nbsp;</td>
        </tr>
        <tr>
          <td><span class="STYLE4">ʵ�ʵ�������:</span></td>
          <td class="STYLE4"><%=Rs("sjtsyl")%>&nbsp;</td>
          <td><span class="STYLE4">����Ӧ��ϵ��:</span></td>
          <td class="STYLE4"><%=strYlxs%>&nbsp;</td>
        </tr>
      </table>
      <table width="98%" border="1" cellspacing="0" bordercolor="#33FFCC">
        <tr>
          <th colspan="8" scope="col"><div align="left" class="STYLE4">���ڵ���</div></th>
        </tr>
        <tr>
          <td width="30" align="center"><span class="STYLE4">ID</span>
            </th>
          <td align="center"><span class="STYLE4">����Ա</span>
            </th>
          <td align="center"><span class="STYLE4">��������</span>
            </th>
          <td align="center"><span class="STYLE4">������ʽ</span>
            </th>
        </tr>
        <%	Dim m
	TmpSql="select * from [ins_tsxx] where lsh='"&s_lsh&"'"
	set TmpRs=server.createobject("adodb.recordset")
	TmpRs.open TmpSql,Conn,1,1
	If Not(TmpRs.eof Or TmpRs.bof) Then
  		For m=1 to TmpRs.RecordCount%>
        <tr>
          <td><div align="center"><span class="STYLE4"><%=m%></span></div></td>
          <td><div align="center"><span class="STYLE4"><%=TmpRs("tsy")%>&nbsp;</span></div></td>
          <td><div align="center"><span class="STYLE4"><%=TmpRs("tsnr")%>&nbsp;</span></div></td>
          <td><div align="center"><span class="STYLE4"><%=TmpRs("cc")%>&nbsp;</span></div></td>
        </tr>
        <%TmpRs.movenext
  		next
  End If
  TmpRs.Close%>
      </table>
      <table width="98%" border="1" cellspacing="0" bordercolor="#33FFCC">
        <tr>
          <th colspan="8" scope="col"><div align="left" class="STYLE4">�������</div></th>
        </tr>
        <tr>
          <td width="30" align="center"><span class="STYLE4">ID</span>
            </th>
          <td align="center"><span class="STYLE4">����Ա</span>
            </th>
          <td align="center"><span class="STYLE4">����ʱ��</span>
            </th>
          <td align="center"><span class="STYLE4">���շ�ʽ</span>
            </th>
        </tr>
        <%	Dim n,TmpArr,TsyArr()
        n=0
	TmpSql="select * from [ext_tsxx] where lsh like '%"&s_lsh&"%' "
	set TmpRs=server.createobject("adodb.recordset")
	TmpRs.open TmpSql,Conn,1,1
	If Not(TmpRs.eof Or TmpRs.bof) Then
		redim  TsyArr(TmpRs.RecordCount-1)
		If Instr(","&Rs("lsh")&"," , ","&s_lsh&",")>0 Then
  			For m=0 to TmpRs.RecordCount
				TsyArr(m)=TmpRs("tsy")&"||"&TmpRs("kssj")&"/"&TmpRs("jssj")&"||"&TmpRs("cc")
				TmpRs.movenext
			next
		End If
	End If


	If TmpRs.RecordCount>0 Then
		For i=0 to ubound(TsyArr)
			TmpArr=split(TsyArr(i),"||")
  		%>
        <tr>
          <td><div align="center"><span class="STYLE4"><%=i+1%></span></div></td>
          <td><div align="center"><span class="STYLE4"><%=TmpArr(0)%>&nbsp;</span></div></td>
          <td><div align="center"><span class="STYLE4"><%=TmpArr(1)%>&nbsp;</span></div></td>
          <td><div align="center"><span class="STYLE4"><%=TmpArr(2)%>&nbsp;</span></div></td>
        </tr>
        <%next
     End If
     TmpRs.Close%>
      </table>
      <table width="98%" border="1" cellspacing="0" bordercolor="#33FFCC">
        <tr>
          <td height="100"><span class="STYLE4"><strong>��ע:</strong></span></td>
          <td height="100" colspan="6"><span class="STYLE4"><%=Rs("bz")%>&nbsp;</span></td>
        </tr>
        <tr>
          <td colspan="3">&nbsp;</td>
          <td width="60"><span class="STYLE4">���:</span></td>
          <td width="220"><span class="STYLE4"></span>&nbsp;</td>
          <td width="60"><span class="STYLE4">��׼:</span></td>
          <td width="220"><span class="STYLE4"></span>&nbsp;</td>
        </tr>
      </table></td>
  </tr>
</table>
<!-- #include file="Inc/Foot.asp" -->
<map name="Map">
  <area shape="rect" coords="18,14,69,24" href="News.asp">
</map>
</BODY>
</HTML>