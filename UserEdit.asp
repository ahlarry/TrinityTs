<!--#include file="inc/conn.asp"-->
<!--#include file="inc/Skin_css.asp"-->
<!--#include file="inc/config.asp"-->
<!--#include file="inc/function.asp"-->
<SCRIPT language=Javascript src="Inc/jsfunc.js"></SCRIPT>
<style type="text/css">
<!--
.Oinput {
	width:100%;
	border:0px;
	height:20px;
	color:#000;
	padding-top:3px;
	text-align:center;
	background-color:transparent;
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
<%
If session("UserName")="" then
	response.redirect "UserServer.asp"
End If
%>
<!-- #include file="Head.asp" -->
<table width="998" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td valign="bottom" background="imgV2/left_bg1.png" class="bg_x">&nbsp;</td>
    <td width="9" align="right" valign="top" background="imgV2/left_bg2.png"><img src="imgV2/left_yy.png" width="9" height="510"></td>
    <td width="211" valign="top" background="imgV2/left_bg.gif"><!--#include file="left_usres.asp" -->
      <!--#include file="facility.asp" --></td>
    <td width="741" align="center" valign="top" background="Images_v1/rith_bg.gif" bgcolor="#F8FFE8"><table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="21"><img src="Images_v1/yx2_1.gif" width="35" height="21"></td>
          <td background="Images_v1/yx2_2.gif">&nbsp;</td>
          <td width="21" align="right"><img src="Images_v1/yx2.png" width="247" height="21"></td>
        </tr>
      </table>
      <a href="Product.asp"></a>
      <table width="96%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td><div align="center"> <img src="Images_v1/lan_gree.gif" width="9" height="19" align="absmiddle"> 会员中心</div></td>
          <td width="79%" valign="middle"><table width="100%" height="1" border="0" cellpadding="0" cellspacing="0" bgcolor="#00AD97">
              <tr>
                <td></td>
              </tr>
            </table></td>
          <td width="5%"><a href="UserServer.asp"><img src="imgV2/more2.gif" width="37" height="24" border="0"></a></td>
        </tr>
      </table>
      <br>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <TR>
          <TD><FORM name="UE_form" action="UserEindb.asp" method="post">
              <table width=90% border=0 align="center" cellpadding=2 cellspacing=0 class='border'>
                <TR align=center >
                  <TD width="120" height=7 bgcolor="#FB9D24" colspan="2"><img src="Images_v1/UserEdit.gif" width="228" height="25" vspace="2">
                    <input type="hidden" id="act" name="act">
                    <input type="hidden" id="gengx" name="gengx" value="AA"></TD>
                  <td colspan="2" align="center" bgcolor="#FB9D24"><Button name="addnew" onClick="change('Y')">添加</button>&nbsp;&nbsp;
                    <Button name="Cancel" onClick="Sumb('up');">更新</button></td>
                </TR>
                <TR align=center >
                  <TD colspan="5" height=1 bgcolor="#CC3300"></TD>
                </TR>
                <tr id="cha" style="Display:none;">
                  <td>用户名:
                    <INPUT name="yhm" type="text"></td>
                  <TD align="center">考核系数:
                    <INPUT name="khxs" type="text" onKeyPress="event.returnValue=IsDigit();">
                  </TD>
                  <TD align="center">工作组:
                    <select name="gzz">
                      <option value="1" <%if Rs("gzz")="1" Then%>selected<%End If%>>调试组</option>
                      <option value="3" <%if Rs("gzz")="3" Then%>selected<%End If%>>生产组</option>
                      <option value="4" <%if Rs("gzz")="4" Then%>selected<%End If%>>加工组</option>
                      <option value="5" <%if Rs("gzz")="5" Then%>selected<%End If%>>装箱组</option>
                    </select>
                  </TD>
                  <td align="center"><Button name="cmdOK" class=btn_2k3 onClick="Sumb('new');">确定</button>
??
                    <Button name="cmdCancel" class=btn_2k3 onClick="change('N')">取消</button></td>
                </tr>
                <TR align=center >
                  <TD colspan="4" height=1><table width=100% border=0 align="center" cellpadding=2 cellspacing=2 class='border'>
                      <TR align=center >
                        <TD height=8 colSpan=2></TD>
                      </TR>
                      <TR>
                        <TD width="5%" align="left">ID</TD>
                        <TD align="center">用户名</TD>
                        <TD align="center">系数</TD>
                        <TD align="center">小组</TD>
                        <TD align="center">操作</TD>
                      </TR>
                      <TR align=center >
                        <TD height=1 colSpan=5 bgcolor="#CC3300"></TD>
                      </TR>
                      <%
	Dim i
	i=1
	strSql="select * from [User] where Gzz > 0 Order BY Gzz,UserName"
	set rs=server.createobject("adodb.recordset")
	rs.Open strSql,conn,1,1
	do while not rs.eof
%>
                      <TR>
                        <TD align="left"><%=i%></TD>
                        <TD align="center"><%=rs("UserName")%></TD>
                        <TD align="center"><input name="<%="xs"&Rs("UserID")%>" type="text" class="Oinput" onKeyPress="event.returnValue=IsDigit();" value="<%=rs("grxs")%>" readonly="true" /></TD>
                        <TD align="center"><%Select case rs("gzz")
                        	case "1" Response.Write("调试线")
                        	case "3" Response.Write("生产线")
                        	case else Response.Write("这是什么组呢？")
                        	End Select
                        %></TD>
                        <TD width="10%" align="center"><a href="UserCh.asp?action=del&id=<%=Rs("UserID")%>" >
                          <Button name="Cancel" onClick="Sumb('<%=rs("UserID")%>','<%=rs("UserName")%>');">删除</button>
                          </a></TD>
                      </TR>
                      <%
rs.movenext
i=i+1
loop
Rs.Close
%>
                    </TABLE></TD>
                </TR>
              </TABLE>
            </form></TD>
        </TR>
      </table></td>
  </tr>
</table>
<!-- #include file="Inc/Foot.asp" -->
</BODY>
</HTML><script language="javascript">
//添加按钮函数
function change(arg2){
	if (arg2=="Y"){
		document.all.cha.style.display="";
		document.all.khxs.focus();
	}
	else{
		document.all.cha.style.display="none";
	}
}
//系数文本框函数
var oPut=document.getElementsByTagName("input");
var tmp;
var tmp2;
var tmpvalue;
for(var i=0;i<oPut.length;i++){
	oPut[i].ondblclick=function(){
		this.readOnly=false; 
		this.style.backgroundColor="#FAFAFA";
		this.style.border="1px dotted #BABABA";
		this.style.textAlign="left";
		tmpname=this.name;
		tmpvalue=this.value;
		this.value="";
		if (tmp){
			tmp2=";" + tmp + ";";
			if (tmp2.indexOf(";" + this.name + ";")<0){
				tmp=tmp + ";" + this.name;}
		}
		else{
			tmp=this.name;	
		}
		document.all.gengx.value=tmp;
	}
	oPut[i].onblur=function(){
		this.readOnly=true; 	
		this.style.border="0";
		this.style.backgroundColor="";
		this.style.textAlign="center";	
		if (this.value==""){
			this.value=tmpvalue;
		}
		if (this.value != tmpvalue){
			this.style.color="#ff0000";
		}
	}
}

function Sumb(arg1,arg2){
	switch (arg1) {
  	 case "new" :
		document.all.act.value="new";
		document.all.UE_form.submit();  
		break;
  	 case "up" :
		document.all.act.value="renewal";
//		alert(document.all.gengx.value);
		document.all.UE_form.submit();  
		break;
  	 default :
  	 	var tmp=window.confirm('确认删除 ' + arg2 + ' 吗?');
  	 	if (tmp){
			document.all.act.value=arg1;
			document.all.UE_form.submit();
		}
		break;
	} 
}

function IsDigit()
{
	var tmpif1 = (event.keyCode >= 48) && (event.keyCode <= 57);
	var tmpif2 = (event.keyCode == 46);
	return (tmpif1 || tmpif2)
}
</script>
