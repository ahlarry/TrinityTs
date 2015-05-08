<!--#include file="Inc/Sysconn.asp" -->
<!--#include file="inc/Skin_css.asp"-->
<!--#include file="Inc/user_dbinf.asp"-->
<style type="text/css">
<!--
.STYLE1 {
	color: #FF6666
}
-->
</style>
<!-- #include file="Head.asp" -->
<%
If session("UserName")="" then
	response.redirect "UserServer.asp"
End If
Dim strzb
strzb=Request("zub")
if strzb="" Then strzb="1" End if
%>
<table width="998" border="0" cellspacing="0" cellpadding="0">
  <tr>
  
  <td valign="bottom" background="imgV2/left_bg1.png" class="bg_x">&nbsp;</td>
  <td width="9" align="right" valign="top" background="imgV2/left_bg2.png"><img src="imgV2/left_yy.png" width="9" height="510"></td>
  <td width="211" valign="top" background="imgV2/left_bg.gif"><!--#include file="login.asp" -->
    <!--#include file="Left_ass.asp" --></td>
  <td width="741" align="center" valign="top" background="Images_v1/rith_bg.gif" bgcolor="#F8FFE8">
  
  <table width="100%" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td width="21"><img src="Images_v1/yx2_1.gif" width="35" height="21"></td>
      <td background="Images_v1/yx2_2.gif">&nbsp;</td>
      <td width="21" align="right"><img src="Images_v1/yx2.png" width="247" height="21"></td>
    </tr>
  </table>
  <table width="96%" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td><img src="imgV2/ass_add.gif" border="0"></td>
      <td width="79%" valign="middle"><table width="100%" height="1" border="0" cellpadding="0" cellspacing="0" bgcolor="#00AD97">
          <tr>
            <td></td>
          </tr>
        </table></td>
      <td width="5%"><img src="imgV2/more2.gif" width="37" height="24" border="0"></td>
    </tr>
  </table>
  <form id="ass_add" name="ass_add" action="ass_indb.asp?action=add" method="post" onSubmit='return checkinf();'>
    <table width="96%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><input name="zub" type="radio" value="1" onClick="location.href='assess_add.asp?zub='+ this.value" <%If strzb="1" Then%> checked="checked" <%End If%> >
          挤出线</td>
        <td><input name="zub" type="radio" value="2" onClick="location.href='assess_add.asp?zub='+ this.value" <%If strzb="2" Then%> checked="checked" <%End If%>>
          修理组</td>
        <td><input name="zub" type="radio" value="3" onClick="location.href='assess_add.asp?zub='+ this.value" <%If strzb="3" Then%> checked="checked" <%End If%>>
          抛光组</td>
        <td><input name="zub" type="radio" value="4" onClick="location.href='assess_add.asp?zub='+ this.value" <%If strzb="4" Then%> checked="checked" <%End If%>>
          加工组</td>
        <td><input name="zub" type="radio" value="5" onClick="location.href='assess_add.asp?zub='+ this.value" <%If strzb="5" Then%> checked="checked" <%End If%>>
          技术组</td>
        <td><input name="zub" type="radio" value="0" onClick="location.href='assess_add.asp?zub='+ this.value" <%If strzb="0" Then%> checked="checked" <%End If%>>
          管理组</td>          
      </tr>
      <tr>
        <td height="1" colspan="6">&nbsp;</td>
      </tr>
      <tr>
        <td height="1" colspan="6" bgcolor="#D6DCE3"></td>
      </tr>
      <tr>
        <td height="1" colspan="6">&nbsp;</td>
      </tr>
    </table>
    <table width="96%" border="0" cellspacing="0" cellpadding="0">
      <tr>
      
      <td class=rtd width="20%">考评项目:</td>
      <td class=ltd >
      
      <select name="kpxm" onChange="ChaKpxm(this.value);ChaKpnr(document.all.kp_sel.value);";>
        <option value="">请选择</option>
        <%	Dim strkpxm
   	Set rs=Server.CreateObject("Adodb.RecordSet")
	strsql="select * from [sys_kpxm] where kpdx="&strzb&" Order BY ID"
	rs.Open strsql,conn,1,1
	do while not rs.eof
		If Rs("kpxm")<>strkpxm Then 
	%>
        <option value="<%=Rs("kpxm")%>"><%=Rs("kpxm")%></option>
        <%End If
strkpxm=Rs("kpxm")        
rs.movenext
loop
Rs.Close
%>
        <option value="其他...">其他...</option>
        </tr>
        
        <tr>
        <td height="5" colspan=2 class=ctd>
        &nbsp;
        </td>
        </tr>
        
        <tr>
        <td class=rtd width="20%">
        考评内容:
        </td>
        
        <td class=ltd>
        
        <select name="kp_sel" onChange="ChaKpnr(this.value)";>        
        <option value="">请选择</option>
      </select>
      <input id="kpnr" name="kpnr" type="text" size="40" style="Display:none;">
      </td>
      
      </tr>
      
      <tr>
        <td height="5" colspan=2 class=ctd>&nbsp;</td>
      </tr>
      <tr>
        <td class=rtd width="20%">考评分:</td>
        <td class=ltd><span class="STYLE1" id=span_kpf>0</span>
          <input id="kpf" name="kpf" type="text" onKeyDown="valNum(event);" onpaste="valClip(event);" />
        </td>
      </tr>
      <tr>
        <td height="5" colspan=2 class=ctd>&nbsp;</td>
      </tr>
      <tr>
        <td class=rtd>责任人:</td>
        <td class=ltd ><select name="zrr"  id="zrr">
            <%for i = 0 to ubound(c_allzy)%>
            <option value="<%=c_allzy(i)%>"><%=c_allzy(i)%></option>
            <%next%>
          </select></td>
      </tr>
      <tr>
        <td height="5" colspan=2 class=ctd>&nbsp;</td>
      </tr>
      <tr>
        <td class="rtd">备注:</td>
        <td class="ltd" ><textarea name="beiz" cols="75" rows="4"></textarea></td>
      </tr>
      <tr>
        <td height="5" colspan=2 class=ctd>&nbsp;</td>
      </tr>
      <tr>
        <td class=rtd>考核时间</td>
        <td colspan="2" class=ltd><script language=javascript>
  		var myDate=new dateSelector();
  		myDate.year;
 		myDate.inputName='kpsj';  
  		myDate.display();
		</script>
        </td>
      </tr>
      <tr>
        <td height="20" colspan=2 class=ctd>&nbsp;</td>
      </tr>
      <tr>
        <td class=ctd colspan=2><div align="center">
            <input type=submit value=" · 确 定 · ">
          </div></td>
      </tr>
    </table>
  </form>
</table>
<!-- #include file="Inc/Foot.asp" -->
</BODY>
</HTML><script language="javascript">
var j;
j=0;
goaler = new Array();
<%set rs=conn.execute("select * from [sys_kpxm] where kpdx="&strzb&" Order BY ID")
if rs.eof then%>
goaler[0] = new Array("","无考评内容","");
<%else
i=0
do while not rs.eof%>
goaler[<%=i%>] = new Array("<%=rs("kpxm")%>","<%=rs("kpnr")%>","<%=rs("dwfz")%>");
<%rs.movenext
i=i+1
loop
end if
rs.close
%>
goaler[<%=i%>] = new Array("其他...","其他","∞");
j=<%=i%>;

function ChaKpxm(locationid)
{
document.all.kp_sel.length = 0; 
var locationid=locationid;
var i;
for (i=0;i <= j; i++)
{
if (goaler[i][0] == locationid)
{ 
document.all.kp_sel.options[document.all.kp_sel.length] = new Option(goaler[i][1], goaler[i][1]);
} 
}
} 

function ChaKpnr(arg){
	var i;
	for (i=0;i<goaler.length;i++){
		if (goaler[i][1] == arg){
			if (arg=="其他"){
				document.all.span_kpf.innerHTML="";
				document.all.kpf.value=0;
				document.all.kpnr.value=arg;
				document.all.kpf.style.display="";
				document.all.kpnr.style.display="";
			}				
			else
			{
				document.all.span_kpf.innerHTML="<strong>" + goaler[i][2] +"</strong>";
				document.all.kpf.value=goaler[i][2];
				document.all.kpnr.value=arg;
				document.all.kpf.style.display="none";
				document.all.kpnr.style.display="none";
			}		
		}
	}
}

function checkinf(){
	var objdm=document.ass_add
	if (objdm.kpxm.value==""){alert("考评项目不能为空!"); objdm.kpxm.focus(); return false;}
	if (objdm.kpnr_sel.value==""){alert("考评内容不能为空!"); objdm.kpnr_sel.focus(); return false;}
	if (objdm.kpsj.value==""){alert("考评时间不能为空!"); objdm.kpsj.focus(); return false;}
	if (objdm.kpf.value==""){alert("考评分不能为空!");return false;}
	return true;
}
</script>
