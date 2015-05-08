<!--#include file="Inc/Sysconn.asp" -->
<!--#include file="inc/Skin_css.asp"-->
<!--#include file="Inc/user_dbinf.asp"-->
<!--#include file="Head.asp" -->
<SCRIPT language=Javascript src="Inc/jsfunc.js"></SCRIPT>
<%
If session("UserName")="" then
	response.redirect "UserServer.asp"
End If
%>
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
          <td><img src="imgV2/spc_add.gif" border="0"></td>
          <td width="79%" valign="middle"><table width="100%" height="1" border="0" cellpadding="0" cellspacing="0" bgcolor="#00AD97">
              <tr>
                <td></td>
              </tr>
            </table></td>
          <td width="5%"><img src="imgV2/more2.gif" width="37" height="24" border="0"></td>
        </tr>
      </table>
      <div align="center">
      <form id="spc_add" name="spc_add" action="spc_indb.asp?action=add" method="post">
        <table width="96%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <th class=th height=25>项目名称</th>
            <th class=th>项目内容</th>
          </tr>
          <tr>
            <td class=rtd width="20%">任务类型</td>
            <td class=ltd><select name="rwlx" onChange="Chanlx(this.value)">
                <option value="装箱">装箱出厂</option>
                <option value="返厂">返厂修理</option>
                <option value="修理">厂外修理</option>
                <option value="调试">厂外调试</option>
              </select>
            </td>
          </tr>
        </table>
        <table id="lxxli" width="96%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td class="rtd" width="20%">修理单号</td>
            <td class="ltd"><input type=text name="xldh" size="15"></td>
          </tr>
          <tr>
            <td class="rtd">任务内容</td>
            <td class="ltd"><textarea name="rwnr" cols="75" rows="4"></textarea></td>
          </tr>
          <tr>
            <td class=rtd>分值</td>
            <td class=ltd><input type=text name="fz" size=8 onKeyPress="javascript:validationNumber(this, 'u_float', 10000, '');">
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
            </td>
          </tr>
          <tr>
            <td class=rtd>责任人</td>
            <td class=ltd><input name="zrr" value=<%=Rs("zrr")%>>
              <font color="#ff0000">&nbsp;&nbsp;可多人，用逗号分隔</font></td>
          </tr>
        </table>
        <table id="lxlsh" width='90%' border="0" cellspacing="0" bordercolor="#33FFCC" style="Display:none;">
          <tr id='lshtr'>
            <td class=rtd>流水号:
              <input name="lsh" type="text" size="9"></td>
            <td align="center">模具类别:
              <input name="mjlb" type="text" size="10"></td>
            <td>任务分值:
              <input name="rwzf" type="text" size="10"></td>
            <td><input value="添加" type="button" onClick="AddLsh()">
              <input value="删除" type="button" onClick="DelLsh()">
            </td>
          </tr>
        </table>
        <table width='90%' border="0" cellspacing="0">
          <tr>
            <td height="1" colspan="4" bgcolor="#D6DCE3"></td>
          </tr>
        </table>
        <table id="lxtsy" width='90%' border="0" cellspacing="0" bordercolor="#33FFCC" style="Display:none;">
          <tr id='tsytr'>
            <td class=rtd>调试人:
              <select name="tsy" id="tsy">
                <%for i = 0 to ubound(c_allzy)%>
                <option value="<%=c_allzy(i)%>"><%=c_allzy(i)%></option>
                <%next%>
              </select>
            </td>
            <td align="center">开始调试:
              <input name="tskssj" type="text" onFocus="CalendarWebControl.show(this,false,this.value);" size="10"></td>
            <td> 结束调试:
              <input name="tsjssj" type="text" onFocus="CalendarWebControl.show(this,false,this.value);" size="10"></td>
            <td><input value="添加" type="button" onClick="AddTsy()">
              <input value="删除" type="button" onClick="DelTsy()">
            </td>
          </tr>
        </table>
        <table id="lxbeiz" width='90%' border="0" cellspacing="0" bordercolor="#33FFCC" style="Display:none;">
          <tr>
            <td class=rtd><span class="STYLE4">备注:</span>
              <input type="hidden" id="LshNum" name="LshNum"/>
              <input type="hidden" id="TsyNum" name="TsyNum"/></td>
            <td><textarea name="beiz" cols="82" rows="4"></textarea></td>
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
</table>
<!-- #include file="Inc/Foot.asp" -->
</BODY>
</HTML><script language="javascript">
document.all.jssj.value=new Date().getFullYear()+"-"+Number(new Date().getMonth()+1)+"-"+new Date().getDate();
//任务类型转换函数
function Chanlx(arg)
{
if (arg=="调试"){
	document.all.lxxli.style.display="none";
	document.all.lxlsh.style.display="";
	document.all.lxtsy.style.display="";
	document.all.lxbeiz.style.display="";
}
else
{
	document.all.lxxli.style.display="";
	document.all.lxlsh.style.display="none";
	document.all.lxtsy.style.display="none";
	document.all.lxbeiz.style.display="none";	
}
}

//增加和删除行
i=1;
function AddLsh(){
   var newTR = lshtr.cloneNode(true);
   newTR.id="a"+(++i)
   lshtr.parentNode.insertAdjacentElement("beforeEnd",newTR);
   ResetLsh();
   document.spc_add.LshNum.value=lxlsh.rows.length;
}

function ResetLsh(){
   var RowCount=lxlsh.rows.length-1
   var ReName=RowCount
   for (var i=0;i<lxlsh.rows[RowCount].cells.length;i++){
      
      var str=lxlsh.rows[RowCount].cells[i].innerHTML
      str=(str.replace(/name=[\d]*/i,"name="+ReName));
      lxlsh.rows[RowCount].cells[i].innerHTML=str 
   }
}

function DelLsh(){
	if (lxlsh.rows.length >1){
		lxlsh.deleteRow(lxlsh.rows.length-1);    
	}
	document.spc_add.LshNum.value=lxlsh.rows.length;
}

var j=1;
function AddTsy(){
   var newTR = tsytr.cloneNode(true);
   newTR.id="a"+(++j)
   tsytr.parentNode.insertAdjacentElement("beforeEnd",newTR);
   ResetTsy();
   document.spc_add.TsyNum.value=lxtsy.rows.length;
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
	document.spc_add.TsyNum.value=lxtsy.rows.length;
}

function View(tb){
  for (var r=0;r<tb.rows.length;r++){
     for (var c=0;c<tb.rows[r].cells.length;c++){
        alert(tb.rows[r].cells[c].innerHTML)
     }
  }
}
</script>
