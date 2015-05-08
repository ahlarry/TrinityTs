<script language="JavaScript" type="text/javascript">
//检测任务书提交信息
function CheckM()
{
	var objdm=document.all	
	if (objdm.ddh.value==""){alert("订单号不能为空!\n请输入订单号!"); objdm.ddh.focus();return false;}	
	if (objdm.lsh.value==""){alert("流水号不能为空!\n请输入流水号!"); objdm.lsh.focus(); return false;}	
	if (objdm.khmc.value==""){alert("客户名称不能为空!\n请输入客户名称!"); objdm.khmc.focus(); return false;}
	if (objdm.qysd.value==""){alert("牵引速度不能为空!\n请输入牵引速度!"); objdm.qysd.focus(); return false;}	
	if (objdm.tscs.value==""){alert("额定调试次数不能为空!\n请输入额定调试次数!"); objdm.tscs.focus(); return false;}	
	if (objdm.tsyl.value==""){alert("额定调试用料不能为空!\n请输入额定调试用料!"); objdm.tsyl.focus(); return false;}	
//	if (objdm.jhjs.value==""){alert("计划结束时间不能为空!\n请输入计划结束时间!"); objdm.jhjs.focus(); return false;}	
//	if (objdm.scxh.value==""){alert("调试生产线号不能为空!\n请输入调试生产线号!"); objdm.scxh.focus(); return false;}	
//	if (objdm.mz.value==""){alert("米重不能为空!\n请输入米重!"); objdm.mz.focus(); return false;}		
	if (objdm.smlqk.value=="代用料厂家" && objdm.smltxt.value==""){alert("代用料厂家不能为空!\n请输入代用料厂家!"); objdm.smltxt.focus(); return false;}	
	if (objdm.gjlqk.value=="代用料厂家" && objdm.gjltxt.value==""){alert("代用料厂家不能为空!\n请输入代用料厂家!"); objdm.gjltxt.focus(); return false;}	
}
</script>
