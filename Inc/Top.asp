<%
sql="select * from main"
Set rs_main= Server.CreateObject("ADODB.Recordset")
rs_main.open sql,conn,1,1
contact=ubbcode(rs_main("Contact"))
About=ubbcode(rs_main("About"))
History=ubbcode(rs_main("History"))
Structure=ubbcode(rs_main("Structure"))
Job=ubbcode(rs_main("Job"))
sale=ubbcode(rs_main("sale"))
salea=ubbcode(rs_main("salea"))
service=ubbcode(rs_main("service"))
services=ubbcode(rs_main("services"))
baoyang=ubbcode(rs_main("baoyang"))
%>
<script language="javascript">
<!--
function winopen(url)
{
window.open(url,"search","toolbar=0,location=0,directories=0,status=0,menubar=0,scrollbars=1,resizable=yes,width=640,height=450,top=200,left=100");
}
function MM_openBrWindow(theURL,winName,features) { //v2.0
window.open(theURL,winName,features);
}
//-->
</script>
<script>
function eshop(id) { window.open("Eshop.asp?cpbm="+id,"","height=400,width=640,left=200,top=0,resizable=yes,scrollbars=yes,status=no,toolbar=no,menubar=no,location=no");}
</script> 