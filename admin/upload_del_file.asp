<%
function del_img(path,filename)
  whichfile=server.mappath(path&filename)
  Set fs = CreateObject("Scripting.FileSystemObject")
  Set thisfile = fs.GetFile(whichfile)
  thisfile.Delete True
end function

'给行id,路径,字段名,表
function del_image(id,path,filename,table)
 set delrs=server.createobject("ADODB.Recordset")
 delrs.open "select "&filename&" from "&table&" where id="&id,conn,1,3
 if delrs(""&filename&"")<>"" then
  del_img path,delrs(""&filename&"") '删除图片文件
 end if
 delrs(""&filename&"")=""  '设置表字段为空
 delrs.update
 set delrs=nothing
end function
%>
