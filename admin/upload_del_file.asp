<%
function del_img(path,filename)
  whichfile=server.mappath(path&filename)
  Set fs = CreateObject("Scripting.FileSystemObject")
  Set thisfile = fs.GetFile(whichfile)
  thisfile.Delete True
end function

'����id,·��,�ֶ���,��
function del_image(id,path,filename,table)
 set delrs=server.createobject("ADODB.Recordset")
 delrs.open "select "&filename&" from "&table&" where id="&id,conn,1,3
 if delrs(""&filename&"")<>"" then
  del_img path,delrs(""&filename&"") 'ɾ��ͼƬ�ļ�
 end if
 delrs(""&filename&"")=""  '���ñ��ֶ�Ϊ��
 delrs.update
 set delrs=nothing
end function
%>
