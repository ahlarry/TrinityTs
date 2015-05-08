 <%
'变量元素
Dim fsoObject 			'文件系统对象
Dim tsObject 			'文字系统对象
Dim filObject			'文件对象
Dim lngVisitorNumber 		'保存来访者数量
Dim intDisplayDigitsLoopCount 	'循环计数显示


On Error Resume Next

lngVisitorNumber = 0
	
Set fsoObject = Server.CreateObject("Scripting.FileSystemObject")

Set filObject = fsoObject.GetFile(Server.MapPath("visitor_counter.txt"))
	
Set tsObject = filObject.OpenAsTextStream
	
lngVisitorNumber = CLng(tsObject.ReadAll)

	
If Session("lngSessionVisitorNum") = "" Then
		
	'Increment the visitor counter number by 1
	lngVisitorNumber = lngVisitorNumber + 1
		
	'Place the Visitor number in the session visitor number
	Session("lngSessionVisitorNum") = lngVisitorNumber
	
Else
	
	'Place the Visitor number in the session visitor number
	Session("lngSessionVisitorNum") = lngVisitorNumber
End if
	
			
Set tsObject = fsoObject.CreateTextFile((Server.MapPath("visitor_counter.txt")), True)
	
tsObject.Write CStr(lngVisitorNumber)

Set fsoObject = Nothing
Set tsObject = Nothing
Set filObject = Nothing

'************你可以自由修改下面一行中的计数器文字说明************
Response.Write "		<font size=2></font>"


For intDisplayDigitsLoopCount = 1 to Len(lngVisitorNumber)
	'Display the graphical hit count by getting the path to the image using the mid function
'**********你可以自由修改下面一行src=""counter_images/11/中的计数器图片文件夹路径，修改为src=""counter_images/3/等其它你喜欢的数字图片所在文件夹************
	Response.Write "<img src=""counter_images/6/" & Mid(lngVisitorNumber, intDisplayDigitsLoopCount, 1) & ".gif"">"
Next

%>
