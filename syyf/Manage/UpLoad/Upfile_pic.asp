<!-- #include file="../Include/Chk.asp" -->
<head>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	font-size:12px;
	background-color: #ffffff;
}
-->
</style>


<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../Images/main.css" rel="stylesheet" type="text/css">
</head>
<body>
<!--#include FILE="upload.inc"-->
<%
dim upload,file,formName,formPath,iCount,filename,fileExt
set upload=new upload_5xSoft ''建立上传对象
formPath="../../uploadfiles/uploadimages/"
 ''在目录后加(/)
if right(formPath,1)<>"/" then formPath=formPath&"/" 
iCount=0
for each formName in upload.file ''列出所有上传了的文件
 set file=upload.file(formName)  ''生成一个文件对象
 if file.filesize<100 then
 	response.write "请先选择你要上传的图片　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
	response.end
 end if
 	
 if file.filesize>1000000 then
 	response.write "<font size=3><br>图片大小超过了限制　[ <a href=# onclick=history.go(-1)>重新上传</a> ]</font>"
	response.end
 end if

 fileExt=lcase(right(file.filename,4))

 if fileEXT<>".jpg" and fileEXT<>".gif" then
 	response.write "<font size=3><br>文件格式只能为jpg和gif格式　[ <a href=# onclick=history.go(-1)>重新上传</a> ]</font>"
	response.end
 end if 
 randomize
 ranNum=int(90000*rnd)+10000
 filename=formPath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&fileExt
 filename1=year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&fileExt
 
 if file.FileSize>0 then         ''如果 FileSize > 0 说明有文件数据
 file.SaveAs Server.mappath(filename)   ''保存文件
response.write "<script>parent.form1.img.value='"&FileName1&"'</script>"

  iCount=iCount+1
 end if
 set file=nothing
next
set upload=nothing  ''删除此对象

Response.Write "<script>parent.form1.img.value='"&FileName1&"';</script>"

Response.Write("上传成功。")
%>
</body>
</html>