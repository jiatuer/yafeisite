<!-- #include file="../Include/conn.asp" -->
<!-- #include file="Include/Chk.asp" -->
<!-- #include file="Include/fckeditor/fckeditor.asp" -->

<%
	id=request("id")
	if id="" or not isnumeric(id) then
	Response.Write "<script language='javascript'>alert('参数错误!');document.location.href('New_Manage.asp');</script>"
	Response.End()
	end if
	audit=request("audit")
	SQL="Select * from xin_Message where id="&id
	set rs=server.createobject("adodb.recordset")
	rs.open SQL,conn,1,3
	if rs.eof and rs.bof then
	Response.Write "<script language='javascript'>alert('参数错误!');document.location.href('Message_Manage.asp');</script>"
	Response.End()
	end if
if Request("Action")=1 then


	rs("re_info")=Trim(Request("re_info"))
	
	if Request.Form("audit")=1 then rs("audit")=audit


	rs.Update
	rs.Close
	Set rs=nothing
	Response.Write "<script language='javascript'>alert('操作成功!');document.location.href('Message_Manage.asp');</script>"
	Response.End()
else
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title></title>
<link href="Images/main.css" rel="stylesheet" type="text/css" />
</head>

<body>
</body>
</html>
<%
end if
%>