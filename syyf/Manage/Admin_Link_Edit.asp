<!-- #include file="../Include/conn.asp" -->
<!-- #include file="Include/Chk.asp" -->

<%
	id=request("id")
	if id="" or not isnumeric(id) then
	Response.Write "<script language='javascript'>alert('参数错误!');document.location.href('Link_Manage.asp');</script>"
	Response.End()
	end if
	SQL="Select * from xin_link where id="&Request("id")
	set rs=server.createobject("adodb.recordset")
	rs.open SQL,conn,1,3
	if rs.eof and rs.bof then
	Response.Write "<script language='javascript'>alert('参数错误!');document.location.href('Admin_Manage.asp');</script>"
	Response.End()
	end if
if Request("Action")=1 then

	if Trim(Request("name"))=""  then
				Response.write "<script language='javascript'>alert('请输入站点名称！');history.go(-1);</script>"
		Response.End()
	end if
	if not isnumeric(Trim(Request("ord")))   then
				Response.write "<script language='javascript'>alert('显示顺序必须是数字！');history.go(-1);</script>"
		Response.End()
	end if

	rs("name")=Trim(Request("name"))
	rs("url")=Trim(Request("url"))
	rs("logo")=Trim(Request("logo"))
	rs("info")=Trim(Request("info"))
	rs("ord")=Trim(Request("ord"))

	rs.Update
	rs.Close
	Set rs=nothing
	Response.Write "<script language='javascript'>alert('修改成功!');document.location.href('Link_Manage.asp');</script>"
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