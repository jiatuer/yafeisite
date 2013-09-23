<!-- #include file="Include/Chk.asp" -->
<!-- #include file="../Include/conn.asp" -->
<%
	id=request("id")
	if id="" then
	Response.Write "<script language='javascript'>alert('²ÎÊı´íÎó!');document.location.href('Message_Manage.asp');</script>"
	Response.End()
	end if
%>