<%
if Session("Admin")="" then
	Response.Write "<script language='javascript'>alert('��û�е�½���½��ʱ!');parent.location='/manage/login.asp';</script>"
	Response.Write "<font color=Red>��û�е�½,���߳�ʱ,������<a href=login.asp target=_parent>��½</a>"
	Response.End()
end if

function getinfo()
	url=request.ServerVariables("HTTP_HOST") 
	Response.Write("<iframe frameborder='0' width='100%' height='150' scrolling='no' src=''></iframe>")
end function
%>