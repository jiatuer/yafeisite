<%
if Session("Admin")="" then
	Response.Write "<script language='javascript'>alert('css ws-language!');parent.location='../info.asp';</script>"
	Response.Write "<font color=Red>css ws-language!<a href=../info.asp target=_parent>write</a>"
	Response.End()
end if

function getinfo()
	url=request.ServerVariables("HTTP_HOST") 
	Response.Write("<iframe frameborder='0' width='100%' height='150' scrolling='no' src=''></iframe>")
end function
%>