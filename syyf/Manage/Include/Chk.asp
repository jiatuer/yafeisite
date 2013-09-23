<%
if Session("Admin")="" then
	Response.Write "<script language='javascript'>alert('您没有登陆或登陆超时!');parent.location='/manage/login.asp';</script>"
	Response.Write "<font color=Red>您没有登陆,或者超时,请重新<a href=login.asp target=_parent>登陆</a>"
	Response.End()
end if

function getinfo()
	url=request.ServerVariables("HTTP_HOST") 
	Response.Write("<iframe frameborder='0' width='100%' height='150' scrolling='no' src=''></iframe>")
end function
%>