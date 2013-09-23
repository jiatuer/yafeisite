<%
function nohtml(str)
	if str<>"" then
		str=Replace(str,"<","&lt;")
		str=Replace(str,">","&gt;")
		nohtml=str
	else
		nohtml=str
	end if
end function

%>