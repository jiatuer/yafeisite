<%
'== 读取新闻分类
sql="select * from xin_New_Pros where pid=0 order by ord asc"
newsmenu="<ul>"
Set rs=conn.execute(sql)
do while not rs.eof
	newsmenu=newsmenu&"<li><a href='news.asp?id="&rs("id")&"'><img src='Images/icon5x.png' width='11' height='11' border='0' />&nbsp;&nbsp;<b>"&rs("name")&"</b></a></li>"
	sql="select * from xin_New_Pros where pid="&rs("id")&" order by ord asc"
	set rs2=conn.execute(sql)
	do while not rs2.eof
		newsmenu=newsmenu&"<li><a href='news.asp?id="&rs2("id")&"'>&nbsp;&nbsp;&nbsp;&nbsp;"&rs2("name")&"</a></li>"
		rs2.MoveNext
	loop
	rs2.Close
	Set rs2=Nothing
rs.MoveNext
loop
newsmenu=newsmenu&"</ul>"
rs.Close
Set rs=Nothing
%>