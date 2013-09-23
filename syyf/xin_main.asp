<%
'== 读取左侧主分类
sql="select top 5 * from xin_Product_Info where pid=0 order by ord asc"
productmenu="<ul>"
Set rs=conn.execute(sql)
do while not rs.eof
	productmenu=productmenu&"<li><a href='product.asp?id="&rs("id")&"' target='_blank'><img src='Images/do1.gif' width='5' height='7' border='0' />&nbsp;&nbsp;"&rs("title")&"</a></li>"&vbcrlf
rs.MoveNext
loop
productmenu=productmenu&"</ul>"
rs.Close
Set rs=Nothing

'== 读取文字链接
sql="select * from xin_link order by ord asc"
link="<ul>"
Set rs=conn.execute(sql)
do while not rs.eof

	link=link&"<li><a href='"&rs("url")&"' target='_blank'>"&rs("name")&"</a></li>"&vbcrlf
rs.MoveNext
loop
link=link&"</ul>"
rs.Close
Set rs=Nothing
%>