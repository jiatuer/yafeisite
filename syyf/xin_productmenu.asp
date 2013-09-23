<%
'== 读取产品分类
sql="select * from xin_Product_Info where pid=0 order by ord asc"
productmenu="<ul>"
Set rs=conn.execute(sql)
do while not rs.eof
	productmenu=productmenu&"<li><a href='product.asp?id="&rs("id")&"'><img src='Images/icon5x.png' width='11' height='11' border='0' />&nbsp;&nbsp;<b>"&rs("title")&"</b></a></li>"
	sql="select * from xin_Product_Info where pid="&rs("id")&" order by ord asc"
	set rs2=conn.execute(sql)
	do while not rs2.eof
		productmenu=productmenu&"<li><a href='product.asp?id="&rs2("id")&"'>&nbsp;&nbsp;&nbsp;&nbsp;"&rs2("title")&"</a></li>"
		rs2.MoveNext
	loop
	rs2.Close
	Set rs2=Nothing

rs.MoveNext
loop
productmenu=productmenu&"</ul>"
rs.Close
Set rs=Nothing

%>