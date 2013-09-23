<%
'==读取产品列表
id=Request("id")
if id="" then
	sql="select * from xin_Product order by id desc"
	menuname="展示"
else
	sql="select * from xin_Product where tid="&id&"  or tid in (Select id from xin_Product_Info where pid="&id&")  order by id desc"
	sql2="select * from xin_Product_Info where id="&id
	set rs1=conn.execute(sql2)
	if not rs1.eof then	
		menuname=rs1("title")
	else
		menuname="没有此分类"
	end if
end if

productlist="<ul>"
set rs=server.CreateObject("ADODB.RECORDSET")
rs.open sql,conn,1,1
if not rs.eof then
pages = 9 '定义每页显示的记录数
rs.pageSize = pages '定义每页显示的记录数
allPages = rs.pageCount'计算一共能分多少页
page = Request.QueryString("page")'通过浏览器传递的页数
'if语句属于基本的排错处理
if isEmpty(page) or Cint(page) < 1 then
page = 1
elseif Cint(page) > allPages then
page = allPages 
end if
rs.AbsolutePage = page
do while not rs.eof and pages > 0


	productlist=productlist&"<li><a href='product_view.asp?id="&rs("id")&"'><img src='./UpLoadfiles/upLoadImages/"&rs("img")&"' /></a><span><a href='product_view.asp?id="&rs("id")&"'>"&rs("title")&"</a></span></li>"
	
pages = pages - 1
rs.MoveNext
loop
productlist=productlist&"</ul>"
end if
rs.Close
Set rs=Nothing

'== 构造分页
pager=	"<a href='?page=1&id="&request("id")&"'>首页</a> &nbsp;"&_
		"<a href='?page="&page-1&"&id="&request("id")&"'>上一页</a> &nbsp;"&_
		"<a href='?page="&page+1&"&id="&request("id")&"'>下一页</a> &nbsp;"&_
		"<a href='?page="&allpages&"&id="&request("id")&"'>尾页</a> &nbsp;"&_
		"页次："&page&"/"&allpages&"页"
%>