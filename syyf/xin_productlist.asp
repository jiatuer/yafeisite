<%
'==��ȡ��Ʒ�б�
id=Request("id")
if id="" then
	sql="select * from xin_Product order by id desc"
	menuname="չʾ"
else
	sql="select * from xin_Product where tid="&id&"  or tid in (Select id from xin_Product_Info where pid="&id&")  order by id desc"
	sql2="select * from xin_Product_Info where id="&id
	set rs1=conn.execute(sql2)
	if not rs1.eof then	
		menuname=rs1("title")
	else
		menuname="û�д˷���"
	end if
end if

productlist="<ul>"
set rs=server.CreateObject("ADODB.RECORDSET")
rs.open sql,conn,1,1
if not rs.eof then
pages = 9 '����ÿҳ��ʾ�ļ�¼��
rs.pageSize = pages '����ÿҳ��ʾ�ļ�¼��
allPages = rs.pageCount'����һ���ֶܷ���ҳ
page = Request.QueryString("page")'ͨ����������ݵ�ҳ��
'if������ڻ������Ŵ���
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

'== �����ҳ
pager=	"<a href='?page=1&id="&request("id")&"'>��ҳ</a> &nbsp;"&_
		"<a href='?page="&page-1&"&id="&request("id")&"'>��һҳ</a> &nbsp;"&_
		"<a href='?page="&page+1&"&id="&request("id")&"'>��һҳ</a> &nbsp;"&_
		"<a href='?page="&allpages&"&id="&request("id")&"'>βҳ</a> &nbsp;"&_
		"ҳ�Σ�"&page&"/"&allpages&"ҳ"
%>