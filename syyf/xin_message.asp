
<%
'==��������
if Request("action")=1 then
	if Trim(Request("name"))="" or Trim(Request("info"))="" then
		Response.write "<script language='javascript'>alert('��Ϣ��д��������\n�������������ݱ�����д��');history.go(-1);</script>"
		Response.End()
	end if
	
	sql="select * from xin_Message"
	set rs=server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,1,3
	
	rs.addnew
	rs("name")=nohtml(Trim(Request("name")))
	rs("tel")=nohtml(Trim(Request("tel")))
	rs("mail")=nohtml(Trim(Request("mail")))
	rs("info")=nohtml(Trim(Request("info")))
	rs("addtime")=now()
	rs("audit")=0
	rs.update
	
	rs.close
	set rs=nothing
	 Response.Write "<script language='javascript'>alert('���Գɹ�,��ȴ�����Ա�����ظ�!');document.location.href('message.asp');</script>"
	Response.End()
end if



'==��ȡ�����б�

sql="select * from xin_Message where audit=1 order by id desc"
messagelist=""
set rs=server.CreateObject("ADODB.RECORDSET")
rs.open sql,conn,1,1
if not rs.eof then
pages = 5 '����ÿҳ��ʾ�ļ�¼��
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

if rs("re_info")="�޻ظ�" then
	messagelist=messagelist&"<div class=""gbtop""><ul><li>������"&rs("name")&"</li> &nbsp;<li>ʱ�䣺"&rs("addtime")&"</li></ul></div>" & "<div class=""gbinfo"">"&rs("info")&"</div>"&"<div class=""gbc""></div>" 
	
else
	messagelist=messagelist&"<div class=""gbtop""><ul><li>������"&rs("name")&"</li> &nbsp;<li>ʱ�䣺"&rs("addtime")&"</li></ul></div>" & "<div class=""gbinfo"">"&rs("info")&"</div>"&"<div class=""gbhf"">����Ա�ظ���"&rs("re_info")&"</div>" & "<div class=""gbc""></div>"
end if	
pages = pages - 1
rs.MoveNext
loop
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