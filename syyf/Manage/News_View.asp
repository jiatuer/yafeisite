<!-- #include file="../Include/conn.asp" -->
<!-- #include file="Include/Chk.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title></title>
<link href="Images/main.css" rel="stylesheet" type="text/css" />
</head>

<body>
<%
	id=request("id")
		if id="" then
	Response.Write "<script language='javascript'>alert('��������!');document.location.href('News_Manage.asp');</script>"
	Response.End()
	end if
	
	sql="select * from xin_New where id="&id
	set rs=server.CreateObject("adodb.recordset")
	rs.open sql,conn,1,1
%>
<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td><table width="100%" border="0" cellpadding="0" cellspacing="1" class="ta">
      <tr>
        <td width="87%" height="25" align="center" valign="middle" class="t2"><%=rs("titles")%></td>
        </tr>
      <tr>
        <td height="25" align="right" valign="middle" class="tb">��Դ��<%=rs("from")%> &nbsp;ʱ�䣺<%=rs("addtime")%>&nbsp; �����<%=rs("hit")%>&nbsp;</td>
        </tr>
      <tr>
        <td height="40" align="left" valign="top" class="tb"><%=rs("info")%></td>
      </tr>
    </table>
	</td>
  </tr>
</table>
</body>
</html>
