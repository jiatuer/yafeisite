<!-- #include file="../Include/conn.asp" -->
<!-- #include file="Include/Chk.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title></title>
<link href="Images/main.css" rel="stylesheet" type="text/css" />
<style type="text/css">
<!--
.STYLE1 {color: #FF0000}
-->
</style>
</head>

<body>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="33" background="Images/l_mu_bg2.gif"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="117" height="33" background="Images/l_mu_bg1.gif" class="tt" style="background-repeat:no-repeat">��ҳ����</td>
        <td width="646" height="33">&nbsp;</td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="10"></td>
  </tr>
  <tr>
    <td>
	<%
	sql="select * from xin_About order by tid asc,ord asc"
	set rs=server.CreateObject("ADODB.RECORDSET")
	rs.open sql,conn,1,1
	if not rs.eof then
	pages = 8 '����ÿҳ��ʾ�ļ�¼��
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
	%>
	<table width="100%" border="0" cellpadding="0" cellspacing="1" class="ta">
      <tr>
        <td width="87%" height="25" align="left" valign="middle" class="t2">��ҳ�б� <span class="STYLE1">˵�������Ҳ���޸ģ������޸���Ӧ�ĵ�ҳ���ݣ�</span></td>
        </tr>
      <tr>
        <td height="25" align="right" valign="middle" class="tb"><table width="100%" border="0" cellpadding="0" cellspacing="1" class="ta">
          <tr>
            <td width="25%" height="25" align="center" valign="middle" class="tb">����</td>
            <td width="25%" align="center" valign="middle" class="tb">��ʾλ��(0:ͷ������ 1:�ײ�����)</td>
            <td width="25%" align="center" valign="middle" class="tb">˳��(��ֵԽСԽ��ǰ)</td>
            <td width="25%" height="25" align="center" valign="middle" class="tb">����</td>
          </tr>
          <tr>
            <td height="25" colspan="4" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
             <%
			 do while not rs.eof and pages > 0
			 %>
	          <tr>
                <td width="25%" height="25" align="center" valign="middle"><%=rs("name")%></td>
                <td width="25%" align="center" valign="middle"><%=rs("tid")%></td>
                <td width="25%" align="center" valign="middle"><%=rs("ord")%></td>
                <td width="25%" height="25" align="center" valign="middle"><a href="About_Edit.asp?id=<%=rs("id")%>">�޸�</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; |&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href="About_Del.asp?id=<%=rs("id")%>" ONCLICK="javascript:return confirm('��ʾ����ȷ��Ҫ����ִ�иò�����')">ɾ��</a></td>
              </tr>
			  <%
			  pages = pages - 1
			  rs.movenext
			  loop
			  else
			  Response.Write("<font color='Red'>���ݿ��������ݣ�</font>")
			  End if
			  %>
            </table></td>
            </tr>
        </table></td>
        </tr>
      <tr>
        <td height="40" align="right" valign="middle" class="tb"><a href="?page=1">��ҳ</a>&nbsp; <a href="?page=<%=page-1%>">��һҳ</a> &nbsp; <a href="?page=<%=page+1%>">��һҳ</a>&nbsp; <a href="?page=<%=allpages%>">βҳ</a> &nbsp;
      ҳ�Σ�<%=page%>/<%=allpages%>ҳ</td>
      </tr>
    </table>
	<%
	rs.close
	set rs=nothing
	%>
	</td>
  </tr>
</table>
</body>
</html>
