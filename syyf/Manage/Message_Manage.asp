<!-- #include file="../Include/conn.asp" -->
<!-- #include file="Include/Chk.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title></title>
<link href="Images/main.css" rel="stylesheet" type="text/css" />
</head>

<body>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="33" background="Images/l_mu_bg2.gif"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="117" height="33" background="Images/l_mu_bg1.gif" class="tt" style="background-repeat:no-repeat">���Թ���</td>
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
	id=Request("id")
	if id="" then
	lname="��������"
	sql="select * from xin_Message order by id desc"
	elseif id=1 then
	lname="δ�ظ�����"
	sql="select * from xin_Message where re_info='�޻ظ�' order by id desc"
	elseif id=2 then
	lname="δ�������"
	sql="select * from xin_Message where audit<>1 order by id desc"
	end if
	set rs=server.CreateObject("ADODB.RECORDSET")
	rs.open sql,conn,1,1
	if not rs.eof then
	pages = 10 '����ÿҳ��ʾ�ļ�¼��
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
        <td width="87%" height="25" align="left" valign="middle" class="t2"><%=lname%></td>
        </tr>
      <tr>
        <td height="25" align="right" valign="middle" class="tb"><table width="100%" border="0" cellpadding="0" cellspacing="1" class="ta">
          <tr>
            <td width="12%" height="25" align="center" valign="middle" class="tb">����</td>
            <td width="15%" align="center" valign="middle" class="tb">�绰</td>
            <td width="20%" align="center" valign="middle" class="tb">����</td>
            <td width="15%" align="center" valign="middle" class="tb">ʱ��</td>
            <td width="18%" align="center" valign="middle" class="tb">״̬</td>
            <td width="20%" height="25" align="center" valign="middle" class="tb">����</td>
          </tr>
          <tr>
            <td height="25" colspan="6" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
             <%
			 do while not rs.eof and pages > 0
			 %>
	          <tr>
                <td width="12%" height="30" align="center" valign="middle"><%=rs("name")%></td>
                <td width="15%" height="30" align="center" valign="middle"><%=rs("tel")%></td>
                <td width="20%" height="30" align="center" valign="middle"><%=rs("mail")%></td>
                <td width="15%" height="30" align="center" valign="middle"><%=rs("addtime")%></td>
                <td width="18%" height="30" align="center" valign="middle">
				<%
				if rs("audit")=0 then
				audit="<font color='red'>δ���</font>"
				else
				audit="�����"
				end if
				if rs("re_info")="�޻ظ�" then
				reinfo="<font color='red'>δ�ظ�</font>"
				else
				reinfo="�ѻظ�"
				end if
				Response.Write(audit&","&reinfo)
				%>				</td>
                <td width="20%" height="30" align="center" valign="middle"><a href="Message_Edit.asp?id=<%=rs("id")%>">��ˡ��ظ�</a> | <a href="Message_Del.asp?id=<%=rs("id")%>" ONCLICK="javascript:return confirm('��ʾ����ȷ��Ҫ����ִ�иò�����')">ɾ��</a></td>
              </tr>
			  <%
			  pages = pages - 1
			  rs.movenext
			  loop
			  else
			  Response.Write("<font color='Red'>�������ݣ�</font>")
			  End if
			  %>
            </table></td>
            </tr>
        </table></td>
        </tr>
      <tr>
        <td height="40" align="right" valign="middle" class="tb"><a href="?page=1&id=<%=Request("id")%>">��ҳ</a>&nbsp; <a href="?page=<%=page-1%>&id=<%=Request("id")%>">��һҳ</a> &nbsp; <a href="?page=<%=page+1%>&id=<%=Request("id")%>">��һҳ</a>&nbsp; <a href="?page=<%=allpages%>&id=<%=Request("id")%>">βҳ</a> &nbsp;
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
