<!-- #include file="../Include/conn.asp" -->
<!-- #include file="Include/Chk.asp" -->
<!-- #include file="Include/fckeditor/fckeditor.asp" -->

<%
	id=request("id")
	if id="" or not isnumeric(id) then
	Response.Write "<script language='javascript'>alert('��������!');document.location.href('News_Manage.asp');</script>"
	Response.End()
	end if
	audit=request("audit")
	SQL="Select * from xin_Message where id="&id
	set rs=server.createobject("adodb.recordset")
	rs.open SQL,conn,1,3
	if rs.eof and rs.bof then
	Response.Write "<script language='javascript'>alert('��������!');document.location.href('Message_Manage.asp');</script>"
	Response.End()
	end if
if Request("Action")=1 then


	rs("re_info")=Trim(Request("re_info"))
	
	if Request.Form("audit")=1 then rs("audit")=audit


	rs.Update
	rs.Close
	Set rs=nothing
	Response.Write "<script language='javascript'>alert('�����ɹ�!');document.location.href('Message_Manage.asp');</script>"
	Response.End()
else
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title></title>
<link href="Images/main.css" rel="stylesheet" type="text/css" />
</head>

<body>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  
  <tr>
    <td><table width="100%" border="0" cellpadding="0" cellspacing="0" class="ta">
      <tr>
        <td width="50%" height="25" align="left" valign="middle" class="t2">��ˡ��ظ�����</td>
        <td width="50%" align="right" valign="middle" class="t2">&nbsp;</td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="8"></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellpadding="0" cellspacing="1" class="ta">
      <tr>
        <td colspan="2" align="left" valign="top" bgcolor="#FFFFFF">
          <table width="100%" border="0" cellspacing="0" cellpadding="0">
		  <form name="form1" method="post" action="?action=1&id=<%=Request("id")%>">
            <tr>
              <td width="20%" height="30" align="right" valign="middle" class="bline"><span class="t3">��&nbsp;&nbsp;&nbsp;&nbsp;����</span></td>
              <td width="80%" height="30" align="left" valign="middle" class="bline"><%=rs("name")%></td>
            </tr>
            <tr class="bline">
              <td width="20%" height="30" align="right" valign="middle" class="bline"><span class="t3">��&nbsp;&nbsp;&nbsp;&nbsp;����</span></td>
              <td width="80%" height="30" align="left" valign="middle" class="bline"><%=rs("tel")%></td>
            </tr>
            <tr class="bline">
              <td width="20%" height="30" align="right" valign="middle" class="bline"><span class="t3">��&nbsp;&nbsp;&nbsp;&nbsp;�䣺</span></td>
              <td width="80%" height="30" align="left" valign="middle" class="bline"><%=rs("mail")%></td>
            </tr>
            <tr>
              <td width="20%" height="30" align="right" valign="middle" class="bline"><span class="t3">����ʱ�䣺</span></td>
              <td width="80%" height="30" align="left" valign="middle" class="bline"><%=rs("addtime")%></td>
            </tr>
            <tr>
             <td width="20%" height="30" align="right" valign="middle" class="bline"><span class="t3">�������ݣ�</span></td>
              <td width="80%" align="left" valign="middle" class="bline"><%=rs("info")%></td>
            </tr>
            <tr>
              <td width="20%" height="30" align="right" valign="middle" class="bline"><span class="t3">��&nbsp;&nbsp;&nbsp;&nbsp;�ˣ�</span></td>
              <td width="80%" align="left" valign="middle" class="bline"><input name="audit" type="checkbox" class="main" id="audit" value="1" <%if rs("audit")=1 then Response.Write("checked")%>>
                <span class="ts">*ѡ��Ϊ���</span></td>
            </tr>
            <tr>
              <td width="20%" height="100" align="right" valign="middle" class="bline"><span class="t3">�ظ����ݣ�</span></td>
              <td width="80%" align="left" valign="middle" class="bline">
                <p>
                  <textarea name="re_info" cols="40" rows="8" id="re_info"><%=rs("re_info")%></textarea>
                  <span class="ts">*�������ó���200��</span></p>                </td>
            </tr>
            <tr>
              <td height="50" align="center" valign="middle">&nbsp;</td>
              <td height="50" align="left" valign="middle"><input name="Submit" type="submit" class="inputbut" value="�ύ">
&nbsp;
<input name="Submit2" type="reset" class="inputbut" value="����" onClick="history.back()"></td>
            </tr>
			</form>
          </table>
                
        </td>
        </tr>
      
    </table></td>
  </tr>
</table>
</body>
</html>
<%
end if
%>