<!-- #include file="../Include/conn.asp" -->
<!-- #include file="Include/Chk.asp" -->
<!-- #include file="Include/fckeditor/fckeditor.asp" -->

<%

if Request("Action")=1 then
		if Trim(Request("name"))="" or Trim(Request("tid"))="" or Trim(Request("info"))="" then
				Response.write "<script language='javascript'>alert('��Ϣ��д��������\n���ơ���ʾλ�á����ݱ�����д��');history.go(-1);</script>"
		Response.End()
	end if


	
	SQL="Select * from xin_About where name='"&Trim(Request("name"))&"'"
	set rs=server.createobject("adodb.recordset")
	rs.open SQL,conn,1,3
		if not rs.eof then
		Response.Write "<script language='javascript'>alert('�Ѵ����������,����������!');history.back();</script>"
		Response.End()
	end if

	rs.AddNew
	rs("url")=Trim(Request("url"))
	rs("name")=Trim(Request("name"))
	rs("tid")=Trim(Request("tid"))
	rs("ord")=Trim(Request("ord"))
	rs("info")=Trim(Request("info"))

	rs.Update
	rs.Close
	Set rs=nothing
	Response.Write "<script language='javascript'>alert('��ӳɹ�!');document.location.href('About_Manage.asp');</script>"
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
        <td width="50%" height="25" align="left" valign="middle" class="t2">������ҳ</td>
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
		  <form name="form1" method="post" action="?action=1">
            <tr>
              <td width="20%" height="30" align="right" valign="middle" class="bline"><span class="t3">��&nbsp;&nbsp;&nbsp;&nbsp;�ƣ�</span></td>
              <td width="80%" height="30" align="left" valign="middle" class="bline"><input name="name" type="text" class="main" id="name" size="20" maxlength="20"></td>
            </tr>
            <tr>
              <td height="30" align="right" valign="middle" class="bline"><span class="t3">��ʾλ�ã�</span></td>
              <td height="30" align="left" valign="middle" class="bline"><select name="tid" class="main" id="tid">
                <option value="0">��վͷ������</option>
                <option value="1">��վ�ײ�����</option>
              </select>              </td>
            </tr>
            <tr>
              <td height="30" align="right" valign="middle" class="bline"><span class="t3">˳&nbsp;&nbsp;&nbsp;&nbsp;��</span></td>
              <td align="left" valign="middle" class="bline"><input name="ord" type="text" class="main" id="ord" value="0" size="4" maxlength="4">
                <span class="ts">*��ֵԽСԽ��ǰ</span></td>
            </tr>
            <tr>
              <td height="30" align="right" valign="middle" class="bline"><span class="t3">��&nbsp;&nbsp;&nbsp;&nbsp;ַ��</span></td>
              <td align="left" valign="middle" class="bline"><input name="url" type="text" class="main" id="url" size="50" maxlength="200">
                (����Ҫ���ӵ�����ҳ��,��������ַ,��http://��ͷ)</td>
            </tr>
            <tr>
              <td width="20%" height="30" align="right" valign="middle" class="bline"><span class="t3">��&nbsp;&nbsp;&nbsp;&nbsp;�ݣ�</span></td>
              <td width="80%" align="left" valign="middle" class="bline">
                <p>
 <INPUT  type="hidden" name="info" id="Content" style="display:none">
<IFRAME ID="ewebeditor1" src="eWebEditor/ewebeditor.asp?id=Content&style=s_coolblue" frameborder="0" scrolling="no" width=95% height="350"></IFRAME> </p>                </td>
            </tr>
            <tr>
              <td height="50" colspan="2" align="center" valign="middle"><input name="Submit" type="submit" class="inputbut" value="�ύ">
                &nbsp; <input name="Submit2" type="reset" class="inputbut" value="����"></td>
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