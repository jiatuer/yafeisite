<!-- #include file="../Include/conn.asp" -->
<!-- #include file="../Include/vars.asp" -->
<!-- #include file="Include/Chk.asp" -->
<%

if Request.QueryString("action")="1" then
	c_name=Request.Form("c_name")
	c_key=Request.Form("c_key")
	c_des=Request.Form("c_des")
	c_adr=Request.Form("c_adr")
	c_tel=Request.Form("c_tel")
	c_fax=Request.Form("c_fax")
	c_mail=Request.Form("c_mail")
	c_tmp=Request.Form("c_tmp")
	c_copyright=Request.Form("c_copyright")
	
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	path = Server.MapPath("../include/vars.asp")
	set file = fso.OpenTextFile(path,2,TRUE)

	file.write(chr(60)&chr(37)) & vbcrlf
	file.write("'�й��»���վ����ϵͳ,����ʱ�䣺"&now()) & vbcrlf
	file.write("c_name		="""&c_name&"""") & vbcrlf
	file.write("c_key		="""&c_key&"""") & vbcrlf
	file.write("c_des		="""&c_des&"""") & vbcrlf
	file.write("c_adr		="""&c_adr&"""") & vbcrlf
	file.write("c_tel		="""&c_tel&"""") & vbcrlf
	file.write("c_fax		="""&c_fax&"""") & vbcrlf
	file.write("c_mail		="""&c_mail&"""") & vbcrlf
	file.write("c_tmp		="""&c_tmp&"""") & vbcrlf
	file.write("c_copyright	="""&c_copyright&"""") & vbcrlf
	file.write(chr(37)&chr(62)) & vbcrlf  
	
	file.close
	
	set file = nothing
	set fso = nothing
	Response.Write("<script>alert('ϵͳ���ø��³ɹ�!');</script>")
end if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title></title>
<link href="Images/main.css" rel="stylesheet" type="text/css" />
<style type="text/css">
<!--
.STYLE1 {color: #FF3300}
-->
</style>
</head>

<body>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  
  <tr>
    <td><table width="100%" border="0" cellpadding="0" cellspacing="0" class="ta">
      <tr>
        <td width="50%" height="25" align="left" valign="middle" class="t2">��վ������Ϣ����</td>
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
          <table width="90%" border="0" cellspacing="0" cellpadding="0">
		  <form name="form1" method="post" action="?action=1">
            <tr>
              <td width="20%" height="30" align="right" valign="middle" class="t3">��˾���ƣ�</td>
              <td width="80%" height="30" align="left" valign="middle"><input name="c_name" type="text" class="main" id="c_name" value="<%=c_name%>" size="50" maxlength="100">
                <span class="ts">*��վ���Ʋ��ó���100��</span></td>
            </tr>
			<tr>
              <td height="30" align="right" valign="middle" class="t3">�ؼ��֣�</td>
              <td height="30" align="left" valign="middle"><input name="c_key" type="text" class="main" id="c_key" value="<%=c_key%>" size="50" maxlength="100">
                <span class="ts">                *��������ʾ</span></td>
            </tr>
			<tr>
              <td height="30" align="right" valign="middle" class="t3">��վ������</td>
              <td height="30" align="left" valign="middle"><input name="c_des" type="text" class="main" id="c_des" value="<%=c_des%>" size="50" maxlength="100">
                <span class="ts">                *��������ʾ</span></td>
            </tr>
            <tr>
              <td height="30" align="right" valign="middle" class="t3">��˾��ַ��</td>
              <td height="30" align="left" valign="middle"><input name="c_adr" type="text" class="main" id="c_adr" value="<%=c_adr%>" size="50" maxlength="100">
                <span class="ts">                *��������ʾ</span></td>
            </tr>
            <tr>
              <td height="30" align="right" valign="middle" class="t3">��˾�绰��</td>
              <td height="30" align="left" valign="middle"><input name="c_tel" type="text" class="main" id="c_tel" value="<%=c_tel%>" size="50" maxlength="100">
                <span class="ts"> *��������ʾ</span></td>
            </tr>
            <tr>
              <td height="30" align="right" valign="middle" class="t3">��˾���棺</td>
              <td height="30" align="left" valign="middle"><input name="c_fax" type="text" class="main" id="c_fax" value="<%=c_fax%>" size="50" maxlength="100">
                <span class="ts"> *��������ʾ</span></td>
            </tr>
            <tr>
              <td height="30" align="right" valign="middle" class="t3">E-mail:</td>
              <td height="30" align="left" valign="middle"><input name="c_mail" type="text" class="main" id="c_mail" value="<%=c_mail%>" size="50" maxlength="100">
                <span class="ts"> *��������ʾ</span></td>
            </tr>
            <tr>
              <td height="30" align="right" valign="middle" class="t3">ģ��·����</td>
              <td height="30" align="left" valign="middle"><input name="c_tmp" type="text" class="main" id="c_tmp" value="<%=c_tmp%>" size="50" maxlength="100">
                <span class="ts">                *Ĭ��ģ��·��Ϊ��template/ </span></td>
            </tr>
            <tr>
              <td height="30" align="right" valign="middle" class="t3">��Ȩ��Ϣ��<br>
                <span class="STYLE1">���ݿ��в�Ҫ���س�����������&lt;br&gt;��<br>
�����˫�������Ϊ������!<br>
��������������Include/vars.asp�ļ����ֶ��޸�c_copyright ��һ�д��롣</span></td>
              <td height="30" align="left" valign="middle">&nbsp;
                  <textarea name="c_copyright" cols="90" rows="8" id="c_copyright"><%=c_copyright%></textarea></td>
            </tr>
            <tr>
              <td height="100" align="center" valign="middle">&nbsp;</td>
              <td height="50" align="left" valign="middle"><p>
  <input name="Submit" type="submit" class="inputbut" value="�޸�">
              </p>
                <p class="ts">*<span class="t1">���������ֶ����벻Ҫ����˫����</span>������ϵͳ����ִ����뱣֤��../include/vars.asp�ж�д��Ȩ��<br>
* �������ͳ�ƣ��뻻�ɵ����ţ��������,��&lt;script language='javascript' type='text/javascript' src='http://js.users.51.la/11111.js'&gt;&lt;/script&gt;<br>
                  <span class="ej">*���������ִ�к�����뽫../include/vars_bak.asp�������¸��ǵ�../include/vars.asp</span></p></td>
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
