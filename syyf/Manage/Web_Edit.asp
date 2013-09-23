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
	file.write("'中国新华网站管理系统,生成时间："&now()) & vbcrlf
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
	Response.Write("<script>alert('系统配置更新成功!');</script>")
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
        <td width="50%" height="25" align="left" valign="middle" class="t2">网站基本信息管理</td>
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
              <td width="20%" height="30" align="right" valign="middle" class="t3">公司名称：</td>
              <td width="80%" height="30" align="left" valign="middle"><input name="c_name" type="text" class="main" id="c_name" value="<%=c_name%>" size="50" maxlength="100">
                <span class="ts">*网站名称不得超过100字</span></td>
            </tr>
			<tr>
              <td height="30" align="right" valign="middle" class="t3">关键字：</td>
              <td height="30" align="left" valign="middle"><input name="c_key" type="text" class="main" id="c_key" value="<%=c_key%>" size="50" maxlength="100">
                <span class="ts">                *留空则不显示</span></td>
            </tr>
			<tr>
              <td height="30" align="right" valign="middle" class="t3">网站描述：</td>
              <td height="30" align="left" valign="middle"><input name="c_des" type="text" class="main" id="c_des" value="<%=c_des%>" size="50" maxlength="100">
                <span class="ts">                *留空则不显示</span></td>
            </tr>
            <tr>
              <td height="30" align="right" valign="middle" class="t3">公司地址：</td>
              <td height="30" align="left" valign="middle"><input name="c_adr" type="text" class="main" id="c_adr" value="<%=c_adr%>" size="50" maxlength="100">
                <span class="ts">                *留空则不显示</span></td>
            </tr>
            <tr>
              <td height="30" align="right" valign="middle" class="t3">公司电话：</td>
              <td height="30" align="left" valign="middle"><input name="c_tel" type="text" class="main" id="c_tel" value="<%=c_tel%>" size="50" maxlength="100">
                <span class="ts"> *留空则不显示</span></td>
            </tr>
            <tr>
              <td height="30" align="right" valign="middle" class="t3">公司传真：</td>
              <td height="30" align="left" valign="middle"><input name="c_fax" type="text" class="main" id="c_fax" value="<%=c_fax%>" size="50" maxlength="100">
                <span class="ts"> *留空则不显示</span></td>
            </tr>
            <tr>
              <td height="30" align="right" valign="middle" class="t3">E-mail:</td>
              <td height="30" align="left" valign="middle"><input name="c_mail" type="text" class="main" id="c_mail" value="<%=c_mail%>" size="50" maxlength="100">
                <span class="ts"> *留空则不显示</span></td>
            </tr>
            <tr>
              <td height="30" align="right" valign="middle" class="t3">模板路径：</td>
              <td height="30" align="left" valign="middle"><input name="c_tmp" type="text" class="main" id="c_tmp" value="<%=c_tmp%>" size="50" maxlength="100">
                <span class="ts">                *默认模板路径为：template/ </span></td>
            </tr>
            <tr>
              <td height="30" align="right" valign="middle" class="t3">版权信息：<br>
                <span class="STYLE1">内容框中不要按回车，若换行用&lt;br&gt;，<br>
如果有双引号请改为单引号!<br>
假如操作出错，请改Include/vars.asp文件。手动修改c_copyright 这一行代码。</span></td>
              <td height="30" align="left" valign="middle">&nbsp;
                  <textarea name="c_copyright" cols="90" rows="8" id="c_copyright"><%=c_copyright%></textarea></td>
            </tr>
            <tr>
              <td height="100" align="center" valign="middle">&nbsp;</td>
              <td height="50" align="left" valign="middle"><p>
  <input name="Submit" type="submit" class="inputbut" value="修改">
              </p>
                <p class="ts">*<span class="t1">以上所有字段中请不要包含双引号</span>，否则系统会出现错误，请保证对../include/vars.asp有读写的权限<br>
* 如果调用统计，请换成单引号，以免出错,如&lt;script language='javascript' type='text/javascript' src='http://js.users.51.la/11111.js'&gt;&lt;/script&gt;<br>
                  <span class="ej">*如果本操作执行后出错，请将../include/vars_bak.asp代码重新覆盖到../include/vars.asp</span></p></td>
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
