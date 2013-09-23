<!-- #include file="../Include/conn.asp" -->
<!-- #include file="Include/Chk.asp" -->

<%
	id=request("id")
	if id="" or not isnumeric(id) then
	Response.Write "<script language='javascript'>alert('参数错误!');document.location.href('Link_Manage.asp');</script>"
	Response.End()
	end if
	SQL="Select * from xin_link where id="&Request("id")
	set rs=server.createobject("adodb.recordset")
	rs.open SQL,conn,1,3
	if rs.eof and rs.bof then
	Response.Write "<script language='javascript'>alert('参数错误!');document.location.href('Admin_Manage.asp');</script>"
	Response.End()
	end if
if Request("Action")=1 then

	if Trim(Request("name"))=""  then
				Response.write "<script language='javascript'>alert('请输入站点名称！');history.go(-1);</script>"
		Response.End()
	end if
	if not isnumeric(Trim(Request("ord")))   then
				Response.write "<script language='javascript'>alert('显示顺序必须是数字！');history.go(-1);</script>"
		Response.End()
	end if

	rs("name")=Trim(Request("name"))
	rs("url")=Trim(Request("url"))
	rs("logo")=Trim(Request("logo"))
	rs("info")=Trim(Request("info"))
	rs("ord")=Trim(Request("ord"))

	rs.Update
	rs.Close
	Set rs=nothing
	Response.Write "<script language='javascript'>alert('修改成功!');document.location.href('Link_Manage.asp');</script>"
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
        <td width="50%" height="25" align="left" valign="middle" class="t2">添加友情链接</td>
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
          <table width="70%" border="0" cellspacing="0" cellpadding="0">
		  <form name="form1" method="post" action="?action=1&id=<%=Request("id")%>">
            <tr>
              <td width="30%" height="30" align="right" valign="middle" class="t3">站点名称：</td>
              <td width="70%" height="30" align="left" valign="middle"><input name="name" type="text" class="main" id="name" value="<%=rs("name")%>" size="30" maxlength="30"></td>
            </tr>
            <tr>
              <td width="30%" height="30" align="right" valign="middle" class="t3">站点网址：</td>
              <td width="70%" height="30" align="left" valign="middle"><input name="url" type="text" class="main" id="url" value="<%=rs("url")%>" size="30" maxlength="100">
                <span class="ts">*请以http://开头</span></td>
            </tr>
            <tr>
              <td width="30%" height="30" align="right" valign="middle" class="t3">站点Logo：</td>
              <td width="70%" height="30" align="left" valign="middle"><input name="logo" type="text" class="main" id="logo" value="<%=rs("logo")%>" size="30" maxlength="100">
                <span class="ts">*请以http://开头</span></td>
            </tr>
            <tr>
              <td width="30%" height="30" align="right" valign="middle" class="t3">显示顺序：</td>
              <td width="70%" height="30" align="left" valign="middle"><input name="ord" type="text" class="main" id="ord" value="<%=rs("ord")%>" size="4" maxlength="4">
                <span class="ts">*数值越小越靠前</span></td>
            </tr>
            <tr>
              <td width="30%" height="50" align="right" valign="middle" class="t3">站点说明：</td>
              <td width="70%" height="30" align="left" valign="middle"><textarea name="info" cols="30" rows="3" id="info"><%=rs("info")%></textarea></td>
            </tr>
            <tr>
              <td height="50" colspan="2" align="center" valign="middle"><input name="Submit" type="submit" class="inputbut" value="修改">
                &nbsp; <input name="Submit2" type="reset" class="inputbut" value="重置"></td>
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