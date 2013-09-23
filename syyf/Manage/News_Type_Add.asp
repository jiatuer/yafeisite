<!-- #include file="../Include/conn.asp" -->
<!-- #include file="Include/Chk.asp" -->

<%
if Request("pid")="" then
	n="主栏目"
	pid=0
else
	Set t=conn.execute("Select name from xin_New_Pros where id="&Request("pid"))
	if not t.eof then
		n=t(0)
		pid=Request("pid")
	else
		n="主栏目"
		pid=0
	end if
end if


if Request("Action")=1 then
	if Trim(Request("name"))=""  then
				Response.write "<script language='javascript'>alert('请输入栏目名称！');history.go(-1);</script>"
		Response.End()
	end if
		if not isnumeric(Trim(Request("ord")))   then
				Response.write "<script language='javascript'>alert('显示顺序必须是数字！');history.go(-1);</script>"
		Response.End()
	end if
	SQL="Select * from xin_New_Pros where name='"&Trim(Request("name"))&"'"
	set rs=server.createobject("adodb.recordset")
	rs.open SQL,conn,1,3
	if not rs.eof then
		Response.Write "<script language='javascript'>alert('已存在这个栏目,请重新输入!');history.back();</script>"
		Response.End()
	end if
	rs.AddNew
	rs("name")=Trim(Request("name"))
	rs("ord")=Trim(Request("ord"))
	rs("pid")=Trim(Request("pid"))

	rs.Update
	rs.Close
	Set rs=nothing
	Response.Write "<script language='javascript'>alert('您成功添加了一个栏目!');document.location.href('News_Type_Manage.asp');</script>"
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
        <td width="50%" height="25" align="left" valign="middle" class="t2">新增栏目</td>
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
          <table width="50%" border="0" cellspacing="0" cellpadding="0">
		  <form name="form1" method="post" action="?action=1">
            <tr>
              <td height="30" align="right" valign="middle" class="t3">所属栏目：</td>
              <td height="30" align="left" valign="middle"><%=n%>
			  <input name="pid" type="hidden" id="pid" value="<%=pid%>"></td>
            </tr>
            <tr>
              <td width="40%" height="30" align="right" valign="middle" class="t3">类别名称：</td>
              <td width="60%" height="30" align="left" valign="middle"><input name="name" type="text" class="main" id="name" size="20" maxlength="20"></td>
            </tr>
            <tr>
              <td width="40%" height="30" align="right" valign="middle" class="t3">显示顺序：</td>
              <td width="60%" height="30" align="left" valign="middle"><input name="ord" type="text" class="main" id="ord" value="0" size="4" maxlength="4">
                <span class="ts">*数值越小越靠前</span></td>
            </tr>
            <tr>
              <td height="50" colspan="2" align="center" valign="middle"><input name="Submit" type="submit" class="inputbut" value="提交">
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