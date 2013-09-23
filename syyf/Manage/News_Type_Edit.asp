<!-- #include file="../Include/conn.asp" -->
<!-- #include file="Include/Chk.asp" -->

<%
	id=request("id")
		if id="" or not isnumeric(id) then
	Response.Write "<script language='javascript'>alert('参数错误!');document.location.href('News_Type_Manage.asp');</script>"
	Response.End()
	end if

	SQL="Select * from xin_New_Pros where id="&request("id")
	set rs=server.createobject("adodb.recordset")
	rs.open SQL,conn,1,3
	if rs.eof and rs.bof then
	Response.Write "<script language='javascript'>alert('参数错误!');document.location.href('News_Type_Manage.asp');</script>"
	Response.End()
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



	rs("name")=Trim(Request("name"))
	rs("ord")=Trim(Request("ord"))
	rs("pid")=Trim(Request("pid"))

	rs.Update
	rs.Close
	Set rs=nothing
	Response.Write "<script language='javascript'>alert('修改成功!');document.location.href('News_Type_Manage.asp');</script>"
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
        <td width="50%" height="25" align="left" valign="middle" class="t2">修改栏目</td>
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
		  <form name="form1" method="post" action="?action=1&id=<%=request("id")%>">
            <tr>
              <td height="30" align="right" valign="middle" class="t3">所属栏目：</td>
              <td height="30" align="left" valign="middle">	 
			   <select name="pid" class="main" id="pid">
	   <%if rs("pid")=0 then 
	   		Response.Write("<option value='0' selected >主栏目</option>")
		else
			sql="Select id,name from xin_New_Pros where id="&rs("pid")
			Set t=conn.execute(sql)
			if not t.eof then
			 	Response.Write("<option value='"&t("id")&"' selected>"&t("name")&"</option>")
			else
				Response.Write("<option value='0' selected >错误</option>")
			end if
		end if
		sql="Select id,name from xin_New_Pros where pid=0 order by ord asc"
		Set t=conn.execute(sql)
		do while not t.eof
			Response.Write("<option value='"&t("id")&"'>"&t("name")&"</option>")
		t.MoveNext
		loop
		t.Close
		Set t=nothing		
	   %>
      </select></td>
            </tr>
            <tr>
              <td width="40%" height="30" align="right" valign="middle" class="t3">栏目名称：</td>
              <td width="60%" height="30" align="left" valign="middle"><input name="name" type="text" class="main" id="name" value="<%=rs("name")%>" size="20" maxlength="20"></td>
            </tr>
            <tr>
              <td width="40%" height="30" align="right" valign="middle" class="t3">显示顺序：</td>
              <td width="60%" height="30" align="left" valign="middle"><input name="ord" type="text" class="main" id="ord" value="<%=rs("ord")%>" size="4" maxlength="4">
                <span class="ts">*数值越小越靠前</span></td>
            </tr>
            <tr>
              <td height="50" colspan="2" align="center" valign="middle"><input name="Submit" type="submit" class="inputbut" value="修改">
                &nbsp;
                <input name="Submit2" type="reset" class="inputbut" value="返回" onClick="history.back()"></td>
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