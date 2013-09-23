<!-- #include file="../Include/conn.asp" -->
<!-- #include file="Include/Chk.asp" -->
<!-- #include file="Include/fckeditor/fckeditor.asp" -->

<%
	id=request("id")
	if id="" or not isnumeric(id) then
	Response.Write "<script language='javascript'>alert('参数错误!');document.location.href('News_Manage.asp');</script>"
	Response.End()
	end if
		
	SQL="Select * from xin_About where id="&id
	set rs=server.createobject("adodb.recordset")
	rs.open SQL,conn,1,3
	if rs.eof and rs.bof then
	Response.Write "<script language='javascript'>alert('参数错误!');document.location.href('News_Manage.asp');</script>"
	Response.End()
	end if

if Request("Action")=1 then
		if Trim(Request("name"))="" or Trim(Request("tid"))="" or Trim(Request("info"))="" then
				Response.write "<script language='javascript'>alert('信息填写不完整：\n名称、显示位置、内容必须填写！');history.go(-1);</script>"
		Response.End()
	end if

	rs("url")=Trim(Request("url"))
	rs("name")=Trim(Request("name"))
	rs("tid")=Trim(Request("tid"))
	rs("ord")=Trim(Request("ord"))
	rs("info")=Trim(Request("info"))

	rs.Update
	rs.Close
	Set rs=nothing
	Response.Write "<script language='javascript'>alert('修改成功!');document.location.href('About_Manage.asp');</script>"
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
        <td width="50%" height="25" align="left" valign="middle" class="t2">修改单页</td>
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
              <td width="20%" height="30" align="right" valign="middle" class="bline"><span class="t3">名&nbsp;&nbsp;&nbsp;&nbsp;称：</span></td>
              <td width="80%" height="30" align="left" valign="middle" class="bline"><input name="name" type="text" class="main" id="name" value="<%=rs("name")%>" size="20" maxlength="20"></td>
            </tr>
            <tr>
              <td height="30" align="right" valign="middle" class="bline"><span class="t3">显示位置：</span></td>
              <td height="30" align="left" valign="middle" class="bline"><select name="tid" class="main" id="tid">
                <option value="<%=rs("tid")%>" selected>
				<%
				if rs("tid")=0 then
				response.Write("网站头部导航")
				else
				response.Write("网站底部导航")
				end if
				%></option>
				<%if rs("tid")=0 then%>
               
                <option value="1">网站底部导航</option>
				<%else%>
				 <option value="0">网站头部导航</option>
				 <%end if%>
              </select>              </td>
            </tr>
            <tr>
              <td height="30" align="right" valign="middle" class="bline"><span class="t3">顺&nbsp;&nbsp;&nbsp;&nbsp;序：</span></td>
              <td align="left" valign="middle" class="bline"><input name="ord" type="text" class="main" id="ord" value="<%=rs("ord")%>" size="4" maxlength="4">
                <span class="ts">*数值越小越靠前</span></td>
            </tr>
            <tr>
              <td height="30" align="right" valign="middle" class="bline"><span class="t3">网&nbsp;&nbsp;&nbsp;&nbsp;址：</span></td>
              <td align="left" valign="middle" class="bline"><input name="url" type="text" class="main" id="url" value="<%=rs("url")%>" size="50" maxlength="200">
                (若需要链接到其他页面,请输入网址,以http://开头)</td>
            </tr>
            <tr>
              <td width="20%" height="30" align="right" valign="middle" class="bline"><span class="t3">内&nbsp;&nbsp;&nbsp;&nbsp;容：</span></td>
              <td width="80%" align="left" valign="middle" class="bline">
                <p>
 <textarea name="info" id="Content" style="display:none"><%=rs("info")%></textarea>
		<IFRAME ID="ewebeditor1" src="eWebEditor/ewebeditor.asp?id=Content&style=s_coolblue" frameborder="0" scrolling="no" width=100% height="350"></IFRAME>    </p>                </td>
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