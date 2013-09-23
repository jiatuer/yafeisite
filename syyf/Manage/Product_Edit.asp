<!-- #include file="../Include/conn.asp" -->
<!-- #include file="Include/Chk.asp" -->
<!-- #include file="Include/fckeditor/fckeditor.asp" -->

<%
	id=request("id")
	if id="" or not isnumeric(id) then
	Response.Write "<script language='javascript'>alert('参数错误!');document.location.href('Product_Manage.asp');</script>"
	Response.End()
	end if
	
	SQL="Select * from xin_Product where id="&id
	set rs=server.createobject("adodb.recordset")
	rs.open SQL,conn,1,3
	if rs.eof and rs.bof then
	Response.Write "<script language='javascript'>alert('参数错误!');document.location.href('News_Manage.asp');</script>"
	Response.End()
	end if
	
if Request("Action")=1 then
	img=Trim(Request("img"))
	if img="" then
		img="noimages.gif"
	end if

	rs("title")=Trim(Request("title"))
	rs("tid")=Trim(Request("tid"))
	rs("img")=img
	rs("info")=Trim(Request("info"))
	rs("author")=Trim(Request("author"))
	rs("addtime")=Trim(Request("addtime"))
	rs("hit")=Trim(Request("hit"))
	rs("from")=Trim(Request("from"))
	rs("pro1")=Trim(Request("pro1"))
	rs("pro2")=Trim(Request("pro2"))
	rs("pro3")=Trim(Request("pro3"))
	rs("pro4")=Trim(Request("pro4"))

	rs.Update
	rs.Close
	Set rs=nothing
	Response.Write "<script language='javascript'>alert('产品修改成功!');document.location.href('Product_Manage.asp');</script>"
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
        <td width="50%" height="25" align="left" valign="middle" class="t2">发布产品</td>
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
		  <form name="form1" method="post" action="?action=1&id=<%=request("id")%>" onSubmit="return Validator.Validate(this,3);">
            <tr>
              <td width="20%" height="30" align="right" valign="middle" class="bline"><span class="t3">产品名称：</span></td>
              <td height="30" colspan="2" align="left" valign="middle" class="bline"><input name="title" type="text" class="main" id="name" value="<%=rs("title")%>" size="30" maxlength="50" dataType="Require"  msg="不能为空"></td>
            </tr>
            <tr>
              <td height="30" align="right" valign="middle" class="bline"><span class="t3">所属分类：</span></td>
              <td height="30" colspan="2" align="left" valign="middle" class="bline"><select name="tid" class="main" id="tid" dataType="Require"  msg="不能为空">
			  
			<%
			Set Rs1 = Conn.Execute("select * from xin_Product_Info where pid=0")
			Do While Not Rs1.EOF
			s=""
			if Cint(rs("tid"))=Cint(Rs1("id")) then
			s="selected"
			end if
			Response.Write "<option value='" & Rs1("id") & "'"&s&">"& Rs1("title") &"</option>"
			
				Set Rs2=Conn.Execute("select * from xin_Product_Info where pid="&Rs1("id")&"")
				Do While Not Rs2.EOF
				j=""
				if Cint(rs("tid"))=Cint(Rs2("id")) then
				j="selected"
				end if
				Response.Write "<option style='color:#999999' value='" & Rs2("id") & "' "&j&">--"& Rs2("title") &"</option>"
				Rs2.movenext
				loop
				rs2.close
			
			Rs1.MoveNext
			loop
			Rs1.Close
			
			%>

              </select>			</td>
            </tr>
			      <tr class="bline">
              <td height="30" align="right" valign="middle" class="bline"><span class="t3">其他项：</span></td>
              <td height="30" colspan="2" align="left" valign="middle" class="bline">来源：<input name="from" type="text" class="main" id="from" value="<%=rs("from")%>" size="15" maxlength="20"> 作者：<input name="author" type="text" class="main" id="author" value="<%=rs("author")%>" size="18"></td> 
		        </tr>
			<tr class="bline">
              <td height="30" align="right" valign="middle" class="bline"><span class="t3">其他项：</span></td>
              <td height="30" colspan="2" align="left" valign="middle" class="bline">时间：<input name="addtime" type="text" class="main" id="addtime" value="<%=rs("addtime")%>" size="15" maxlength="20"> 点击：<input name="hit" type="text" class="main" id="hit" value="<%=rs("hit")%>" size="18"></td> 
			</tr>
            <tr>
              <td height="40" align="right" valign="middle" class="bline"><span class="t3">上传图片：</span></td>
              <td width="40%" height="110" align="left" valign="middle" class="bline"><iframe allowTransparency="true" name='i1' src='upload/upload_pic.asp'; frameborder=0 width="350" scrolling=no height="40"> </iframe>
                <input name="img" type="hidden" id="img" value="<%=rs("img")%>"></td>
              <td width="40%" align="left" valign="middle" class="bline"><img src="../uploadfiles/upLoadImages/<%=rs("img")%>" width="180" height="120">此项为列表中的缩略图</td>
            </tr>
            <tr>
              <td width="20%" height="30" align="right" valign="middle" class="bline"><span class="t3">内容介绍：</span></td>
              <td colspan="2" align="left" valign="middle" class="bline">
                <p>
    <textarea name="info" id="Content" style="display:none"><%=rs("info")%></textarea>
		<IFRAME ID="ewebeditor1" src="eWebEditor/ewebeditor.asp?id=Content&style=s_coolblue" frameborder="0" scrolling="no" width=100% height="350"></IFRAME>
                </p>                </td>
            </tr>
            <tr>
              <td height="50" colspan="3" align="center" valign="middle"><input name="Submit" type="submit" class="inputbut" value="提交">
                &nbsp;
                <input name="Submit2" type="reset" class="inputbut" value="返回" onClick="history.back()"></td>
              </tr>
			</form>
			<script type="text/javascript" src="Include/Validator.js"></script>
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