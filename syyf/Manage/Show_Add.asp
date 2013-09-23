<!-- #include file="../Include/conn.asp" -->
<!-- #include file="Include/Chk.asp" -->
<!-- #include file="Include/fckeditor/fckeditor.asp" -->

<%

if Request("Action")=1 then
	img=Trim(Request("img"))
	if img="" then
		img="noimages.gif"
	end if

	
	SQL="Select * from xin_Show"
	set rs=server.createobject("adodb.recordset")
	rs.open SQL,conn,1,3

	rs.AddNew
	rs("title")=Trim(Request("title"))
	rs("tid")=Trim(Request("tid"))
	rs("img")=img
	rs("info")=Trim(Request("info"))
	rs("author")=Trim(Request("author"))
	rs("addtime")=Trim(Request("addtime"))
	rs("hit")=Trim(Request("hit"))
	rs("from")=Trim(Request("from"))
	rs("show1")=Trim(Request("show1"))
	rs("show2")=Trim(Request("show2"))
	rs("show3")=Trim(Request("show3"))
	rs("show4")=Trim(Request("pro4"))

	rs.Update
	rs.Close
	Set rs=nothing
	Response.Write "<script language='javascript'>alert('产品发布成功!');document.location.href('Show_Add.asp');</script>"
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
		  <form name="form1" method="post" action="?action=1" onSubmit="return Validator.Validate(this,3);">
            <tr>
              <td width="20%" height="30" align="right" valign="middle" class="bline"><span class="t3">添加名称：</span></td>
              <td width="80%" height="30" align="left" valign="middle" class="bline"><input name="title" type="text" class="main" id="title" size="30" maxlength="50" dataType="Require"  msg="不能为空"></td>
            </tr>
            <tr>
              <td height="30" align="right" valign="middle" class="bline"><span class="t3">所属分类：</span></td>
              <td height="30" align="left" valign="middle" class="bline"><select name="tid" class="main" id="tid" dataType="Require"  msg="不能为空">
<%
Set Rs1 = Conn.Execute("select * from xin_Show_Info where pid=0")
Do While Not Rs1.EOF
Response.Write "<option value='" & Rs1("id") & "'>"& Rs1("title") &"</option>"

	Set Rs2=Conn.Execute("select * from xin_Show_Info where pid="&Rs1("id")&"")
	do while not Rs2.eof
	Response.Write "<option style='color:#999999' value='" & Rs2("id") & "'>--"& Rs2("title") &"</option>"
	Rs2.movenext
	loop
	rs2.close

Rs1.MoveNext
loop
Rs1.Close

%>
              </select>              </td>
            </tr>
			 <tr>
              <td height="30" align="right" valign="middle" class="bline"><span class="t3">其他项：</span></td>
              <td height="40" align="left" valign="middle" class="bline">作者：
                <input name="author" type="text" class="main" id="author" size="15" maxlength="250" value="admin"> 
                来源： 
                <input name="from" type="text" class="main" id="from" size="20" maxlength="250" value="本站"></td>
            </tr>
            <tr>
              <td height="30" align="right" valign="middle" class="bline"><span class="t3">其他项：</span></td>
              <td height="40" align="left" valign="middle" class="bline">时间：
                <input name="addtime" type="text" class="main" id="addtime" size="30" maxlength="250" value="<%If isEdit Then Response.Write Rs("WriteTime"):Else Response.Write (Now())%>">
                点击：
                <input name="hit" type="text" class="main" id="hit" size="10" maxlength="250" value="0"></td>
            </tr>
            <tr>
              <td height="30" align="right" valign="middle" class="bline"><span class="t3">上传图片（用于缩略图）：  </span></td>
              <td height="40" align="left" valign="middle" class="bline"><iframe allowTransparency="true" name='i1' src='upload/upload_pic.asp'; frameborder=0 width="350" scrolling=no height="40"> </iframe>
                <input name="img" type="hidden" id="img"></td>
            </tr>
            <tr>
              <td width="20%" height="30" align="right" valign="middle" class="bline"><span class="t3">内容介绍：</span></td>
              <td width="80%" align="left" valign="middle" class="bline">
                <p><INPUT  type="hidden" name="info" id="Content" style="display:none">
<IFRAME ID="ewebeditor1" src="eWebEditor/ewebeditor.asp?id=Content&style=s_coolblue" frameborder="0" scrolling="no" width=95% height="350"></IFRAME>
                </p>                </td>
            </tr>
            <tr>
              <td height="50" colspan="2" align="center" valign="middle"><input name="Submit" type="submit" class="inputbut" value="提交">
                &nbsp; <input name="Submit2" type="reset" class="inputbut" value="重置"></td>
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