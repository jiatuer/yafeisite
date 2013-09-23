<!-- #include file="../Include/conn.asp" -->
<!-- #include file="Include/Chk.asp" -->
<!-- #include file="Include/fckeditor/fckeditor.asp" -->

<%
	id=request("id")
	if id="" or not isnumeric(id) then
	Response.Write "<script language='javascript'>alert('参数错误!');document.location.href('News_Manage.asp');</script>"
	Response.End()
	end if
	img=request("img")
	SQL="Select * from xin_New where id="&id
	set rs=server.createobject("adodb.recordset")
	rs.open SQL,conn,1,3
	if rs.eof and rs.bof then
	Response.Write "<script language='javascript'>alert('参数错误!');document.location.href('News_Manage.asp');</script>"
	Response.End()
	end if
if Request("Action")=1 then
		if Trim(Request("titles"))="" or Trim(Request("lid"))="" or Trim(Request("info"))="" then
				Response.write "<script language='javascript'>alert('信息填写不完整：\n标题、栏目、内容必须填写！');history.go(-1);</script>"
		Response.End()
	end if

'===== Begin 图片新闻处理
	  nr=Request("info")
	  nr=replace(nr,"<IMG","<img")
	  nr=replace(nr,"SRC=","src=")
	  nr=replace(nr,"","")
	  nr=replace(nr,"","")
	  nr=replace(nr,"","")
	  if instr(nr,"<img")<>0 and instr(nr,"src=")<>0 then
	     nr=replace(nr,"""","")
	     aa=Split(nr,"<img")
	     bb=aa(1)
	     cc=split(bb,">")
	     dd=cc(0)
	     ee=split(dd,"src=")
	     ff=ee(1)
	     gg=split(ff," ")
	     img=gg(0) 
	  end if
	
	if right(img,3)<>"jpg" then img=""
'Response.Write(img)
'Response.End()

'===== End 图片新闻处理	




	rs("titles")=Trim(Request("titles"))
	rs("lid")=Trim(Request("lid"))
	rs("from")=Trim(Request("from"))
	rs("info")=Trim(Request("info"))
		rs("author")=Trim(Request("author"))
	rs("addtime")=Trim(Request("addtime"))
	rs("hit")=Trim(Request("hit"))
	rs("news1")=Trim(Request("news1"))
	rs("news2")=Trim(Request("news2"))
	rs("news3")=Trim(Request("news3"))
	rs("news4")=Trim(Request("news4"))
	if img<>"" and Request.Form("img")=1 then rs("img")=img


	rs.Update
	rs.Close
	Set rs=nothing
	Response.Write "<script language='javascript'>alert('修改成功!');document.location.href('News_Manage.asp');</script>"
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
        <td width="50%" height="25" align="left" valign="middle" class="t2">修改新闻</td>
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
              <td width="20%" height="30" align="right" valign="middle" class="bline"><span class="t3">新闻标题：</span></td>
              <td width="80%" height="30" align="left" valign="middle" class="bline"><input name="titles" type="text" class="main" id="titles" value="<%=rs("titles")%>" size="30" maxlength="100">
                <span class="ts">                *标题字数限100字之内 	</span></td>
            </tr>
            <tr>
              <td height="30" align="right" valign="middle" class="bline"><span class="t3">所属栏目：</span></td>
              <td height="30" align="left" valign="middle" class="bline"><select name="lid" class="main" id="lid">
<%
Set Rs1 = Conn.Execute("select * from xin_New_Pros where pid=0")
Do While Not Rs1.EOF
s=""
if Cint(rs("lid"))=Cint(Rs1("id")) then
s="selected"
end if
Response.Write "<option value='" & Rs1("id") & "' "&s&">"& Rs1("name") &"</option>"

	Set Rs2=Conn.Execute("select * from xin_New_Pros where pid="&Rs1("id")&"")
	Do While Not Rs2.EOF
	j=""
	if Cint(rs("lid"))=Cint(Rs2("id")) then
	j="selected"
	end if
	Response.Write "<option style='color:#999999' value='" & Rs2("id") & "' "&j&">--"& Rs2("name") &"</option>"
	Rs2.movenext
	loop
	rs2.close

Rs1.MoveNext
loop
Rs1.Close
%>
              </select>              </td>
            </tr>
            <tr class="bline">
              <td height="30" align="right" valign="middle" class="bline"><span class="t3">其他项：</span></td>
              <td height="30" align="left" valign="middle" class="bline">来源：<input name="from" type="text" class="main" id="from" value="<%=rs("from")%>" size="15" maxlength="20"> 作者：<input name="author" type="text" class="main" id="author" value="<%=rs("author")%>" size="18"></td>
            </tr>
			<tr class="bline">
              <td height="30" align="right" valign="middle" class="bline"><span class="t3">其他项：</span></td>
              <td height="30" align="left" valign="middle" class="bline">时间：<input name="addtime" type="text" class="main" id="addtime" value="<%=rs("addtime")%>" size="15" maxlength="20"> 点击：<input name="hit" type="text" class="main" id="hit" value="<%=rs("hit")%>" size="18"></td>
            </tr>
			<tr class="bline">
              <td height="30" align="right" valign="middle" class="bline"><span class="t3">备用项：</span></td>
              <td height="30" align="left" valign="middle" class="bline">备用一：<input name="news1" type="text" class="main" id="news1" value="<%=rs("news1")%>" size="15"> 备用二：<input name="news2" type="text" class="main" id="news2" value="<%=rs("news2")%>" size="18"></td>
            </tr>
			<tr class="bline">
              <td height="30" align="right" valign="middle" class="bline"><span class="t3">备用项：</span></td>
              <td height="30" align="left" valign="middle" class="bline">备用三：<input name="news3" type="text" class="main" id="news3" value="<%=rs("news3")%>" size="15"> 备用四：<input name="news4" type="text" class="main" id="news4" value="<%=rs("news4")%>" size="18"></td>
            </tr>
            <tr>
              <td height="30" align="right" valign="middle" class="bline"><span class="t3">图片新闻：</span></td>
              <td height="30" align="left" valign="middle" class="bline"><input name="img" type="checkbox" class="main" id="img" value="1" <%if rs("img")<>"" then Response.Write("checked")%>>
                <span class="ts">*选中为图片新闻,系统会自动获取第一个jpg图片</span></td>
            </tr>
            <tr>
              <td width="20%" height="30" align="right" valign="middle" class="bline"><span class="t3">新闻内容：</span></td>
              <td width="80%" align="left" valign="middle" class="bline">
                <p>
 <textarea name="info" id="Content" style="display:none"><%=rs("info")%></textarea>
		<IFRAME ID="ewebeditor1" src="eWebEditor/ewebeditor.asp?id=Content&style=s_coolblue" frameborder="0" scrolling="no" width=100% height="350"></IFRAME>          </p>                </td>
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