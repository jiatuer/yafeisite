<!-- #include file="../Include/conn.asp" -->
<!-- #include file="Include/Chk.asp" -->
<!-- #include file="Include/fckeditor/fckeditor.asp" -->

<%
img=request("img")
if Request("Action")=1 then
		if Trim(Request("titles"))="" or Trim(Request("lid"))="" or Trim(Request("info"))="" then
				Response.write "<script language='javascript'>alert('��Ϣ��д��������\n���⡢��Ŀ�����ݱ�����д��');history.go(-1);</script>"
		Response.End()
	end if

'===== Begin ͼƬ���Ŵ���
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

'===== End ͼƬ���Ŵ���	

	
	SQL="Select * from xin_New"
	set rs=server.createobject("adodb.recordset")
	rs.open SQL,conn,1,3

	rs.AddNew
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
	Response.Write "<script language='javascript'>alert('������ӳɹ�!');document.location.href('News_Add.asp');</script>"
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
        <td width="50%" height="25" align="left" valign="middle" class="t2">¼������</td>
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
              <td width="20%" height="30" align="right" valign="middle" class="bline"><span class="t3">���ű��⣺</span></td>
              <td width="80%" height="30" align="left" valign="middle" class="bline"><input name="titles" type="text" class="main" id="titles" size="30" maxlength="100">
                <span class="ts">                *����������100��֮�� 	</span></td>
            </tr>
            <tr>
              <td height="30" align="right" valign="middle" class="bline"><span class="t3">������Ŀ��</span></td>
              <td height="30" align="left" valign="middle" class="bline"><select name="lid" class="main" id="lid">
<%
Set Rs1 = Conn.Execute("select * from xin_New_Pros where pid=0")
Do While Not Rs1.EOF
Response.Write "<option value='" & Rs1("id") & "'>"& Rs1("name") &"</option>"
	
	Set Rs2=Conn.Execute("select * from xin_New_Pros where pid="&Rs1("id")&"")
	do while not Rs2.eof
	Response.Write "<option style='color:#999999' value='" & Rs2("id") & "'>--"& Rs2("name") &"</option>"
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
              <td height="30" align="right" valign="middle" class="bline"><span class="t3">�����</span></td>
              <td height="40" align="left" valign="middle" class="bline">���ߣ�
                <input name="author" type="text" class="main" id="author" size="15" maxlength="60" value="admin"> 
                ��Դ�� 
                <input name="from" type="text" class="main" id="from" size="20" maxlength="100" value="��վ"></td>
            </tr>
            <tr>
              <td height="30" align="right" valign="middle" class="bline"><span class="t3">�����</span></td>
              <td height="40" align="left" valign="middle" class="bline">ʱ�䣺
                <input name="addtime" type="text" class="main" id="addtime" size="30" maxlength="60" value="<%If isEdit Then Response.Write Rs("WriteTime"):Else Response.Write (Now())%>">
                �����
                <input name="hit" type="text" class="main" id="hit" size="10" maxlength="50" value="0"></td>
            </tr>
            <tr>
              <td height="30" align="right" valign="middle" class="bline"><span class="t3">�����</span></td>
              <td height="40" align="left" valign="middle" class="bline">����һ��
                <input name="news1" type="text" class="main" id="news1" size="18" value="">
                ���ö��� 
                <input name="news2" type="text" class="main" id="news2" size="18" value=""></td>
            </tr>
            <tr>
              <td height="30" align="right" valign="middle" class="bline"><span class="t3">�����</span></td>
              <td height="40" align="left" valign="middle" class="bline">��������
                <input name="news3" type="text" class="main" id="news3" size="18" value=""> 
                �����ģ� 
                <input name="news4" type="text" class="main" id="news4" size="18" value=""></td>
            </tr>
            <tr>
              <td width="20%" height="30" align="right" valign="middle" class="bline"><span class="t3">�������ݣ�<br>
              </span></td>
              <td width="80%" align="left" valign="middle" class="bline">
                <p>
                 <INPUT  type="hidden" name="info" id="Content" style="display:none">
<IFRAME ID="ewebeditor1" src="eWebEditor/ewebeditor.asp?id=Content&style=s_coolblue" frameborder="0" scrolling="no" width=95% height="350"></IFRAME>
                </p>                </td>
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