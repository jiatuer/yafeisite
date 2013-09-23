<%
 action = request("action")
%>
<HTML>
<HEAD>
<!--#include file="Admin_top.asp"-->

<TABLE width=98% height=30 border=0 align="center" cellPadding=0 cellSpacing=0 
      background=Images/main_bg.gif>
  <TBODY>
    <TR>
      <TD>　<strong><font color="#006699">来自何方下拉菜单管理</font></strong></TD>
    </TR>
  </TBODY>
</TABLE>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="5">&nbsp;</td>
  </tr>
</table>
<TABLE height=181 cellSpacing=1 cellPadding=5 width=98% align=center 
      bgColor=#cdcdcd border=0>
  <TBODY>
    <%
 Select Case Action
  Case 1
%>
    <TR bgColor=#fafafa>
      <TD height=80><div align=center>
        <table width="98%" height="94" border="0" align="center" id="table4">
          <tr>
            <td bgcolor="#CDCDCD" align="center" width="26%" height="19">序号</td>
            <td bgcolor="#CDCDCD" align="center" width="54%" height="19">名称</td>
            <td bgcolor="#CDCDCD" align="center" height="19" width="16%">操作</td>
          </tr>
          <%
 Set rs = Server.CreateObject("ADODB.RecordSet") 
 Sql = "select * from xin_Froms"
 Rs.Open sql,conn,1,3
 If not rs.eof then
 Count = rs.recordcount
 for i = 1 to Count
%>
          <tr>
            <td width="26%" bgcolor="#F0F0F0" align="center" height="22"><%=Rs("From_id")%></td>
            <td width="54%" bgcolor="#F0F0F0" align="center" height="22"><b><%=Rs("From_Value")%></b></td>
            <td bgcolor="#F0F0F0" align="center" height="22" width="16%"> <a href="Admin_From_Tion.asp?Action=4&id=<%=rs("From_id")%>">删</a></td>
          </tr>
          <%
 rs.movenext
 Next
 Else
%>
          <tr>
            <td bgcolor="#F0F0F0" align="center" colspan="3" height="23">暂无记录</td>
          </tr>
          <%
 End If
 Rs.close
 Set rs = nothing
%>
          <tr>
            <td bgcolor="#E3E3E3" align="center" colspan="3" height="20"><b> <u><a href="Admin_From.asp?action=2">新增下拉菜单</a></u></b> <font color="#800000"><b> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href="Admin_From_Tion.asp?Action=5"> <font color="#FF0000">所有数据清空(慎用)</font></a></b></font></td>
          </tr>
        </table>
      </div></TD>
    </TR>
    <%
  Case 2
%>
    <TR bgColor=#fafafa>
      <TD height=80><div align=center>
        <table width="98%" height="63" border="0" align="center" id="table4">
          <tr>
            <td bgcolor="#CDCDCD" align="center" width="96%" height="21"> 增加记录</td>
          </tr>
          <form method="POST" action="Admin_From_Tion.asp?action=2" name="YuQaIFS_Feedback_Add">
            <tr>
              <td bgcolor="#F0F0F0" align="center" height="22"><input type="text" name="Title" size="20">
                    <input type="submit" value="提交" name="B1"></td>
            </tr>
          </form>
          <tr>
            <td bgcolor="#E3E3E3" align="center" height="22"><b>注：最多不能超过10个字</b></td>
          </tr>
        </table>
      </div></TD>
    </TR>
    <%
  Case 3
%>
    <TR bgColor=#fafafa>
      <TD height=80><div align=center>
        <%
 id = request("id")
 Set rs = Server.CreateObject("ADODB.RecordSet") 
 Sql = "select * from xin_Froms Where From_id = "&id
 Rs.Open sql,conn,1,3
 If not rs.eof then
   Title_Name = rs("From_Value")
 Else
   response.write"<script language=JavaScript>{window.alert('对不起,您查询的记录不存在');window.location.href='javascript:history.go(-1)'}</script>" 
 End if
 Rs.close
 Set rs = nothing  
%>
        <table width="98%" height="63" border="0" align="center" id="table4">
          <tr>
            <td bgcolor="#CDCDCD" align="center" width="96%" height="21"> 修改记录</td>
          </tr>
          <form method="POST" action="Admin_From_Tion.asp?action=3&id=<%=id%>" name="YuQaIFS_Feedback_Modi">
            <tr>
              <td bgcolor="#F0F0F0" align="center" height="22"><input type="text" name="Title2" size="20" value="<%=Title_Name%>">
                    <input type="submit" value="提交" name="B1"></td>
            </tr>
          </form>
          <tr>
            <td bgcolor="#E3E3E3" align="center" height="22"><b>注：最多不能超过10个字</b></td>
          </tr>
        </table>
      </div></TD>
    </TR>
    <%
 Case else
%>
    <TR bgColor=#fafafa>
      <TD height=80><div align=center> 错误的参数 </div></TD>
    </TR>
    <%
 End Select
%>
  </TBODY>
</TABLE>
</BODY></HTML>