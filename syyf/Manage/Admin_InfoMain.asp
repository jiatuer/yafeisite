<HTML>
<HEAD>
<!--#include file="Admin_top.asp"-->
<%
 Set rs = Server.CreateObject("ADODB.RecordSet") 
 Sql = "select * from xin_Feedback Where Feedback_fs = 'send'"
 Rs.Open sql,conn,1,3
 if not rs.eof then
   Send = rs.recordcount
 else
   Send = 0
 End if
 Rs.close
 Sql = "select * from xin_Feedback Where Feedback_fs = 'Mail'"
 Rs.Open sql,conn,1,3
 if not rs.eof then
   Mail = rs.recordcount
 else
   Mail = 0
 End if 
 Rs.close
 Set rs = nothing 
%>
<TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
  <TBODY>
  <TR>
    <TD width=1 bgColor=#e3e3e3></TD>
    <TD width=744><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td>&nbsp;</td>
      </tr>
    </table>
    <TABLE width=699 height=22 border=0 align="center" cellPadding=3 cellSpacing=1 
      bgColor=#cdcdcd>
      <TBODY>
        <TR bgColor=#fafafa>
          <TD width="100%" height=20 bgcolor="#E3E3E3"><p align="center"><b>&nbsp;&nbsp; 网站共有<u><font color=#FF0000><%=Send+Mail%></font></u>条反馈信息,请注意查收!</b></TD>
        </TR>
      </TBODY>
    </TABLE></TD>
  </TR></TBODY></TABLE>
</BODY></HTML>