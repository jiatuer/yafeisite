<HTML>
<HEAD>
<!--#include file="Admin_top.asp"-->

<TABLE width=98% height=30 border=0 align="center" cellPadding=0 cellSpacing=0 
      background=Images/main_bg.gif>
  <TBODY>
    <TR>
      <TD>　<strong><font color="#006699">反馈信息管理</font></strong></TD>
    </TR>
  </TBODY>
</TABLE>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="5">&nbsp;</td>
  </tr>
</table>
<TABLE width=98% border=0 align="center" cellPadding=5 cellSpacing=1 
      bgColor=#cdcdcd>
  <TBODY>
    <TR bgColor=#fafafa>
      <TD height=80><div align=center>
        <script language="JavaScript">
          <!--
            function MM_jumpMenu(targ,selObj,restore){ //v3.0
               eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
               if (restore) selObj.selectedIndex=0;
            }
          //-->
          </script>
        <table width="98%" height="94" border="0" align="center" id="table4">
          <tr>
            <td bgcolor="#F0F0F0" align="center" height="30" colspan="7"><p align="right">请选择相应的分类:
              <select size="1" name="D1" onChange="MM_jumpMenu('parent',this,0)">
                        <option>请选择</option>
                        <option value="Admin_Msg.asp?TypeClass=All">所有信息</option>
                        <%
  Set rs = Server.CreateObject("ADODB.RecordSet") 
  Sql = "select * from xin_Types"
  Rs.Open sql,conn,1,1
  If not rs.eof then
    Count = rs.recordcount
    For i = 1 to count
      response.write "<option value='Admin_Msg.asp?TypeClass="&rs("Type_Value")&"'>"&rs("Type_Value")&"</option>"
      rs.movenext
    Next
  End if
  rs.close
  Set rs=Nothing 
%>
                      </select>
            </td>
          </tr>
          <tr>
            <td bgcolor="#CDCDCD" align="center" height="19"><b>序号</b></td>
            <td bgcolor="#CDCDCD" align="center" height="19"><b>主题(点击可查看)</b></td>
            <td bgcolor="#CDCDCD" align="center" height="19"><b>类型</b></td>
            <td bgcolor="#CDCDCD" align="center" height="19"><b>反馈日期</b></td>
            <td bgcolor="#CDCDCD" align="center" height="19"><b>IP</b></td>
            <td bgcolor="#CDCDCD" align="center" height="19"><b>状态</b></td>
            <td bgcolor="#CDCDCD" align="center" height="19"><b>操作</b></td>
          </tr>
          <%
 TypeClass = Request("TypeClass")
 Set rs = Server.CreateObject("ADODB.RecordSet")
 If TypeClass = "" or TypeClass = "All" Then
    Sql = "select * from xin_Feedback where Feedback_fs = 'send' order by Feedback_id desc"
 Else
    Sql = "select * from xin_Feedback where Feedback_fs = 'send' and Feedback_Type = '"& TypeClass &"' order by Feedback_id desc"
 End If
 Rs.Open sql,conn,1,3
 If not rs.eof then
   If not isempty(request("page")) then   
    pagecount=cint(request("page"))   
   else   
    pagecount=1
   end if
   Rs.pagesize=15
   Rs.AbsolutePage=pagecount
   do while not Rs.eof
   i=i+1
%>
          <tr>
            <td bgcolor="#F0F0F0" align="center" height="22"><%=Rs("Feedback_id")%></td>
            <td bgcolor="#F0F0F0" align="center" height="22"><a href="#" onClick="javascript:window.open('Admin_Msg_View.asp?id=<%=Rs("Feedback_id")%>','InfoDetail','toolbar=no,scrollbars=no,resizable=yes,top=0,left=0,width=455 height=465'); "><u>
              <% If Len(Rs("Feedback_title"))>=10 Then Response.write left(Rs("Feedback_title"),10)+"..." Else Response.write Rs("Feedback_title") End if%>
            </u></a></td>
            <td bgcolor="#F0F0F0" align="center" height="22"><font color="#FF0000"><%=Rs("Feedback_type")%></font></td>
            <td bgcolor="#F0F0F0" align="center" height="22"><%=Rs("Feedback_datetime")%></td>
            <td bgcolor="#F0F0F0" align="center" height="22"><%=Rs("Feedback_ip")%></td>
            <td bgcolor="#F0F0F0" align="center" height="22"><% Select Case Rs("Feedback_zt")
					Case 0
					  response.write "<font color=#FF0000><b>未读</b></font>"
					Case 1
					  response.write "<b>已读</b>"
					End Select%></td>
            <td bgcolor="#F0F0F0" align="center" height="22"><a href="Admin_Action.asp?Action=1&id=<%=Rs("Feedback_id")%><% If Rs("Feedback_zt") = 0 Then Response.write "&zt=w" End if%>">删</a></td>
          </tr>
          <%
 If i>=Rs.pagesize Then exit do
 Rs.movenext
 loop
 Else
%>
          <tr>
            <td bgcolor="#F0F0F0" align="center" colspan="7" height="23">暂无记录</td>
          </tr>
          <%
 End If
%>
          <tr>
            <td bgcolor="#E3E3E3" align="center" colspan="6" height="20"><p>每页显示 <font color="#FF0000"><b><%=Rs.pagesize%></b></font> 条纪录 共有 <font color="#FF0000"><b><%=Rs.recordcount%></b></font> 条纪录 目前在第 <font color="#FF0000"><b><%=CINT(pagecount)%></b></font> 页
              <% if pagecount=1 and Rs.pagecount<>pagecount and Rs.pagecount<>0 then%>
                      <a href="?page=<%=cstr(pagecount+1)%>">下一页</a>
                      <% end if %>
                      <% if Rs.pagecount>1 and Rs.pagecount=pagecount then %>
                      <a href="?page=<%=cstr(pagecount-1)%>">上一页</a>
                      <%end if%>
                      <% if pagecount<>1 and Rs.pagecount<>pagecount then%>
                      <a href="?page=<%=cstr(pagecount-1)%>">上一页</a> <a href="?page=<%=cstr(pagecount+1)%>">下一页</a>
                      <% end if %>
            </td>
            <td bgcolor="#E3E3E3" align="center" height="20" width="41"><font color="#800000"><b> <a href="Admin_Action.asp?Action=6"> 清空</a></b></font></td>
          </tr>
        </table>
      </div></TD>
    </TR>
  </TBODY>
</TABLE>
</BODY></HTML>
<%
 Rs.close
 Set rs = nothing
%>