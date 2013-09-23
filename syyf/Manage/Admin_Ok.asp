<!-- #include file="../Include/conn.asp" -->
<!--#include file="../Inc/xin_Config.asp" -->
<!--#include file="../Inc/xin_Function.asp"-->
<!--#include file="../Inc/xin_Md5.asp" -->
<%
 If Config_Web_safe Then
   From_url = Cstr(Request.ServerVariables("HTTP_REFERER"))
   Serv_url = Cstr(Request.ServerVariables("SERVER_NAME"))
   if mid(From_url,8,len(Serv_url)) <> Serv_url then
     response.write"<script language=JavaScript>{window.alert('对不起,本系统开启了防外部数据提交');window.location.href='javascript:window.close()'}</script>"
     response.end
   end if
 End if
%>
<%
  name=YuQaIFS_Eck(request.form("Admin_User"))
  pass=Md5(YuQaIFS_Eck(request.form("Admin_Pass")),Pmode)
  Sql="select * from xin_Admin where Admin_User='"&name&"'"
  Set rs=Server.CreateObject("ADODB.RecordSet") 
  Rs.Open sql,conn,1,3
  If Not Rs.eof then
    If Rs("Admin_Pass") <> pass Then
      response.write"<script language=JavaScript>{window.alert('对不起,密码错误');window.location.href='javascript:history.go(-1)'}</script>"
    Else
      If rs("Admin_Lock") =0 then
        rs("Admin_Datetime") = now()
        rs("Admin_Ip") = Request.ServerVariables("REMOTE_ADDR")
        rs.Update   
        Session("Admin_User")=rs("Admin_User")
        Session("Admin_Ok")=True
        Session.timeout=900
        Response.Redirect "Admin_InfoMain.asp"
      Else
        response.write"<script language=JavaScript>{window.alert('您的帐号由于您的操作不当，已经被管理员锁定！如有问题，请联系管理员！');window.location.href='javascript:history.go(-1)'}</script>"
      End if
    End if
  Else
    response.write"<script language=JavaScript>{window.alert('对不起,帐号不存在');window.location.href='javascript:history.go(-1)'}</script>"
  End if
  Rs.close
  Set rs = nothing
%>