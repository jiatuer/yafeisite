<!-- #include file="SQL.asp" -->
<!-- #include file="Function.asp" -->
<%
'On Error Resume Next
   dim connstr,datapath,conn,dbpath
   datapath="/DataBase/xinhua.asp"
   connstr="Provider=Microsoft.JET.OLEDB.4.0;Data Source=" & Server.mappath(datapath)
   Set conn=Server.CreateObject("ADODB.Connection")
   conn.open connstr

%>