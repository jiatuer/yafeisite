<% 
Dim Query_Badword,Form_Badword,f_i,Err_Message,Err_Web,f_name

'------定义部份  头----------------------------------------------------------------------

Err_Message = 1		'css1

Err_Web = "Default.Asp"	'css2

Query_Badword="'∥and∥select∥update∥chr∥delete∥%20from∥;∥insert∥master.∥set∥chr(37)"     

'css3     

'----- css4

if request.QueryString<>"" then
Chk_badword=split(Query_Badword,"∥")
FOR EACH Query_Name IN Request.QueryString
for f_i=0 to ubound(Chk_badword)
If Instr(LCase(request.QueryString(Query_Name)),Chk_badword(f_i))<>0 Then
Select Case Err_Message
  Case "1"
Response.Write "<Script Language=JavaScript>alert('css5：and update delete ; ');window.close();</Script>"
  Case "2"
Response.Write "<Script Language=JavaScript>location.href='"&Err_Web&"'</Script>"
  Case "3"
Response.Write "<Script Language=JavaScript>alert('css6：and update delete ！');location.href='"&Err_Web&"';</Script>"
End Select
Response.End
End If
NEXT
NEXT
End if
%>