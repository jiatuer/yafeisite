<%
  '����Σ���ַ�
    Function YuQaIFS_Eck(InChar)
        Dim NewChar
        NewChar = Replace(Inchar,"and","")
        NewChar = Replace(Inchar,"1=1","")
        NewChar = Replace(Inchar,"or","")
        NewChar = Replace(Inchar,"=","")
        NewChar = Replace(Inchar,"'","")
        NewChar = Replace(NewChar,"_","")
        NewChar = Replace(NewChar,"*","")
        NewChar = Replace(NewChar,"	","")
        NewChar = Replace(NewChar,chr(34),"")
        NewChar = Replace(NewChar,chr(39),"")
        NewChar = Replace(NewChar,chr(37),"")
        NewChar = Replace(NewChar,"{","")
        NewChar = Replace(NewChar,"}","")
        ' ������ NewChar = Replace(NewChar,chr(91),"")
        ' ������ NewChar = Replace(NewChar,chr(93),"")
        ' ð�� NewChar = Replace(NewChar,chr(58),"")
        NewChar = Replace(NewChar,chr(59),"")
        ' �Ӻ� NewChar = Replace(NewChar,chr(43),"")
        YuQaIFS_Eck = NewChar
    End Function 
%>