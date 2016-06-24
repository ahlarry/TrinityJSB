'**************************************************
'函数名：R
'作 用：过滤非法的SQL字符
'参 数：strChar-----要过滤的字符
'返回值：过滤后的字符
'**************************************************
Public Function R(strChar)
If strChar = "" Or IsNull(strChar) Then R = "":Exit Function
Dim strBadChar, arrBadChar, tempChar, I
'strBadChar = "$,#,',%,^,&,?,(,),<,>,[,],{,},/,\,;,:," & Chr(34) & "," & Chr(0) & ""
strBadChar = "+,',--,%,^,&,?,(,),<,>,[,],{,},/,\,;,:," & Chr(34) & "," & Chr(0) & ""
arrBadChar = Split(strBadChar, ",")
tempChar = strChar
For I = 0 To UBound(arrBadChar)
tempChar = Replace(tempChar, arrBadChar(I), "")
Next
tempChar = Replace(tempChar, "@@", "@")
R = tempChar
End Function 
'**************************************************
Function HtmlEncode(Content)
IF content="" or isnull(content) Then exit function
  Content = trim(Content)
  Content = Replace(Content,"%20"  , ""       )'特殊字符过滤
  Content = Replace(Content,chr(62),  "＞"  )' >
  Content = Replace(Content,chr(60),  "＜"  )' <
  Content = Replace(Content,chr(39),  "＇"    )' '
  Content = Replace(Content,chr(37),  "％"    )' %
  Content = Replace(Content, vbcrlf,  ""      )
  Content = Replace(Content,chr(34),  "”")' "
  Content = Replace(Content,chr(40),  "（"    )' (
  Content = Replace(Content,chr(41),  "）"    )' )
  Content = Replace(Content,chr(91),  "［"    )' [
  Content = Replace(Content,chr(93),  "］"    )' ]
  Content = Replace(Content,chr(123), "｛"    )' {
  Content = Replace(Content,chr(125), "｝"    )' }
  Content = Replace(Content, CHR(13),   "")  
  Content = Replace(Content,CHR(10), "")
  HtmlEncode = content
End Function 