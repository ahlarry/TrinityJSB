<%
Function TopTable()
	rw("<div id='loading' style=z-index:10000;visibility:hidden;position:'absolute';left:100;top:90;height:150;width:300;background-color:#FFFFFF;><Table cellpaddin=2 cellspacing=0 height=""100%"" width=""100%""><tr><td align=center><br><img src=""images/loading.gif""><br><br>��ҳ���ڼ�����.......  ���Ժ�!<br><br></td></tr></table></Div>")
	rw("<Script language=""javascript"">document.all.loading.style.visibility='visible';document.all.loading.style.left=(screen.width-300)/2;</script>")
	strCode=strCode & "<Table class=xtable width="""&web_info(8)&""" cellpadding=0 cellspacing=0 border=0>" &_
		vbcrlf & "<Tr><Td height=3 class=td_frame>" &_
		vbcrlf & "</td></Tr>" &_

		vbcrlf & "<Tr><Td class=ctd height=22>" &_
			vbcrlf & "<Table cellpadding=2 cellspacing=0 width=""100%"" height=""100%""><tr>" &_
			vbcrlf & "<td align=left width=350>&nbsp;&nbsp;Today: "&XjDate(2)&"</td>" &_
			vbcrlf & "<td align=right  width=*>���ٵ�¼</td>" &_
			vbcrlf & "</tr></table>" &_
		vbcrlf & "</td></Tr>" &_

		vbcrlf & "<Tr><Td class=ctd height=60>" &_
			vbcrlf & "<Table border=0 cellpadding=0 cellspacing=0 width=""100%"" height=""100%""><tr>" &_
			vbcrlf & "<td align=center width=*><img src="""web_info(9)&"""></td>" &_
			vbcrlf & "</tr></table>" &_
		vbcrlf & "</td></Tr>" &_

		vbcrlf & "<Tr><Td  class=ctd height=22>" &_
			vbcrlf & "<Table border=0 cellpadding=0 cellspacing=0 width=""100%"" height=""100%""><tr>" &_
			vbcrlf & "<td align=center width=*>"&mainmenu&"</td>" &_
			vbcrlf & "</tr></table>" &_
		vbcrlf & "</td></Tr>" &_
		vbcrlf & "</table>" &_

		vbcrlf & XjLine(2,web_info(8),"") &_

		vbcrlf & "<Table class=xtable width="""&web_info(8)&""" cellpadding=0 cellspacing=0 border=0>" &_
		vbcrlf & "<Tr><Td  class=ctd height=25>" &_
			vbcrlf & "<Table border=0 cellpadding=5 cellspacing=0 width=""100%"" height=""100%""><tr>" &_
			vbcrlf & "<td align=left width=*>ϵͳ֪ͨ</td>" &_
			vbcrlf & "<td align=right width=300>����Ϣ</td>" &_
			vbcrlf & "</tr></table>" &_
		vbcrlf & "</td></Tr>" &_
		vbcrlf & "</Table>" &_

		vbcrlf & XjLine(2,web_info(8),"") &_

		vbcrlf & "<Table class=xtable width="""&web_info(8)&""" cellpadding=0 cellspacing=0 border=0>" &_
		vbcrlf & "<Tr><Td  class=ctd height=25>" &_
			vbcrlf & "<Table border=0 cellpadding=5 cellspacing=0 width=""100%"" height=""100%""><tr>" &_
			vbcrlf & "<td align=left width=280>��ǰ��λ��--λ��λ��λ��--���µ�λ��λ��</td>" &_
			vbcrlf & "<td align=right width=*>"&pageLink("mtstat")&"</td>" &_
			vbcrlf & "</tr></table>" &_
		vbcrlf & "</td></Tr>" &_
		vbcrlf & "</Table>" &_

		vbcrlf & XjLine(2,web_info(8),"")
	TopTable=strCode
	rw(TopTable)
End Function

Function BottomTable()
	strCode=""
	strCode= XjLine(2,web_info(8),"") &_
		vbcrlf & "<Table class=xtable width="""&web_info(8)&""" cellpadding=0 cellspacing=0 border=0>" &_
		vbcrlf & "<Tr><Td  class=ctd height=22>" &_
			vbcrlf & "<Table border=0 cellpadding=0 cellspacing=0 width=""100%"" height=""100%""><tr>" &_
			vbcrlf & "<td align=center width=*>"&bottommenu&"</td>" &_
			vbcrlf & "</tr></table>" &_
		vbcrlf & "</td></Tr>" &_
		vbcrlf & "</table>" &_

		vbcrlf & XjLine(2,web_info(8),"") &_

		vbcrlf &"<Table class=xtable width="""&web_info(8)&""" cellpadding=0 cellspacing=0 border=0>" &_
		vbcrlf & "<Tr><Td  class=ctd height=22>" &_
			vbcrlf & "<Table border=0 cellpadding=0 cellspacing=0 width=""100%"" height=""100%""><tr>" &_
			vbcrlf & "<td align=center width=*>" &_
				vbcrlf & "��Ȩ����&copy;:���ѿƼ���ģ������ 2006-2007 &nbsp;&nbsp;&nbsp;" &_
				vbcrlf & "�������ݿ����: " & xjweb.opdbnum & " �� &nbsp;&nbsp;&nbsp;" &_
				vbcrlf & "ҳ��ִ��ʱ��: " & Round(((Timer()-StartTime)*1000),2) & " ���� &nbsp;&nbsp;&nbsp;" &_
				vbcrlf & "�������: " &_
			vbcrlf & "</td>" &_
			vbcrlf & "</tr></table>" &_
		vbcrlf & "</td></Tr>" &_
		vbcrlf & "<Tr><Td height=3 class=td_frame>" &_
		vbcrlf & "</td></Tr>" &_
		vbcrlf & "</table>"
	BottomTable=strCode
	rw(BottomTable)
	rw("<Script language=""javascript"">document.all.loading.style.visibility='hidden';</script>")
End Function

Function xjLine(iHeight, iWidth, xColor)
	strCode="<Table cellspacing=0 border=0 cellpadding=0 width="""&iWidth&"""><tr><td"
		If xColor="class" Then
			strCode=strCode & " class=td_frame "
		ElseIf xColor<>"" Then
			strCode=strCode & " style=background-color:" & xColor & "; "
		end if
		strCode=strCode & " height="&iHeight&"></td>" &_
			vbcrlf &  "</tr></table>"
		xjLine = strCode
End Function

Function TbTopic(info)
	strCode=""
	strCode="<Table cellspacing=0 border=0 cellpadding=3 width=""100%"">"&_
		vbcrlf & "<tr><td height=6></td></tr>"&_
		vbcrlf & "<tr><td align=center valign=middle>"&_
		vbcrlf & "<font style=font-size:15;font-weight:bold;>" & info & "</font>"&_
		vbcrlf & "</td></tr>"&_
		vbcrlf & "<tr><td height=3></td></tr>" &_
		vbcrlf & "</table>"
	TbTopic=strCode
End Function

Function XjDate(iKind)
	If Not isNumeric(iKind) Then XjDate=Now() : Exit Function
	Select Case iKind
		Case 1	'2005��1��1��
			XjDate=Year(Now) & "��" & Month(Now) & "��" & Day(Now) & "��"
		Case 2	'2005��1��1��������
			XjDate=Year(Now) & "��" & Month(Now) & "��" & Day(Now) & "�� ����"
				Select Case (Weekday(Now))
					Case 1
						XjDate=XjDate & "��"
					Case 2
						XjDate=XjDate & "һ"
					Case 3
						XjDate=XjDate & "��"
					Case 4
						XjDate=XjDate & "��"
					Case 5
						XjDate=XjDate & "��"
					Case 6
						XjDate=XjDate & "��"
					Case 7
						XjDate=XjDate & "��"
				End Select
		Case 3	'2005-1-1
			XjDate=Year(Now) & "-" & Month(Now) & "-" & Day(Now)
		Case Else
			XjDate=Now()
	End Select
End Function
%>