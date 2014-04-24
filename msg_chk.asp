<!--#include file="include/conn.asp"-->
<%
dim action, strid, TmpStr
TmpStr=0
action=request("action")
select case action
	case "chknew"
		strSql="select * from [ims_message] where incept='" & session("userName") &"' and delr=0"
		call xjweb.exec("",-1)
		set rs=server.createobject("adodb.recordset")
		rs.open strSql,conn,1,3
		if rs.eof or rs.bof then
			response.write("document.write('<font color=#999999>&nbsp;</font>');")
		else
			strid=rs("id")
			response.write("document.write('<a href=# onclick=open_win(\'msg_dis.asp?new=true&id="&strid&"\',\'name\',\'500\',\'400\',\'yes\');><font  style=color:#ff0000;><b>"&rs.recordcount&"</b> 条系统信息</font></a>');")
			Do while not (rs.eof or TmpStr=1)
				If rs("flag")=0 Then
					response.write("document.write('<bgsound src=""images/uctrl/message.wav"">');")
					response.write("open_win('msg_dis.asp?new=true&id="&rs("id")&"','name','500','400','yes');")
					TmpStr=1
				End If
				rs.MoveNext
			Loop			
		end if
		rs.close
	case else
end select
strid=""
TmpStr=""
%>