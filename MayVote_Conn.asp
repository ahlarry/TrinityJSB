<%
'=============================
'Copyright (C) 006
'========�����������=============
Dim Conn,Connstr,NowTime,May_True,May_False,Action,CountTimeStart,CountTimeEnd
'=========�������ݿ�=============
CountTimeStart = timer
mdb = "database/#MayVote.mdb"  '���ݿ�·��
NowTime = "Now()"
May_True = 1
May_False = 0
ConnStr = "DBQ="+Server.MapPath(AdminDbPath &mdb)+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
'ConnStr = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(AdminDbPath &mdb)
Set Conn = Server.Createobject("adodb.connection")
On Error Resume Next 
Conn.Open Connstr

'=======���ݿ�������֤=======
If Err Then
Err.Clear
Set Conn = Nothing
Response.Write "���ݿ����ӳ�������ϵ����Ա��" 
Response.End 
End If
'============================

'======��ȡͶƱϵͳ��������========
Dim MayVote_Name
MayVote_Name = Application("MayVote_Name")
If  MayVote_Name = Empty Then
Dim rsConfig
Set rsConfig = Conn.Execute("Select * From MayVote_Config")
	If rsConfig.BOF And rsConfig.EOF Then
	rsConfig.Close
	Set rsConfig = Nothing
	Response.Write "ͶƱϵͳ�������ݶ�ʧ��ϵͳ�޷��������У�"
	Response.End
	Else
		Application.Lock
        Application("MayVote_Name") = rsConfig("MayVote_Name") 
		Application("MayVote_Url") = rsConfig("MayVote_Url") 
		Application("MayVote_Setting") = rsConfig("MayVote_Setting") 
		Application("MayVote_Ver") = rsConfig("MayVote_Ver")
		Application("MayVote_Copy") = rsConfig("MayVote_Copy") 
		Application.UnLock
	End If 
rsConfig.Close
Set rsConfig = Nothing
End If
'============================

'========���ݿ����ӹر�==========
Sub CloseConn()
On Error Resume Next
	If IsObject(Conn) Then
		Conn.Close
		Set Conn = Nothing
	End If 
End Sub

Sub MayVote_CopyRight()
Response.Write"Powered by <b>ͶƱϵͳ</b><b style='color:#FF9900'>"&Application("MayVote_Ver")&"</b> &nbsp;&nbsp;��Ȩ����&copy;&nbsp;<b><font color='#FF0000'>���ѿƼ���ģ������</font>&nbsp;2006-2007</b>"
End Sub
'============================
'���ƣ��ű�ִ��ʱ��
'���ã� Call CountTime()
'============================
Sub CountTime()
	CountTimeEnd=timer
	CountTimes=Trim(Cstr(CountTimeEnd-CountTimeStart))
	Response.Write CountTimes
End Sub
'============================
%>