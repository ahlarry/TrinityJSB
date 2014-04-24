<%
'=============================
'Copyright (C) 006
'========申明常规变量=============
Dim Conn,Connstr,NowTime,May_True,May_False,Action,CountTimeStart,CountTimeEnd
'=========连接数据库=============
CountTimeStart = timer
mdb = "database/#MayVote.mdb"  '数据库路径
NowTime = "Now()"
May_True = 1
May_False = 0
ConnStr = "DBQ="+Server.MapPath(AdminDbPath &mdb)+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
'ConnStr = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(AdminDbPath &mdb)
Set Conn = Server.Createobject("adodb.connection")
On Error Resume Next 
Conn.Open Connstr

'=======数据库连接验证=======
If Err Then
Err.Clear
Set Conn = Nothing
Response.Write "数据库连接出错，请联系管理员。" 
Response.End 
End If
'============================

'======读取投票系统基本设置========
Dim MayVote_Name
MayVote_Name = Application("MayVote_Name")
If  MayVote_Name = Empty Then
Dim rsConfig
Set rsConfig = Conn.Execute("Select * From MayVote_Config")
	If rsConfig.BOF And rsConfig.EOF Then
	rsConfig.Close
	Set rsConfig = Nothing
	Response.Write "投票系统配置数据丢失！系统无法正常运行！"
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

'========数据库连接关闭==========
Sub CloseConn()
On Error Resume Next
	If IsObject(Conn) Then
		Conn.Close
		Set Conn = Nothing
	End If 
End Sub

Sub MayVote_CopyRight()
Response.Write"Powered by <b>投票系统</b><b style='color:#FF9900'>"&Application("MayVote_Ver")&"</b> &nbsp;&nbsp;版权所有&copy;&nbsp;<b><font color='#FF0000'>三佳科技挤模技术部</font>&nbsp;2006-2007</b>"
End Sub
'============================
'名称：脚本执行时间
'调用： Call CountTime()
'============================
Sub CountTime()
	CountTimeEnd=timer
	CountTimes=Trim(Cstr(CountTimeEnd-CountTimeStart))
	Response.Write CountTimes
End Sub
'============================
%>