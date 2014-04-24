<%
'本页用到类xj_classs的实例xjweb
dim c_depart
c_depart=""
	'取出部门
	set rs = xjweb.exec("select * from c_depart order by depart",1)
	do while not rs.eof
		if c_depart <> "" then 
			c_depart = c_depart & "|||" & rs("depart")
		else
			c_depart = rs("depart")
		end if
		rs.movenext
	loop
	rs.close
c_depart=split(c_depart,"|||")
%>