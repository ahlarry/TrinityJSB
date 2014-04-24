<%
	'本页用到类xj_classs的实例xujian_ims
	dim c_lxrwlx
	c_lxrwlx = ""
	'取出零星任务类型
	set rs = xjweb.exec("select * from c_lxrwlx order by lxrwlx",1)
	do while not rs.eof
		if c_lxrwlx <> "" then 
			c_lxrwlx = c_lxrwlx & "|||" & rs("lxrwlx")
		else
			c_lxrwlx = rs("lxrwlx")
		end if
		rs.movenext
	loop
	rs.close
	c_lxrwlx = split(c_lxrwlx, "|||")
%>