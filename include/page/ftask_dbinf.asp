<%
	'��ҳ�õ���xj_classs��ʵ��xujian_ims
	dim c_lxrwlx
	c_lxrwlx = ""
	'ȡ��������������
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