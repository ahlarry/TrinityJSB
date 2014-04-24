<%
	'本页用到类xj_classs的实例xjweb
	
	dim c_sbcj, c_dwmc, c_jcjxh, c_dmmc, c_mjcl, c_rdogg, c_mtljcc, c_mtjg, c_dxjg, c_sxjg, c_gfxl
	c_sbcj = "" : c_dwmc = "" : c_jcjxh = "" : c_dmmc = "" : c_mjcl = "" : c_rdogg = ""
	c_mtljcc = "" : c_mtjg = "" : c_dxjg = "" : c_sxjg = "" : c_gfxl=""
	'取出断面名称
	set rs = xjweb.exec("select * from c_dmmc order by dmmc",1)
	do while not rs.eof
		if c_dmmc <> "" then 
			c_dmmc = c_dmmc & "|||" & rs("dmmc")
		else
			c_dmmc = rs("dmmc")
		end if
		rs.movenext
	loop
	rs.close
	c_dmmc = split(c_dmmc, "|||")
	
	'取出设备厂家
	set rs = xjweb.exec("select * from c_sbcj order by sbcj",1)
	do while not rs.eof
		if c_sbcj <> "" then 
			c_sbcj = c_sbcj & "|||" & rs("sbcj")
		else
			c_sbcj = rs("sbcj")
		end if
		rs.movenext
	loop
	rs.close
	c_sbcj = split(c_sbcj, "|||")
	
	'取出挤出机型号
	set rs = xjweb.exec("select * from c_jcjxh order by jcjxh",1)
	do while not rs.eof
		if c_jcjxh <> "" then 
			c_jcjxh = c_jcjxh & "|||" & rs("jcjxh")
		else
			c_jcjxh = rs("jcjxh")
		end if
		rs.movenext
	loop
	rs.close
	c_jcjxh = split(c_jcjxh, "|||")
	
	'取出单位名称
	set rs = xjweb.exec("select * from c_dwmc order by dwmc",1)
	do while not rs.eof
		if c_dwmc <> "" then 
			c_dwmc = c_dwmc & "|||" & rs("dwmc")
		else
			c_dwmc = rs("dwmc")
		end if
		rs.movenext
	loop
	rs.close
	c_dwmc = split(c_dwmc, "|||")
	
		'取出模具材料
	set rs = xjweb.exec("select * from c_mjcl order by mjcl",1)
	do while not rs.eof
		if c_mjcl <> "" then 
			c_mjcl = c_mjcl & "|||" & rs("mjcl")
		else
			c_mjcl = rs("mjcl")
		end if
		rs.movenext
	loop
	rs.close
	c_mjcl = split(c_mjcl, "|||")
	
	'取出热电偶规格
	set rs = xjweb.exec("select * from c_rdogg order by rdogg",1)
	do while not rs.eof
		if c_rdogg <> "" then 
			c_rdogg = c_rdogg & "|||" & rs("rdogg")
		else
			c_rdogg = rs("rdogg")
		end if
		rs.movenext
	loop
	rs.close
	c_rdogg = split(c_rdogg, "|||")
	
	'取出模头连接尺寸
	set rs = xjweb.exec("select * from c_mtljcc order by mtljcc",1)
	do while not rs.eof
		if c_mtljcc <> "" then 
			c_mtljcc = c_mtljcc & "|||" & rs("mtljcc")
		else
			c_mtljcc = rs("mtljcc")
		end if
		rs.movenext
	loop
	rs.close
	c_mtljcc = split(c_mtljcc, "|||")
	
	'取出模头结构
	set rs = xjweb.exec("select * from c_mtjg order by mtjg",1)
	do while not rs.eof
		if c_mtjg <> "" then 
			c_mtjg = c_mtjg & "|||" & rs("mtjg")
		else
			c_mtjg = rs("mtjg")
		end if
		rs.movenext
	loop
	rs.close
	c_mtjg = split(c_mtjg, "|||")
	
	'取出定型结构
	set rs = xjweb.exec("select * from c_dxjg order by dxjg",1)
	do while not rs.eof
		if c_dxjg <> "" then 
			c_dxjg = c_dxjg & "|||" & rs("dxjg")
		else
			c_dxjg = rs("dxjg")
		end if
		rs.movenext
	loop
	rs.close
	c_dxjg = split(c_dxjg, "|||")
	
	'取出水箱结构
	set rs = xjweb.exec("select * from c_sxjg order by sxjg",1)
	do while not rs.eof
		if c_sxjg <> "" then 
			c_sxjg = c_sxjg & "|||" & rs("sxjg")
		else
			c_sxjg = rs("sxjg")
		end if
		rs.movenext
	loop
	rs.close
	c_sxjg = split(c_sxjg, "|||")

	'取出规范系列
	set rs = xjweb.exec("select xl from c_gflb group by xl order by xl",1)
	do while not rs.eof
		if c_gfxl <> "" then 
			c_gfxl = c_gfxl & "|||" & rs("xl")
		else
			c_gfxl = rs("xl")
		end if
		rs.movenext
	loop
	rs.close
	c_gfxl = split(c_gfxl, "|||")	
%>