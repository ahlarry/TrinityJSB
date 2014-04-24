<%
	'本页用到类xj_classs的实例xujian_ims
	'从库中取出所有的分值系数(所有的分值系数均相对于模具总分)
	
	dim xs_jg, xs_sj, xs_mtjg, xs_mtsj, xs_dxjg, xs_dxsj, xs_sh, xs_fg, xs_fgsh, xs_fc, xs_bom, xs_tsd, xs_ts
	xs_jg=0 : xs_sj=0 : xs_mtjg＝0 : xs_mtsj=0 : xs_dxjg=0 : xs_dxsj : xs_sh=0 : xs_fg=0 : xs_fgsh=0 : xs_fc=0 : xs_bom=0 : xs_tsd=0 : xs_ts=0
	'取出断面名称
	set rs = xujian_ims.exec("select * from c_fzbl",1)
	if not(rs.eof or rs.bof) then
		xs_jg=rs("jgbl")
		xs_sj=rs("sjbl")
		xs_mtjg=rs("mtjgbl")
'		xs_mtsj=1-xs_mtjg
		xs_dxjg=rs("dxjgbl")
'		xs_dxsj=1-xs_dxjg
		xs_sh=rs("shbl")
		xs_fg=rs("fgbl")
		xs_fgsh=rs("fgshbl")
		xs_fc=rs("fcbl")
		xs_bom=rs("bombl")
		xs_tsd=rs("tsdbl")
		xs_ts=rs("tsbl")
	end if
	rs.close
	Response.Write(xs_mtjg)
%>