<!--#include file = "include/conn.asp"-->
<%
	Call ChkPageAble(7)
	dim ininf
	ininf = lcase(request.form("indbinf"))
	dim strddh, strlsh, strdwmc, strdmmc, strmh, strdiskid, strbz ,dtcpsj,Xstrlsh,lsno 
      Xstrlsh=""
	select case ininf
		case "add"
            lsno = trim(request("lsno"))
			strdiskid = trim(request("diskid"))
			strbz = trim(request("bz"))
         for i=1 to lsno

			strddh = trim(request("ddh"&i))
			strlsh = trim(request("lsh"&i))
			strdwmc = trim(request("dwmc"&i))
			strdmmc = trim(request("dmmc"&i))
			strmh = trim(request("mh"&i))
			dtcpsj = now()
			
			strSql="select * from [doc_bak] where lsh = '"&strlsh&"'"
			Call xjweb.Exec("",-1)
			rs.open strSql, conn, 1,3
			if not (rs.eof or rs.bof) then
				Call JsAlert("��ˮ�� ��" & strlsh & "�� ģ���Ѵ浵��","")
			else
				rs.addnew
				rs("ddh")=strddh
				rs("lsh")=strlsh
				rs("dwmc")=strdwmc
				rs("mh")=strmh
				rs("diskid")=strdiskid
				if strbz<>"" then rs("bz")=strbz
				rs("cpsj")=dtcpsj
				rs.update
				rs.close
				strSql = "update [mtask] set cp = true where lsh = '"&strlsh&"'"
				Call xjweb.Exec(strSql, 1)
			end if				

Xstrlsh=Xstrlsh&strlsh&" "
next

				Call JsAlert("��ˮ�� ��" & Xstrlsh & "�� ģ�ߴ浵��Ϣ��ӳɹ�!","docbak_add.asp")
		case "change"
			strddh = trim(request("ddh"))
			strlsh = trim(request("lsh"))
			strdwmc = trim(request("dwmc"))
			strdmmc = trim(request("dmmc"))
			strmh = trim(request("mh"))
			strdiskid = trim(request("diskid"))
			strbz = trim(request("bz"))
			dtcpsj = now()
			
			strSql="select * from [doc_bak] where lsh = '"&strlsh&"'"
			Call xjweb.Exec("",-1)
			rs.open strSql, conn, 1,3
			if rs.eof or rs.bof then
				Call JsAlert("��ˮ�� ��" & strlsh & "�� ��ģ����δ�浵��","docbak_change.asp")
			else
				rs("ddh")=strddh
				rs("lsh")=strlsh
				rs("dwmc")=strdwmc
				rs("mh")=strmh
				rs("diskid")=strdiskid
				if strbz<>"" then rs("bz")=strbz
				rs("cpsj")=dtcpsj
				rs.update
				rs.close
				Call JsAlert("��ˮ�� ��" & strlsh & "�� ģ�ߴ浵��Ϣ���ĳɹ���","docbak_search.asp")
			end if
		case else
			Call JsAlert("����Ҫ��ը�ˣ�","index.asp")
	end select
%>