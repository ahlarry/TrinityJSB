<!--#include file="include/conn.asp"-->
<%
dim action
action=request("action")
select case action
	case "item"
		strSql="select * from [ims_inform] where not(inform_outdate) and datediff('d',inform_date,now)<7 order by id desc"
		'sql="select * from ims_inform where inform_outdate order by id desc"
		call xjweb.Exec("", -1)
		set rs=server.createobject("adodb.recordset")
		rs.open strSql, conn, 1, 3
		if rs.eof or rs.bof then
			call noinform()
		else
			call haveinform(rs)
		end if
		rs.close
	case else
end select

function noinform()
%>
	document.write('<font color=#aaaaaa>&nbsp;</font>');
<%
end function

function haveinform(rs)
%>
	document.write("<b>通知:</b> <span id='inform' style='width:150;height:10;color=green;font-weight:bold;text-align:left;'>请启用JS</span>");
	document.write("<script language=javascript>");
	document.write("var i = 0;");
	document.write("var j = 0;");
	document.write("var k = 0;");
	document.write("var str=new Array;");
	document.write("k=<%=datediff("d",rs("inform_date"),now)%>;")
	<%
	dim tempstr
	if len(rs("inform_topic"))>10 then tempstr=left(rs("inform_topic"),8) & "......" else tempstr=rs("inform_topic")
	%>
	document.write("str[0]='<a id=ainfo href=inform_dis.asp?id=<%=rs("id")%>><%=tempstr%></a>';")
	document.write("str[1]='发布日期:<%=rs("inform_date")%>';")
	document.write("str[2]='发 布 人:<%=rs("inform_lzr")%>';")
	document.write("window.inform.innerHTML = str[0];");
	document.write("window.ainfo.alt=str[1]+'<br>'+str[2];")
	document.write("function disinform()");
	document.write("{");
	document.write("j++;");
	document.write("if(j%2==1&&j<10){window.inform.style.filter='glow(Color=#ff0000,Strength=1)';}else{window.inform.style.filter='';}");
	document.write("if(j>20) j=0;");
	document.write("window.setTimeout('disinform()',eval(100*(k+1)));");
	document.write("}");
	document.write("window.setTimeout('disinform()',1);");
	document.write("</script>");
<%
end function
%>