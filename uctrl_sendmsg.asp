<!--#include file="include/conn.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->
<body onLoad="return ax();">
<%
Call ChkPageAble(0)
CurPage="�û����� �� ����ϵͳ����Ϣ"
strPage="uctrl"
'Call FileInc(0, "js/login.js")
xjweb.header()
Call TopTable()
Call Main()
Call BottomTable()
xjweb.footer()
closeObj()

Sub Main()
%>

<Table class=xtable cellspacing=0 cellpadding=0 width="<%=web_info(8)%>">
  <Tr>
    <Td class=ctd height=300><%Call uctrlSendMsg()%>
      <%Response.Write(XjLine(10,"100%",""))%>
    </Td>
  </Tr>
</Table>
<%
End Sub

Function uctrlSendMsg()
	Call TbTopic("����վ�ڶ���Ϣ")
	dim strincept, strtitle
	strincept=request("incept")
	strtitle=request("title")
%>
<table border="0" cellpadding="3" cellspacing="0" class=xtable width="95%" align="center">
  <form action="msg_indb.asp?action=send" method="post"  onsubmit="return msgsendchk();">
    <tr>
      <td class=rtd width="10%">�ռ���</td>
      <td class=ltd width="*"><span id="span_incept" name="span_incept">
        <%if strincept<>"" then%>
        <%=strincept%>
        <%else%>
        ��ѡ���ռ���
        <%end if%>
        </span></td>
    </tr>
    <tr>
      <td class=rtd>Ⱥ���б�</td>
      <td class=ltd><input type=radio id=msgqf name=msgqf value=all class=radio onclick="checkqf();">
        ȫ&nbsp;&nbsp;&nbsp;��
        <input type=radio id=radio name=msgqf value=zy class=radio onclick="checkqf();" />
        ������Ա
        <input type=radio id=msgqf name=msgqf value=zz class=radio onclick="checkqf();">
        �����鳤
        <input type=radio id=msgqf name=msgqf value=jl class=radio onclick="checkqf();">
        ���о���
        <input type=radio id=msgqf name=msgqf value=no class=radio onclick="checkqf();">
        ��ѡȺ��<br />
        <input type=checkbox id=qfzy1 name=qfzy1 value=xz1 onclick="checkxz();">
        ��&nbsp;һ&nbsp;��&nbsp;&nbsp;
        <input type=checkbox id=qfzy2 name=qfzy2 value=xz2 onclick="checkxz();">
        ��&nbsp;��&nbsp;��&nbsp;&nbsp;
        <input type=checkbox id=qfzy3 name=qfzy3 value=xz3 onclick="checkxz();">
        ��&nbsp;��&nbsp;��&nbsp;&nbsp;
        <input type=checkbox id=qfzy4 name=qfzy4 value=xz4 onclick="checkxz();">
        ��&nbsp;��&nbsp;��&nbsp;&nbsp;
        <input type=checkbox id=qfzy5 name=qfzy5 value=xz5 onclick="checkxz();">
        ��&nbsp;��&nbsp;��&nbsp;&nbsp;
        <input type=checkbox id=qfzy0 name=qfzy0 value=xz0 onclick="checkxz();">
        ��&nbsp;��&nbsp;��
         </td>
    </tr>
    <tr>
      <td class=rtd>�ռ����б�</td>
      <td class=ltd><table border="0" width="100%" cellspacing="0" cellpadding="0">
          <tr>
            <%
				dim j
				j=1
				for i=0 to ubound(c_jsb)
					if j>10 then response.write("</tr><tr>") : j=1
			%>
            <td><input type=checkbox id=user<%=i%> name=user<%=i%> value=<%=c_jsb(i)%> class=radio onclick="changesenduser();">
              <label for=user<%=i%>><%=c_jsb(i)%></label>
            </td>
            <%
					j=j+1
				next
			%>
          </tr>
        </table></td>
    </tr>

    <tr>
      <td class=rtd >����</td>
      <td class=ltd><input type="text" name="title" size="70" maxlength=100 tabindex=1 value=<%=strtitle%>>
        (����100�ַ�)</td>
    </tr>
    <tr>
      <td class=rtd>����</td>
      <td class=ltd>
      	<SCRIPT language="javascript">
		var Sm="";
		var Se="AsaiEdit/";//�༭�����ڸ�Ŀ¼����
		var sy="9EC9EC";//����ɫ
		var qy="EDF6FF";//ǳ��ɫ
		var by="FFFFFF";//����ɫ
		var an="content";//������
	 	</script>
		<SCRIPT language="JavaScript" src="AsaiEdit/AsaiEdit.js"></SCRIPT>
		<SCRIPT language="JavaScript" src="AsaiEdit/EditMenu.js"></SCRIPT>
      <textarea name="content" id="content" cols="88" rows="8" style="display:none;"></textarea></td>
    </tr>
    <tr>
      <td class=ctd colspan="2"><input type="submit" value=" ���� "></td>
    </tr>
    <input type="hidden" name="incept" value="<%=strincept%>">
  </form>
</table>
<script language="javascript">
		var allzy=new Array(<%=ubound(c_allzy)%>);
		var allzz=new Array(<%=ubound(c_allzz)%>);
		var alljl=new Array(<%=ubound(c_alljl)%>);
		var allxz0=new Array(<%=ubound(c_xz0)%>);
		var allxz1=new Array(<%=ubound(c_xz1)%>);
		var allxz2=new Array(<%=ubound(c_xz2)%>);
		var allxz3=new Array(<%=ubound(c_xz3)%>);
		var allxz4=new Array(<%=ubound(c_xz4)%>);
		var allxz5=new Array(<%=ubound(c_xz5)%>);

		var i,j;
		for (i = 0; i <<%=ubound(c_allzy)%>; i++)
		{
			<%for i=0 to ubound(c_allzy)%>
				allzy[<%=i%>] ="<%=c_allzy(i)%>";
			<%next%>
		}
		for (i = 0; i <<%=ubound(c_xz0)%>; i++)
		{
			<%for i=0 to ubound(c_xz0)%>
				allxz0[<%=i%>] ="<%=c_xz0(i)%>";
			<%next%>
		}
		for (i = 0; i <<%=ubound(c_xz1)%>; i++)
		{
			<%for i=0 to ubound(c_xz1)%>
				allxz1[<%=i%>] ="<%=c_xz1(i)%>";
			<%next%>
		}
		for (i = 0; i <<%=ubound(c_xz2)%>; i++)
		{
			<%for i=0 to ubound(c_xz2)%>
				allxz2[<%=i%>] ="<%=c_xz2(i)%>";
			<%next%>
		}
		for (i = 0; i <<%=ubound(c_xz3)%>; i++)
		{
			<%for i=0 to ubound(c_xz3)%>
				allxz3[<%=i%>] ="<%=c_xz3(i)%>";
			<%next%>
		}
		for (i = 0; i <<%=ubound(c_xz4)%>; i++)
		{
			<%for i=0 to ubound(c_xz4)%>
				allxz4[<%=i%>] ="<%=c_xz4(i)%>";
			<%next%>
		}
		for (i = 0; i <<%=ubound(c_xz5)%>; i++)
		{
			<%for i=0 to ubound(c_xz5)%>
				allxz5[<%=i%>] ="<%=c_xz5(i)%>";
			<%next%>
		}
		for (i = 0; i <<%=ubound(c_allzz)%>; i++)
		{
			<%for i=0 to ubound(c_allzz)%>
				allzz[<%=i%>] ="<%=c_allzz(i)%>";
			<%next%>
		}
		for (i = 0; i <<%=ubound(c_alljl)%>; i++)
		{
			<%for i=0 to ubound(c_alljl)%>
				alljl[<%=i%>] ="<%=c_alljl(i)%>";
			<%next%>
		}

		function checkqf()
		{
			var obj1 = document.all.msgqf;
			var str1;
			for(i=0;i < obj1.length; i ++)
			{
				if(obj1[i].checked)
					str1 = obj1[i].value;
			}
			switch(str1)
			{
				case "all":
					checkuser();
					break;
				case "zy":
					checkzy();
					break;
				case "zz":
					checkzz();
					break;
				case "jl":
					checkjl();
					break;
				case "no":
					donotqf();
					break;
			}
		}

		function donotqf()
		{
			for(j=0; j<=<%=ubound(c_jsb)%>; j++)
			{
				eval("document.all.user"+j+".checked=false");
			}
			for(i=0;i<6;i++) {
				eval("document.all.qfzy" + i +".checked=false");
			}
			changesenduser();
		}

		function checkuser()
		{
			for(j=0; j<=<%=ubound(c_jsb)%>; j++)
			{
				eval("document.all.user"+j+".checked=true");
			}
			changesenduser();
		}
		function checkzy()
		{
			for(j=0; j<=<%=ubound(c_jsb)%>; j++)
			{
				eval("document.all.user"+j+".checked=false");
			}
			for(i=0;i<6;i++) {
				eval("document.all.qfzy"+i+".checked=false");
			}
			for(j=0; j<=<%=ubound(c_jsb)%>; j++)
			{
				for(i=0; i<allzy.length; i++)
				{
					if(allzy[i]==eval("document.all.user"+j+".value"))
					{
						eval("document.all.user"+j+".checked=true");
					}
				}
			}
			changesenduser();
		}
		function checkzz()
		{
			for(j=0; j<=<%=ubound(c_jsb)%>; j++)
			{
				eval("document.all.user"+j+".checked=false");
			}
			for(i=0;i<6;i++) {
				eval("document.all.qfzy"+i+".checked=false");
			}
			for(j=0; j<=<%=ubound(c_jsb)%>; j++)
			{
				for(i=0; i<allzz.length; i++)
				{
					if(allzz[i]==eval("document.all.user"+j+".value"))
					{
						eval("document.all.user"+j+".checked=true");
					}
				}
			}
			changesenduser();
		}
		function checkjl()
		{
			for(i=0;i<6;i++) {
				eval("document.all.qfzy"+i+".checked=false");
			}
			for(j=0; j<=<%=ubound(c_jsb)%>; j++)
			{
				eval("document.all.user"+j+".checked=false");
			}

			for(j=0; j<=<%=ubound(c_jsb)%>; j++)
			{
				for(i=0; i<alljl.length; i++)
				{
					if(alljl[i]==eval("document.all.user"+j+".value"))
					{
						eval("document.all.user"+j+".checked=true");
					}
				}
			}
			changesenduser();
		}
		function checkxz()
		{
			document.all.msgqf[4].checked=true;
			var xz=new Array();
			for(i=0;i<6;i++) {
				var chkobj=eval("document.all.qfzy" + i);
				var b=eval("allxz"+i);
				if(chkobj.checked){
					xz=xz.concat(b);
				}
			}
			for(j=0; j<=<%=ubound(c_jsb)%>; j++)
			{
				eval("document.all.user"+j+".checked=false");
			}
			for(j=0; j<=<%=ubound(c_jsb)%>; j++)
			{
				for(i=0; i<xz.length; i++)
				{
					if(xz[i]==eval("document.all.user"+j+".value"))
					{
						eval("document.all.user"+j+".checked=true");
					}
				}
			}
			changesenduser();
		}

		function changesenduser()
		{
			var ii=0;
			var strtemp="";
			for(ii=0; ii<=<%=ubound(c_jsb)%>; ii++)
			{
				if(eval("document.all.user" + ii +".checked==true"))
					//alert(eval("document.all.user" + ii + ".value"));
				{
					if(strtemp!="")
						strtemp=strtemp + "|" + eval("document.all.user" + ii + ".value");
					else
						strtemp=eval("document.all.user" + ii + ".value");
				}
			}

			if(strtemp=="")
				document.all.span_incept.innerHTML="��ѡ�������";
			else
				document.all.span_incept.innerHTML=strtemp;
				document.all.incept.value=strtemp;
		}
		function msgsendchk()
		{
			if (document.all.incept.value=="")
				 {alert("��ѡ���ռ��� ��");  return false;}
			if(trim(document.all.title.value)=="")
				{alert("�������������(100������),�Ҳ���Ϊ���ַ� ��"); document.all.title.focus(); return false;}
//			if(trim(document.all.content.value)=="")
//				{alert("�������������,����Ϊ���ַ� ��"); document.all.content.focus(); return false;}

			if (!Asai_validateMode()){return false;}
			document.all(""+an+"").value=IframeID.document.body.innerHTML;
			if(IframeID.document.body.innerHTML==""){
				alert("���ݲ���Ϊ��");
				IframeID.document.body.focus();
				return false;
			}
			return true;
		}
		function lTrim(str)
		{
			if (str.charAt(0) == " ")
			{
			//����ִ���ߵ�һ���ַ�Ϊ�ո�
			str = str.slice(1);//���ո���ִ���ȥ��
			//��һ��Ҳ�ɸĳ� str = str.substring(1, str.length);
			str = lTrim(str); //�ݹ����
			}
			return str;
		}

		//ȥ���ִ��ұߵĿո�
		function rTrim(str)
		{
			var iLength;

			iLength = str.length;
			if (str.charAt(iLength - 1) == " ")
			{
			//����ִ��ұߵ�һ���ַ�Ϊ�ո�
			str = str.slice(0, iLength - 1);//���ո���ִ���ȥ��
			//��һ��Ҳ�ɸĳ� str = str.substring(0, iLength - 1);
			str = rTrim(str); //�ݹ����
			}
			return str;
		}

		//ȥ���ִ����ߵĿո�
		function trim(str)
		{
			return lTrim(rTrim(str));
		}
	</script>
<%
end function
%>
