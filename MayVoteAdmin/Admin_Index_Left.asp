<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�����Ŀ - MayVote��̨����</title>
<link href="Images/style.css" rel="stylesheet" type="text/css">
<SCRIPT language=javascript1.2>
function showsubmenu(sid)
{
whichEl = eval("submenu" + sid);
if (whichEl.style.display == "none")
{
eval("submenu" + sid + ".style.display='';");
}
else
{
eval("submenu" + sid + ".style.display='none';");
}
}
</SCRIPT>
</head>

<body  topmargin="0">
<br>
<table width=160 align=center cellpadding=0 cellspacing=0>
  <tr> 
    <td height=25 align="center" valign="middle" class=menu_title><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td height="25" align="center" class="meun_title3"><font color="#FFFFFF"><a href='Admin_Index_Main.asp' target=main><b><font color="#FFFFFF">������ҳ</font></b></a> 
            <font color="#FFFFFF">|</font> <a href='Admin_Login.asp?Action=Logout' target='_top'><b><font color="#FFFFFF">�˳�</font></b></a></font></td>
        </tr>
      </table>
      
    </td>
  </tr>
  <tr> 
    <td>
        <table width=100% border="0" align=center cellpadding=4 cellspacing=0>
        <tr> 
            <td class="sec_menu"><table width="150" border="0" align="center" cellpadding="4" cellspacing="0">
              <tr> 
                <td>�û�����<% = Session("UserName")%></td>
              </tr>
              <tr> 
                <td>�û��飺<%If Session("System") = 1 Then 
Response.Write "��������Ա"
Else
Response.Write"��ͨ����Ա"
End If%></td>
              </tr>
            </table> </td>
          </tr>
        </table>
      
    </td>
  </tr>
</table>
<br>
<table width=160 align=center cellpadding=0 cellspacing=0>
  <tr> 
    <td height=25 align="center" valign="middle" class=menu_title id=menuTitle100 style="cursor:hand;" onclick="showsubmenu(100)" onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title';><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td height="25" class="meun_title3"><font color="#FFFFFF"><strong>��<font color="#FFFFFF">�������</font></strong></font></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td id='submenu100' style="display:none"> <div class=sec_menu style="width:160"> 
        
        <table width="150" border="0" align="center" cellpadding="4" cellspacing="0">
          <tr> 
            <td><img src="Images/bullet.gif" width="7" height="7">&nbsp;<a href="Admin_Setting.asp" target="main">��������</a></td>
          </tr>
          <tr> 
            <td><img src="Images/bullet.gif" width="7" height="7">&nbsp;<a href="Admin_Users.asp" target="main">�û�����</a></td>
          </tr>
          <tr>
            <td><img src="Images/bullet.gif" width="7" height="7">&nbsp;<a href="Admin_Users.asp?Action=EditPassWord" target="main">�޸ĸ�������</a></td>
          </tr>
        </table>
      </div>
      <div  style="width:160"> 
        <table cellpadding=0 cellspacing=0 align=center width=130>
          <tr> 
            <td height=24></td>
          </tr>
        </table>
      </div></td>
  </tr>
</table>
<table width=160 align=center cellpadding=0 cellspacing=0>
  <tr>
    <td height=25 align="center" valign="middle" class=menu_title id=menuTitle101 style="cursor:hand;" onclick="showsubmenu(101)" onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title';><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td height="25" class="meun_title3"><strong>��<font color="#FFFFFF">ͶƱ����</font></strong></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td id='submenu101' style="display:none"><div class=sec_menu style="width:160">
      <table width=150 border="0" align=center cellpadding=4 cellspacing=0 >
        <tr>
          <td><img src="Images/bullet.gif" width="7" height="7">&nbsp;<a href="Admin_AddVote.asp" target="main">�����Ŀ</a></td>
        </tr>
        <tr>
          <td><img src="Images/bullet.gif" width="7" height="7">&nbsp;<a href="Admin_Edit.asp" target="main">�༭��Ŀ</a></td>
        </tr>
        <tr>
          <td><img src="Images/bullet.gif" width="7" height="7">&nbsp;<a href="Admin_JS_Guide.asp" target="main">JS ������</a></td>
        </tr>
        <tr>
          <td><img src="Images/bullet.gif" width="7" height="7">&nbsp;<a href="Admin_VoteTop.asp" target="main">ͶƱ���а�</a></td>
        </tr>
      </table>
    </div>
        <div  style="width:160">
          <table cellpadding=0 cellspacing=0 align=center width=130>
            <tr>
              <td height=24></td>
            </tr>
          </table>
        </div></td>
  </tr>
</table>
<table width=160 align=center cellpadding=0 cellspacing=0>
  <tr> 
    <td height=25 align="center" valign="middle" class=menu_title id=menuTitle102 style="cursor:hand;" onclick="showsubmenu(102)" onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title';><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td height="25" class="meun_title3"><strong>��<font color="#FFFFFF">���ݿ����</font></strong></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td id='submenu102' style="display:none"> <div class=sec_menu style="width:160">
        
        <table cellpadding=4 cellspacing=0 align=center width=150>
          <tr> 
            <td><img src="Images/bullet.gif" width="7" height="7">&nbsp;<a href="Admin_Database.asp?Action=BackUpData" target="main">�������ݿ�</a></td>
          </tr>
          <tr> 
            <td><img src="Images/bullet.gif" width="7" height="7">&nbsp;<a href="Admin_Database.asp?Action=RestoreData" target="main">�ָ����ݿ�</a></td>
          </tr>
          <tr> 
            <td><img src="Images/bullet.gif" width="7" height="7">&nbsp;<a href="Admin_Database.asp?Action=CompactData" target="main">ѹ�����ݿ�</a></td>
          </tr>
        </table>
      </div>
      <div  style="width:160"> 
        <table cellpadding=0 cellspacing=0 align=center width=130>
          <tr> 
            <td height=24></td>
          </tr>
        </table>
      </div></td>
  </tr>
</table>
<table width=160 align=center cellpadding=0 cellspacing=0>
  <tr> 
    <td height=25 align="center" valign="middle" class=menu_title><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td height="25" class="meun_title3">��<font color="#FFFFFF"><strong>��Ȩ��Ϣ</strong></font></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td> <table width=100% border="0" align=center cellpadding=4 cellspacing=0>
        <tr> 
          <td class="sec_menu"><table width=130 align=center cellpadding=0 cellspacing=0>
              <tr> 
                <td height=70> ��Ȩ���У�&nbsp;���Ѽ���ģ������<br>
                </td>
              </tr>
            </table></td>
        </tr>
      </table>
        <table cellpadding=0 cellspacing=0 align=center width=130>
          <tr> 
            <td height=24></td>
          </tr>
        </table>
      </div></td>
  </tr>
</table>
</body>
</html>