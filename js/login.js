function login_true()
{
  if (document.frm_login.userName.value=="")
	  { alert("���������� �û����� ��"); document.frm_login.userName.focus(); return false;}
  if (document.frm_login.userPwd.value=="")
	  { alert("���������� ��½���� ��"); document.frm_login.userPwd.focus(); return false; }
}