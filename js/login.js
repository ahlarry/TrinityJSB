function login_true()
{
  if (document.frm_login.userName.value=="")
	  { alert("请输入您的 用户名称 ！"); document.frm_login.userName.focus(); return false;}
  if (document.frm_login.userPwd.value=="")
	  { alert("请输入您的 登陆密码 ！"); document.frm_login.userPwd.focus(); return false; }
}