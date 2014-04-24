<script language="javascript">
var Cookie = {
	setCookie	: function (sName, sValue,dExpires)
	{
		if(sName!=null&&sValue!=null)
		{
			if(dExpires==null)
			{
				document.cookie = sName + "=" + escape(sValue) + "; ";
			}
			else
			{
				try
				{
					//var date = new Date(dExpires.replace(/-/g, "\/"));
					document.cookie = sName + "=" + escape(sValue) + "; expires=" + dExpires.toGMTString();
				}
				catch(e)
				{}

			}
		}
	},
	getCookie	: function(sName)
	{
		if(sName!=null)
		{
			var aCookie = document.cookie.split("; ");
			for (var i=0; i < aCookie.length; i++)
			{
				var aCrumb = aCookie[i].split("=");
				if (sName == aCrumb[0])
				{
					return unescape(aCrumb[1]);
				}
			}
		}
		else
		{
			return null;
		}
	},
	deleteCookie	: function(sName)
	{
		if(sName!=null)
		{
			document.cookie = sName + "=" + "; expires=Fri, 31 Dec 1900 23:59:59 GMT;";
		}
	}
};

var ans=false;
function getUserSelect(){
	ans = confirm("是否批量存盘吗?\n确定:批量保存\n取消:单个保存"); 
	if (ans==true){
	   // alert("按订单号批量存盘!")
		Cookie.setCookie("useroperation","batch"); 
  }
  else
  {
       // alert("按流水号单个存盘!")
		Cookie.setCookie("useroperation","single"); 
	 
  }
}
 
 
//-->
</script>