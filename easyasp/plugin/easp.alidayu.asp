<%
'#################################################################################
'## easp.alidayu.asp
'## ------------------------------------------------------------------------------
'##  Feature      : EasyASP 阿里大鱼短信发送插件
'##  Version      : v1.0
'##  For EasyASP  :  3.0+
'##  Author       : Jam(2239559[at]qq.com)
'##  Update Date  : 2016年7月1日 星期五
'##  Description  : 
'##       此插件用于阿里大鱼短信发送，具体API接口信息见https://api.alidayu.com/doc2/apiDetail?apiId=25450
'##       使用方法如下：
'##         Easp("alidayu")			- 根据属性值发送短信，成功返回True，错误返回错误信息
'##       在使用方法前必须设置如下属性值：
'##         Easp("alidayu").AppKey = "12345678"    					- 阿里大鱼分配给你的AppKey
'##         Easp("alidayu").Secret = "1acde2335689748a377654503d"   - AppKey对应的Secret
'##         Easp("alidayu").SignName = "Jam轻博客" 					- 在阿里大鱼管理后台中通过审批的短信签名
'##         Easp("alidayu").RecNum = "13911111111" 					- 要接收短信的手机号码
'##         Easp("alidayu").TplCode = "SMS_11370053" 				- 在阿里大鱼管理后台中通过审批的短信模板ID
'##       以上是必要的属性，下面还有两个非必须属性：
'##         Easp("alidayu").Extend = "2356"     					- 公共回传参数
'##         Easp("alidayu").SMSParam = "{""number"":""1234""}" 		- 短信模板变量，传参规则{"key":"value"}
'#################################################################################
Class EasyASP_alidayu	
	Private s_author,s_version,a_list,b_list,s_method,s_v,s_format,s_simplify,s_secret,s_app_key,s_sign_method
	Private s_sms_type,s_sms_free_sign_name,s_rec_num,s_sms_template_code,s_extend,s_sms_param
	Private Sub Class_Initialize()
		s_author = "jam"
		s_version = "0.1"
		s_method = "alibaba.aliqin.fc.sms.num.send"		'API接口名称，这里是短信发送的接口
		s_v = "2.0"						'API协议版本
		s_format = "json"				'响应格式
		s_simplify = false				'是否采用精简JSON返回格式，仅当format=json时有效，默认值为：false
		s_sign_method = "md5"			'签名的摘要算法
		s_sms_type = "normal"			'短信类型，传入值请填写normal		
		s_secret = ""					'APP的Secret，加密时用
		s_app_key = ""					'应用的AppKey
		s_sms_free_sign_name = "" 		'短信签名，传入的短信签名必须是在阿里大鱼“管理中心-短信签名管理”中的可用签名
		s_rec_num = ""					'短信接收号码。支持单个或多个手机号码，传入号码为11位手机号码，不能加0或+86。群发短信需传入多个号码，以英文逗号分隔，一次调用最多传入200个号码。示例：18600000000,13911111111,13322222222
		s_sms_template_code = ""		'短信模板ID，传入的模板必须是在阿里大鱼“管理中心-短信模板管理”中的可用模板
		s_extend = ""					'公共回传参数，在“消息返回”中会透传回该参数；举例：用户可以传入自己下级的会员ID，在消息返回时，该会员ID会包含在内，用户可以根据该会员ID识别是哪位会员使用了你的应用
		s_sms_param = ""				'短信模板变量，传参规则{"key":"value"}，key的名字须和申请模板中的变量名一致，多个变量之间以逗号隔开。示例：针对模板“验证码${code}，您正在进行${product}身份验证，打死不要告诉别人哦！”，传参时需传入{"code":"1234","product":"alidayu"}
		Easp.Error("error-SMS-AppKey") = "AppKey属性为空。|无法读取AppKey属性，这是个必要属性。|请在代码中用属性 Easp(""alidayu"").AppKey 进行设置。"
		Easp.Error("error-SMS-Secret") = "Secret属性为空。|无法读取Secret属性，这是个必要属性。|请在代码中用属性 Easp(""alidayu"").Secret 进行设置。"
		Easp.Error("error-SMS-SignName") = "SignName属性为空。|无法读取短信签名SignName属性，这是个必要属性。|请在代码中用属性 Easp(""alidayu"").SignName 进行设置。"
		Easp.Error("error-SMS-RecNum") = "RecNum属性为空。|无法读取短信接收号码RecNum属性，这是个必要属性。|请在代码中用属性 Easp(""alidayu"").RecNum 进行设置。"
		Easp.Error("error-SMS-TplCode") = "TplCode属性为空。|无法读取短信模板IDTplCode属性，这是个必要属性。|请在代码中用属性 Easp(""alidayu"").TplCode 进行设置。"
	End Sub
	Private Sub Class_Terminate()
		
	End Sub	
	'应用的AppKey属性
	Public Property Let AppKey(ByVal String)
		s_app_key = String
	End Property
	'APP的Secret，加密时用
	Public Property Let Secret(ByVal String)
		s_secret = String
	End Property
	'短信签名
	Public Property Let SignName(ByVal String)
		s_sms_free_sign_name = String
	End Property
	'短信接收号码
	Public Property Let RecNum(ByVal String)
		s_rec_num = String
	End Property
	'短信模板ID
	Public Property Let TplCode(ByVal String)
		s_sms_template_code = String
	End Property
	'公共回传参数
	Public Property Let Extend(ByVal String)
		s_extend = String
	End Property
	'短信模板变量
	Public Property Let SMSParam(ByVal String)
		s_sms_param = String
	End Property
 
	'发送短信，成功返回True，错误返回错误信息
	Public Default Function SendSMS()
		If Easp.isN(s_app_key) Then :If Easp.Debug Then Easp.Error.Raise "error-SMS-AppKey":Exit Function:End If 
		If Easp.isN(s_secret) Then :If Easp.Debug Then Easp.Error.Raise "error-SMS-Secret":Exit Function:End If 
		If Easp.isN(s_sms_free_sign_name) Then :If Easp.Debug Then Easp.Error.Raise "error-SMS-SignName":Exit Function:End If 
		If Easp.isN(s_rec_num) Then :If Easp.Debug Then Easp.Error.Raise "error-SMS-RecNum":Exit Function:End If 
		If Easp.isN(s_sms_template_code) Then :If Easp.Debug Then Easp.Error.Raise "error-SMS-TplCode":Exit Function:End If 
		'生成参数的键数组，原始字典对象
		Dim a,b,d,k,u,l,j
		Set d = Server.CreateObject("Scripting.Dictionary")
		'公用参数
		d.Add "method",s_method
		d.Add "app_key",s_app_key
		d.Add "timestamp",Easp.Date.Format(Now(),"yyyy-mm-dd hh:mm:ss")
		d.Add "format",s_format
		d.Add "v",s_v
		d.Add "sms_type",s_sms_type
		d.Add "simplify",s_simplify
		d.Add "sign_method",s_sign_method
		'接口参数
		d.Add "sms_free_sign_name",s_sms_free_sign_name
		d.Add "rec_num",s_rec_num
		d.Add "sms_template_code",s_sms_template_code
		d.Add "extend",s_extend
		d.Add "sms_param",s_sms_param
		Sort d
		For Each k In d
			u = u&k&d(k)
			l = l & "&" & k & "=" & Server.URLEncode(d(k))
		Next 
		Set d = Nothing
		'将HASH数组去掉无用内容并加上Secret，然后MD5返回字符串
		u = UCase(Easp("md5utf8")(s_secret&u&s_secret))
		'加上sign签名参数
		l = "sign="&u&l
		Easp.Http.Data = l
		Set j = Easp.Json.Parse(Easp.Http.Get("http://gw.api.taobao.com/router/rest"))
		If j.Has("alibaba_aliqin_fc_sms_num_send_response") Then 
			SendSMS = True
		ElseIf j.Has("error_response") Then 
			SendSMS = j("error_response")("sub_msg")
		End If 
	End Function
	
	'字典按Key排序
	Private Function Sort(dict)
		Dim i,j, temp
		Dim keys,items
		Dim t
		Set t = Server.CreateObject("Scripting.Dictionary")
		keys = dict.Keys
		items = dict.Items
		For i = 0 To  dict.Count - 1
			t.Add keys(i),items(i)
		Next	
		For i = 0 To  dict.Count - 1
			For j = i + 1 To dict.Count - 1			
				If Easp.Str.Compare(keys(i), ">", keys(j)) Then
					temp = keys(i)
					keys(i) = keys(j)
					keys(j) = temp
				End If
			Next
		Next	
		dict.RemoveAll
		For i = 0 To UBound(keys)
			dict.Add keys(i), t.Item(keys(i))
		Next
		t.RemoveAll
		Set t = Nothing
	End Function
End Class 
%>