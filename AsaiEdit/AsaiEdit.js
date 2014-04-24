var Asai_edit;	
var Asai_RangeType;
var Asai_selection;
var Asai_filterScript = false;
var Asai_charset="gb2312";
var Asai_bLoad=false;
var Asai_pureText=true;
var Asai_bTextMode=1;
var colour;

document.write("<style>#PostiFrame{border:1px solid #"+sy+";padding:3px;font-size:12px;}.pm{padding:0px;margin:0px;}.stt{position:absolute;background-color:#"+by+";border:#"+sy+" 1px solid;font-size:12px;padding:5px;FILTER: alpha(opacity=88)progid:DXImageTransform.Microsoft.Shadow(Color=#"+sy+",Direction=120,strength=4);}.stt a{cursor:pointer;}.as1{margin:2px;border:1px solid #"+sy+";background-color:#"+qy+";font-size:12px;cursor:pointer;}.as2{margin:2px;border:1px solid #"+by+";background-color:#"+by+";cursor:pointer;}.Ab{margin:2px;float:left;width:66px;height:22px;line-height:22px;text-align:center;font-size:13px;cursor:pointer;border:1px solid #"+sy+";background-color: #"+qy+";}.Abc{margin:2px;float:left;width:66px;height:22px;line-height:22px;text-align:center;font-size:13px;cursor:pointer;border:1px solid #"+sy+";background-color: #"+qy+";}.Abb{margin:2px;float:left;width:130px;height:22px;line-height:22px;text-align:center;font-size:13px;cursor:pointer;border:1px solid #"+sy+";background-color: #"+qy+";}.Aba{margin:2px;float:left;width:66px;height:22px;line-height:22px;text-align:center;font-size:13px;font-weight:bold;cursor:pointer;border:1px solid #"+sy+";background-color: #"+by+";}#Ar{border-top:1px solid #"+sy+";border-left:1px solid #"+sy+";border-right:1px solid #"+sy+";padding:2px;}#Ac{border:1px solid #"+sy+";padding:2px;}#At{border-bottom:1px solid #"+sy+";border-left:1px solid #"+sy+";border-right:1px solid #"+sy+";padding:2px;height:27px;}.Aj{float:right;margin:2px;width:32px;height:22px;line-height:22px;text-align:center;cursor:pointer;border:1px solid #"+sy+";background-color: #"+qy+";}</style>");

//HtmPop弹出窗口编辑 begin
//字体颜色
function Asai_foreColor()
{
	if (!Asai_validateMode()) return;
	if (Asai_bIsIE5){
		var arr = showModalDialog(""+Sm+""+Se+"HtmPop/selcolor.htm", "", "dialogWidth:280px; dialogHeight:300px; status:0; help:0");
		if (arr != null) FormatText('forecolor', arr);
		else IframeID.focus();
	}else
		{
		FormatText('forecolor', '');
		}
}

//字体背景
function Asai_backColor()
{
	if (!Asai_validateMode()) return;
	if (Asai_bIsIE5)
	{
		var arr = showModalDialog(""+Sm+""+Se+"HtmPop/selcolor.htm", "", "dialogWidth:280px; dialogHeight:300px; status:0; help:0");
		if (arr != null) FormatText('backcolor', arr);
		else IframeID.focus();
	}else
		{
		FormatText('backcolor', '');
		}
}

//上传文件
function UpLoad()
{
  var arr = showModalDialog(""+Sm+""+Se+"HtmPop/FileUp.htm", window, "dialogWidth:22em; dialogHeight:10em; status:0; help:0");
  if (arr != null){  
  var str1;
  var ss;
  ss=arr.split("*")
  a=ss[0];
  b=ss[1];
  c=ss[2];
  str1=""  
  if (c=="jpg" || c=="gif" || c=="png"){
  	str1=str1+"<p align=center><img border=0 src='"+b+"' alt='"+a+"'"
  	str1=str1+" onload='javascript:if(this.width>screen.width-333)this.width=screen.width-333'"
  	str1=str1+" ></A></P>"
  }else
  {
  	str1="<p align=left><img style='width:32px;height:32px' src='AsaiEdit/SysImg/"+c+"48.jpg'><A href='"+b+"'><FONT size=4>"+a+"</FONT></A></P>";
  }
IframeID.document.body.innerHTML+=str1;
   }
  IframeID.focus();
}

//插入表情
function Asai_qq()
{
	var arr = showModalDialog(""+Sm+""+Se+"HtmPop/face.htm", "", "dialogWidth:520px; dialogHeight:480px; status:0; help:0");
	
	if (arr != null)
	{
		Asai_InsertSymbol(arr);
		IframeID.focus();
	}
	else IframeID.focus();
}

//插入特殊符号
function insertSpecialChar()
{
	var arr = showModalDialog(""+Sm+""+Se+"HtmPop/specialchar.htm", "","dialogWidth:555px; dialogHeight:432px; status:0; help:0");
	if (arr != null) Asai_InsertSymbol(arr);
	IframeID.focus() ;
}
function doZoom( sizeCombo ) 
{
	if (sizeCombo.value != null || sizeCombo.value != "")
	if (Asai_bIsIE5){
	var z = IframeID.document.body.runtimeStyle;}
	else{
	var z = IframeID.document.body.style;
	}
	z.zoom = sizeCombo.value + "%" ;
}

//打开其他编辑器
function AsaiPopUp(style, form, field, width, height) {
	window.open(""+Sm+""+Se+"HtmPop/AsaiEditor/popup.htm?style="+style+"&form="+form+"&field="+field, "", "width="+width+",height="+height+",toolbar=no,location=no,status=no,menubar=no,scrollbars=no,resizable=no");
}
//HtmPop弹出窗口编辑 end

//特殊样式编辑 Begin
//引用
function Asai_quote()
{
	Asai_specialtype("<div style='margin:18px;border:1px dotted #CCCCCC;padding:5px;background:#EAF7EC;font-size:12px;font-family:Tahoma;line-height:normal;' title='提示：这是一段引用的内容！'>","</div>");
}
//特殊样式编辑 end


//直接用到的过程 begin
//清空
function ClearReset()
{IframeID.document.body.innerHTML='';
IframeID.focus();}

//增加、减小编辑器大小
function Asai_Size(num)
{
	var obj=document.getElementById("Asai_Composition");
	if (parseInt(obj.offsetHeight)+num>=100) {
		
		obj.height = (parseInt(obj.offsetHeight) + num);
	}
	if (num>0)
	{
		obj.width="100%";
	}
}

//编辑器的设计、源码模式切换
function Asai_setMode(n)
{
	Asai_setStyle();
	var cont;
	var ar=document.getElementById("ar");
	switch (n){
		case 1:
				ar.style.display="";
				if (document.getElementById("Asai_TabHtml").className=="Aba"){
					if (Asai_bIsIE5){
						cont=IframeID.document.body.innerText;
						cont=Asai_correctUrl(cont);
						if (Asai_filterScript)
						cont=Asai_FilterScript(cont);
						IframeID.document.body.innerHTML="<a>　</a>"+cont;
						
					}else{
						var html = IframeID.document.body.ownerDocument.createRange();
						html.selectNodeContents(IframeID.document.body);
						IframeID.document.body.innerHTML = html.toString();
					}
				}
				break;
		case 2:
					ar.style.display="none";	
					Asai_cleanHtml();
					cont=IframeID.document.body.innerHTML;
					cont=Asai_rCode(IframeID.document.body.innerHTML,"<a>　</a>","");
					cont=Asai_correctUrl(cont);
					if (Asai_filterScript){cont=Asai_FilterScript(cont);}
					if (Asai_bIsIE5){					
						IframeID.document.body.innerText=cont;
					}else{								
						var html=document.createTextNode(cont);
						IframeID.document.body.innerHTML = "";
						IframeID.document.body.appendChild(html);
					}
				break;
	}
	Asai_setTab(n);
	Asai_bTextMode=n
}

//格式化Word、Excel
function ClearWord()
{
	var htt;
	htt=IframeID.document.body.innerHTML;
	htt = htt.replace(/<\/?SPAN[^>]*>/gi, "" );
    // 清理类属性
    htt = htt.replace(/<(\w[^>]*) class=([^ |>]*)([^>]*)/gi, "<$1$3") ;
    // 清理样式属性
    htt = htt.replace(/<(\w[^>]*) style="([^"]*)"([^>]*)/gi, "<$1$3") ;
	// 清理空白样式
 	htt = htt.replace( /\s*style="\s*"/gi, '' ) ;
 	htt = htt.replace( /<SPAN\s*[^>]*>\s*&nbsp;\s*<\/SPAN>/gi, '&nbsp;' ) ;
 	htt = htt.replace( /<SPAN\s*[^>]*><\/SPAN>/gi, '' ) ;
    // 清理语言属性
    htt = htt.replace(/<(\w[^>]*) lang=([^ |>]*)([^>]*)/gi, "<$1$3") ;
 	// 清理字体属性
   	htt = htt.replace( /<FONT[^>]*>(.*?)<\/FONT>/gi, "$1" ) ;    
   	htt = htt.replace( /<FONT[^>]*>(.*?)<\/FONT>/gi, "$1" ) ;    
    // 清理XML元素和声明
    htt = htt.replace(/<\\?\?xml[^>]*>/gi, "") ;
    // 清理XML命名空间声明标签: <o:p></o:p>
    htt = htt.replace(/<\/?\w+:[^>]*>/gi, "") ;
    // 清理空白标签
 	htt = htt.replace( /<([^\s>]+)[^>]*>\s*<\/\1>/g, '' ) ;
 	htt = htt.replace( /<([^\s>]+)[^>]*>\s*<\/\1>/g, '' ) ;
    // 清理空格
    htt= htt.replace(/&nbsp;/gi, "" );
    // 转换 <P> 为 <DIV>
    var re = new RegExp("(<P)([^>]*>.*?)(<\/P>)","gi") ;        
    // 修正IE 5.0 版本的显示错误
    htt = htt.replace( re, "<div$2</div>" ) ;
    IframeID.document.body.innerHTML = htt;     
}

//选择后的样式变化
function Asai_setTab(n)
{
	
	var mhtml=document.getElementById("Asai_TabHtml");
	var mdesign=document.getElementById("Asai_TabDesign");
	if (n==1)
	{
		mhtml.className="Ab";
		mdesign.className="Aba";		
	}
	else if (n==2)
	{
		mhtml.className="Aba";
		mdesign.className="Ab";
	}
	else if (n==3)
	{
		mhtml.className="Ab";
		mdesign.className="Ab";
	}
}

//基本的格式化文本命令
function FormatText(command, option)
{
var codewrite
if (Asai_bIsIE5){
		if (option=="removeFormat"){
		command=option;
		option=null;}
		IframeID.focus();
	  	IframeID.document.execCommand(command, false, option);
		Asai_pureText = false;
		IframeID.focus();
		
}else{
		if ((command == 'forecolor') || (command == 'backcolor')) {
			parent.command = command;
			buttonElement = document.getElementById(command);
			IframeID.focus();
			document.getElementById("colourPalette").style.left = getOffsetLeft(buttonElement) + "px";
			document.getElementById("colourPalette").style.top = (getOffsetTop(buttonElement) + buttonElement.offsetHeight) + "px";
		
			if (document.getElementById("colourPalette").style.visibility=="hidden")
				{document.getElementById("colourPalette").style.visibility="visible";
			}else {
				document.getElementById("colourPalette").style.visibility="hidden";
			}
		
			
			var sel = IframeID.document.selection; 
			if (sel != null) {
				colour = sel.createRange();
			}
		}
		else{
		IframeID.focus();
	  	IframeID.document.execCommand(command, false, option);
		Asai_pureText = false;
		IframeID.focus();
		}
	}
}

//清理编辑器内代码
function Asai_correctUrl(cont)
{
	var regExp;
	var url=location.href.substring(0,location.href.lastIndexOf("/")+1);
	cont=Asai_rCode(cont,location.href+"#","#");
	cont=Asai_rCode(cont,url,"");
	cont=Asai_rCode(cont,"<a>　</a>","");
	return cont;
}

function Asai_cleanHtml()
{
	if (Asai_bIsIE5){
	var fonts = IframeID.document.body.all.tags("FONT");
	}else{
	var fonts = IframeID.document.getElementsByTagName("FONT");
	}
	var curr;
	for (var i = fonts.length - 1; i >= 0; i--) {
		curr = fonts[i];
		if (curr.style.backgroundColor == "#ffffff") curr.outerHTML = curr.innerHTML;
	}
}

function Asai_InsertSymbol(str1)
{
	IframeID.focus();
	if (Asai_bIsIE5) Asai_selectRange();
	Asai_edit.pasteHTML(str1);
}

function Asai_selectRange(){
	Asai_selection =	IframeID.document.selection;
	Asai_edit		=	Asai_selection.createRange();
	Asai_RangeType =	Asai_selection.type;
}

//编辑器里面插入内容用
function Asai_specialtype(Mark1, Mark2){
	var strHTML;
	if (Asai_bIsIE5){
		Asai_selectRange();
		if (Asai_RangeType == "Text"){
			if (Mark2==null)
			{
				strHTML = "<" + Mark1 + ">" + Asai_edit.htmlText + "</" + Mark1 + ">"; 
			}else{
				strHTML = Mark1 + Asai_edit.htmlText +  Mark2; 
			}
			Asai_edit.pasteHTML(strHTML);
			IframeID.focus();
			Asai_edit.select();
		}
		else{window.alert("请选择相应内容！")}	
	}
	else{
		if (Mark2==null)
		{
		strHTML	=	"<" + Mark1 + ">" + IframeID.document.body.innerHTML + "</" + Mark1 + ">"; 
		}else{
		strHTML = Mark1 + IframeID.document.body.innerHTML +  Mark2; 
		}
		IframeID.document.body.innerHTML=strHTML
		IframeID.focus();
	}
}

function Asai_getText()
{
	if (Asai_bTextMode==2)
		return IframeID.document.body.innerText;
	else
	{
		Asai_cleanHtml();
		return IframeID.document.body.innerHTML;
	}
}

function getOffsetTop(elm) {
	var mOffsetTop = elm.offsetTop;
	var mOffsetParent = elm.offsetParent;
	while(mOffsetParent){
		mOffsetTop += mOffsetParent.offsetTop;
		mOffsetParent = mOffsetParent.offsetParent;
	}
	return mOffsetTop;
}

function getOffsetLeft(elm) {
	var mOffsetLeft = elm.offsetLeft;
	var mOffsetParent = elm.offsetParent;
	while(mOffsetParent) {
		mOffsetLeft += mOffsetParent.offsetLeft;
		mOffsetParent = mOffsetParent.offsetParent;
	}
	return mOffsetLeft;
}

function rand() {
	return parseInt((1000)*Math.random()+1);
}

//检验连接地址是不是已经输入了
function Asai_UserDialog(what)
{
	if (!Asai_validateMode()) return;
	IframeID.focus();
	if (what == "CreateLink") {
		if (Asai_bIsNC)
		{
			insertLink = prompt("请填写超级链接地址信息：", "http://");			
			if ((insertLink != null) && (insertLink != "") && (insertLink != "undefined")) {
			IframeID.document.execCommand('CreateLink', false, insertLink);
			}else{
			IframeID.document.execCommand('unlink', false, null);
			}
		}
		else {
			IframeID.document.execCommand(what, true, null);
		}
	}
	
	if(what == "InsertImage"){
		imagePath = prompt('请填写图片链接地址信息：', 'http://');			
		if ((imagePath != null) && (imagePath != "")) {
			IframeID.document.execCommand('InsertImage', false, imagePath);
		}
		IframeID.document.body.innerHTML = (IframeID.document.body.innerHTML).replace("src=\"file://","src=\"");
	}
	Asai_pureText = false;
	IframeID.focus();
}

function Asai_GetRangeReference(editor)
{
	editor.focus();
	var objReference = null;
	var RangeType = editor.document.selection.type;
	var selectedRange = editor.document.selection.createRange();
	
	switch(RangeType)
	{
	case 'Control' :
		if (selectedRange.length > 0 ) 
		{
			objReference = selectedRange.item(0);
		}
	break;
	case 'None' :
		objReference = selectedRange.parentElement();
		break;
	case 'Text' :
		objReference = selectedRange.parentElement();
		break;
	}
	return objReference
}

function Asai_CheckTag(item,tagName)
{
	if (item.tagName.search(tagName)!= -1)
	{
		return item;
	}
	if (item.tagName == 'BODY')
	{
		return false;
	}
	item=item.parentElement;
	return Asai_CheckTag(item,tagName);
}

function Asai_FilterScript(content)
{
	content = Asai_rCode(content, 'javascript:', '<b>javascript</b> :');
	content = content.replace(RegExp, "<div style='margin:18px;border:1px dotted #CCCCCC;padding:5px;background:#FDFDDF;font-size:12px;font-family:Tahoma;line-height:normal;cursor:pointer;' title='点击运行该代码，观看即时演示！' onclick=\"preWin=window.open('','','');preWin.document.open();preWin.document.write(this.innerText);preWin.document.close();\">&lt;!-- Script 代码开始 --&gt;<br>$1<br>&lt;!-- Script 代码结束 --&gt;</div>");
	RegExp = /<P>&nbsp;<\/P>/gi;
	content = content.replace(RegExp, "");
	return content;
}

function Asai_rCode(s,a,b,i)
{
	a = a.replace("?","\\?");
	if (i==null)
	{
		var r = new RegExp(a,"gi");
	}else if (i) {
		var r = new RegExp(a,"g");
	}
	else{
		var r = new RegExp(a,"gi");
	}
	return s.replace(r,b); 
}

//编辑器内容check和提交代码 beging
//判断提交时候的状态
function Asai_validateMode()
{
	if (Asai_bTextMode!=2) return true;
	alert("请取消查看“源码”状态，调至“设计”状态再行提交，谢谢!");
	IframeID.focus();
	return false;
}

//编辑器区域整体样式表现
function Asai_InitDocument(hiddenid, charset)
{	
	if (charset!=null)
	Asai_charset=charset;
		var Asai_bodyTag="</head><BODY bgcolor=\"#ffffff\" style='font-size:9pt;'>";
	if (navigator.appVersion.indexOf("MSIE 6.0",0)==-1){
	IframeID.document.designMode="On"
	}
	IframeID.document.open();
	IframeID.document.write ('<html><head>');
	if (Asai_bIsIE5){
	}
	IframeID.document.write(Asai_bodyTag);
	IframeID.document.write("</body>");
	IframeID.document.write("</html>");
	IframeID.document.close();
	IframeID.document.body.contentEditable = "True";
	IframeID.document.charset=Asai_charset;
	Asai_bLoad=true;
	Asai_setStyle();
}
//编辑器区域滚动条样式
function Asai_setStyle()
{
var bs = IframeID.document.body.style;
if (Asai_bTextMode==2) {
bs.fontFamily="Arial";
bs.fontSize="12px";
}else{
bs.fontFamily="Arial";
bs.fontSize="12px";
}
bs.scrollbarShadowColor= '#'+sy+'';//立体滚动条阴影的颜色
bs.scrollbar3dLightColor= '#'+by+'';//滚动条亮边的颜色
bs.scrollbarArrowColor= '#'+sy+'';//上下按钮上三角箭头的颜色
bs.scrollbarBaseColor= '#'+by+'';//滚动条的基本颜色
bs.scrollbarDarkShadowColor= '#'+by+'';//滚动条强阴影的颜色
bs.scrollbarFaceColor= '#'+by+'';//滚动条凸出部分的颜色
bs.scrollbarHighlightColor= '#'+sy+'';//滚动条空白部分的颜色
bs.scrollbarTrackColor= '#'+qy+'';//滚动条的背景颜色
bs.border='0';
}

//关闭弹出菜单
function disall(){ 
if(sty.style.display==""){sty.style.display="none";} 
if(sfn.style.display==""){sfn.style.display="none";} 
if(sfd.style.display==""){sfd.style.display="none";} 
if(sfh.style.display==""){sfh.style.display="none";} 
}
function disty(){ 
if(sty.style.display=="none"){ 
sty.style.display="" ; 
}else{ 
sty.style.display="none" ; 
} 
} 
function disfn(){ 
if(sfn.style.display=="none"){ 
sfn.style.display="" ; 
}else{ 
sfn.style.display="none" ; 
} 
} 
function disfd(){ 
if(sfd.style.display=="none"){ 
sfd.style.display="" ; 
}else{ 
sfd.style.display="none" ; 
} 
} 
function disfh(){ 
if(sfh.style.display=="none"){ 
sfh.style.display="" ; 
}else{ 
sfh.style.display="none" ; 
} 
} 

function submits(){
var html;
html=Asai_getText();
html=Asai_rCode(html,"<a>　</a>","");
fdocument.all(""+an+"").value=html;
}

function ax(){
content=document.all(""+an+"").value;
IframeID.document.body.innerHTML=content;
document.all(""+an+"").value="";
}

function chk(){
if (!Asai_validateMode()){return false;}
document.all(""+an+"").value=IframeID.document.body.innerHTML;
if(IframeID.document.body.innerHTML==""){
alert("内容不能为空");
IframeID.document.body.focus();
return false;}
return true;}

//去掉所有HTML符号检验是否为空值
function Asai_ChekEmptyCode(html)
{
	html = html.replace(/\<[^>]*>/g,"");        
	html = html.replace(/&nbsp;/gi, "");
	html = html.replace(/o:/gi, "");
	html = html.replace(/\s/gi, "");
	return html;
}

function ctlent(eventobject)
{
	if(event.ctrlKey && event.keyCode==13)
	{
		this.document.formasai.submit();
	}
}
function getHTML()
{
	var html;
	if (!Asai_bTextMode) 
	{
	html = IframeID.document.body.innerHTML
	}
	else
	{
	html = IframeID.document.body.innerText
	}
	return html;
}
//编辑器内容check和提交代码 end
