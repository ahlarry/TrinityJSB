<!--#include file="include/conn.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->
<%
Call ChkPageAble(0)
CurPage="Ա������ �� Ա�������б�"
strPage="ygkp"
'Call FileInc(0, "js/login.js")
xjweb.header()
Call TopTable()
%>

<a href="#" title="������ʾ" alt="�⻹����ʾ">tipp</a>
<script Language="JavaScript">
//***********Ĭ�����ö���.*********************
tPopWait=50;//ͣ��tWait�������ʾ��ʾ��
tPopShow=5000;//��ʾtShow�����ر���ʾ
showPopStep=20;
popOpacity=99;

var tfontcolor="#000000";
var tbgcolor="#ecedef";
var tbordercolor="#666666";

//***************�ڲ���������*****************
sPop=null;
curShow=null;
tFadeOut=null;
tFadeIn=null;
tFadeWaiting=null;
document.write("<style type='text/css'id='defaultPopStyle'>");
document.write(".cPopText {background-repeat : repeat-y; background-color: " + tbgcolor + ";color:" + tfontcolor + "; border: 1px " + tbordercolor + " solid;font-color: font-size: 12px; padding-right: 4px; padding-left: 28px; height: 20px; padding-top: 2px; padding-bottom: 2px; filter: Alpha(Opacity=0);text-align:left;}");
document.write("</style>");
document.write("<div id='dypopLayer' style='position:absolute;z-index:1000;' class='cPopText'></div>");

function showPopupText(){
var o=event.srcElement;
MouseX=event.x;
MouseY=event.y;
if(o.alt!=null && o.alt!=""){o.dypop=o.alt;o.alt=""};
if(o.title!=null && o.title!=""){o.dypop=o.title;o.title=""};
if(o.dypop!=sPop) {
sPop=o.dypop;
clearTimeout(curShow);
clearTimeout(tFadeOut);
clearTimeout(tFadeIn);
clearTimeout(tFadeWaiting);
if(sPop==null || sPop=="") {
dypopLayer.innerHTML="";
dypopLayer.style.filter="Alpha()";
dypopLayer.filters.Alpha.opacity=0;
}
else {
if(o.dyclass!=null) popStyle=o.dyclass
else popStyle="cPopText";
curShow=setTimeout("showIt()",tPopWait);
}
}
}
function showIt(){
  dypopLayer.className=popStyle;
  sPop="<b>��ʾ:</b><br>" + sPop;
  dypopLayer.innerHTML=sPop;
  popsoffsetX = 10;   // ��������λ������������Ҳ�ľ��룻3-12 ����
  popsoffsetY = 15;  // ��������λ������·��ľ��룻3-12 ����
  popWidth=dypopLayer.clientWidth;
  popHeight=dypopLayer.clientHeight;
  popLeftAdjust=0;
//  if(MouseX+12+popWidth>document.body.clientWidth) { popLeftAdjust=-popWidth-24; }
  if(MouseX+popsoffsetX+popWidth>document.body.clientWidth) { popLeftAdjust=MouseX+popsoffsetX+popWidth-document.body.clientWidth; }
  else { popLeftAdjust=0; }
//  if(MouseY+12+popHeight>document.body.clientHeight) { popTopAdjust=-popHeight-24; }
  if(MouseY+popsoffsetY+popHeight>document.body.clientHeight) { popTopAdjust=MouseY+popsoffsetY+popHeight-document.body.clientHeight; }
  else { popTopAdjust=0; }
  dypopLayer.style.left=MouseX+popsoffsetX+document.body.scrollLeft-popLeftAdjust;
  dypopLayer.style.top=MouseY+popsoffsetY+document.body.scrollTop-popTopAdjust;
  dypopLayer.style.filter="Alpha(Opacity=0)";
  fadeOut();
}
function fadeOut(){
if(dypopLayer.filters.Alpha.opacity<popOpacity) {
dypopLayer.filters.Alpha.opacity+=showPopStep;
tFadeOut=setTimeout("fadeOut()",1);
}
else {
dypopLayer.filters.Alpha.opacity=popOpacity;
tFadeWaiting=setTimeout("fadeIn()",tPopShow);
}
}
function fadeIn(){
if(dypopLayer.filters.Alpha.opacity>0) {
dypopLayer.filters.Alpha.opacity-=1;
tFadeIn=setTimeout("fadeIn()",1);
}
}
document.onmouseover=showPopupText;
</script> 

<%
Call BottomTable()
xjweb.footer()
closeObj()
%>