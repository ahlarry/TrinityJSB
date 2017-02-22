//****************鼠标提示 开始*****************
var tPopWait=50;	//停留tWait豪秒后显示提示
var tPopShow=30000;	//显示tShow豪秒后关闭提示
var showPopStep=20;
var popOpacity=95;
var tfontcolor="#000000";
var tbgcolor="#ecedef";
var tbordercolor="#666666";
var sPop=null;curShow=null;tFadeOut=null;tFadeIn=null;tFadeWaiting=null;

function showPopupText()
{
  var o=event.srcElement;
  MouseX=event.x;
  MouseY=event.y;
  if(o.alt!=null && o.alt!="") { o.dypop=o.alt;o.alt=""; }
  if(o.title!=null && o.title!="") { o.dypop=o.title;o.title=""; }
  if(o.dypop!=sPop)
  {
    sPop=o.dypop;
    clearTimeout(curShow);
    clearTimeout(tFadeOut);
    clearTimeout(tFadeIn);
    clearTimeout(tFadeWaiting);  
    if(sPop==null || sPop=="")
    {
      div_poplayer.innerHTML="";
      div_poplayer.style.filter="Alpha()";
      div_poplayer.filters.Alpha.opacity=0;
    }
    else
    {
	  sPop=sPop.replace(/\r/gi,"　　")	//table
      sPop=sPop.replace(/\n/gi,"<br>")	//回车
      if(o.dyclass!=null) { popStyle=o.dyclass; }
      else { popStyle="div_pop"; }
      curShow=setTimeout("showIt()",tPopWait);
    }
  }
}

function showIt()
{
  div_poplayer.className=popStyle;
  sPop="<b>提示:</b><br>" + sPop;
  div_poplayer.innerHTML=sPop;
  popsoffsetX = 10;   // 弹出窗口位于鼠标左侧或者右侧的距离；3-12 合适
  popsoffsetY = 15;  // 弹出窗口位于鼠标下方的距离；3-12 合适
  popWidth=div_poplayer.clientWidth;
  popHeight=div_poplayer.clientHeight;
  popLeftAdjust=0;
//  if(MouseX+12+popWidth>document.body.clientWidth) { popLeftAdjust=-popWidth-24; }
  if(MouseX+popsoffsetX+popWidth>document.body.clientWidth) { popLeftAdjust=MouseX+popsoffsetX+popWidth-document.body.clientWidth; }
  else { popLeftAdjust=0; }
//  if(MouseY+12+popHeight>document.body.clientHeight) { popTopAdjust=-popHeight-24; }
  if(MouseY+popsoffsetY+popHeight>document.body.clientHeight) { popTopAdjust=MouseY+popsoffsetY+popHeight-document.body.clientHeight; }
  else { popTopAdjust=0; }
  div_poplayer.style.left=MouseX+popsoffsetX+document.body.scrollLeft-popLeftAdjust;
  div_poplayer.style.top=MouseY+popsoffsetY+document.body.scrollTop-popTopAdjust;
  div_poplayer.style.filter="Alpha(Opacity=0)";
  fadeOut();
}

function fadeOut(){
  if(div_poplayer.filters.Alpha.opacity<popOpacity)
  {
    div_poplayer.filters.Alpha.opacity+=showPopStep;
    tFadeOut=setTimeout("fadeOut()",1);
  }
  else
  {
    div_poplayer.filters.Alpha.opacity=popOpacity;
    tFadeWaiting=setTimeout("fadeIn()",tPopShow);
  }
}

function fadeIn()
{
  if(div_poplayer.filters.Alpha.opacity>0)
  {
    div_poplayer.filters.Alpha.opacity-=1;
    tFadeIn=setTimeout("fadeIn()",1);
  }
}
document.write("<style type='text/css'id='defaultPopStyle'>");
document.write(".div_pop {background-repeat : repeat-y; background-color: " + tbgcolor + ";color:" + tfontcolor + "; border: 1px " + tbordercolor + " solid;font-color: font-size: 12px; padding-right: 4px; padding-left: 28px; height: 20px; padding-top: 2px; padding-bottom: 2px; filter: Alpha(Opacity=0);text-align:left;}");
document.write("</style>");
document.write("<div id='div_poplayer' style='position:absolute;z-index:1000;' class='div_pop'></div>");
document.onmouseover=showPopupText;
//****************鼠标提示 结束*****************