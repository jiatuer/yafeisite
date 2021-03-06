//***********默认设置定义.*********************
tPopWait=50;//停留tWait毫秒后显示提示。
tPopShow=3000;//显示tShow毫秒后关闭提示
showPopStep=20;
popOpacity=99;

//***************内部变量定义*****************
sPop=null;
curShow=null;
tFadeOut=null;
tFadeIn=null;
tFadeWaiting=null;

document.write("<style type='text/css'id='defaultPopStyle'>");
document.write(".cPopText {  background-color: #ffffff; border: 1px #000000 solid; font-size: 12px; padding-right: 4px; padding-left: 4px; height: 20px; padding-top: 2px; padding-bottom: 2px; filter: Alpha(Opacity=0)}");

document.write("</style>");
document.write("<div id='zzzpopLayer' style='position:absolute;z-index:1000;' class='cPopText'></div>");


function showPopupText(){
var o=event.srcElement;
	MouseX=event.x;
	MouseY=event.y;
	if(o.alt!=null && o.alt!=""){o.zzzpop=o.alt;o.alt=""};
        if(o.title!=null && o.title!=""){o.zzzpop=o.title;o.title=""};
	if(o.zzzpop!=sPop) {
			sPop=o.zzzpop;
			clearTimeout(curShow);
			clearTimeout(tFadeOut);
			clearTimeout(tFadeIn);
			clearTimeout(tFadeWaiting);	
			if(sPop==null || sPop=="") {
				zzzpopLayer.innerHTML="";
				zzzpopLayer.style.filter="Alpha()";
				zzzpopLayer.filters.Alpha.opacity=0;	
				}
			else {
				if(o.dyclass!=null) popStyle=o.dyclass 
					else popStyle="cPopText";
				curShow=setTimeout("showIt()",tPopWait);
			}
			
	}
}

function showIt(){
		zzzpopLayer.className=popStyle;
		zzzpopLayer.innerHTML=sPop;
		popWidth=zzzpopLayer.clientWidth;
		popHeight=zzzpopLayer.clientHeight;
		if(MouseX+12+popWidth>document.body.clientWidth) popLeftAdjust=-popWidth-24
			else popLeftAdjust=0;
		if(MouseY+12+popHeight>document.body.clientHeight) popTopAdjust=-popHeight-24
			else popTopAdjust=0;
		zzzpopLayer.style.left=MouseX+12+document.body.scrollLeft+popLeftAdjust;
		zzzpopLayer.style.top=MouseY+12+document.body.scrollTop+popTopAdjust;
		zzzpopLayer.style.filter="Alpha(Opacity=0)";
		fadeOut();
}

function fadeOut(){
	if(zzzpopLayer.filters.Alpha.opacity<popOpacity) {
		zzzpopLayer.filters.Alpha.opacity+=showPopStep;
		tFadeOut=setTimeout("fadeOut()",1);
		}
		else {
			zzzpopLayer.filters.Alpha.opacity=popOpacity;
			tFadeWaiting=setTimeout("fadeIn()",tPopShow);
			}
}

function fadeIn(){
	if(zzzpopLayer.filters.Alpha.opacity>0) {
		zzzpopLayer.filters.Alpha.opacity-=1;
		tFadeIn=setTimeout("fadeIn()",1);
		}
}
document.onmouseover=showPopupText;