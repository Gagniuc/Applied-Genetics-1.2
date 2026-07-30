var ie=document.all
var dom=document.getElementById
var ns4=document.layers
var calunits=document.layers? "" : "px"

var bouncelimit=32 
var direction="up"

var persistclose=0 
var startX = 30 
var startY = 5 
var verticalpos="fromtop" 

var TRange = null;
var dupeRange = null;
var TestRange = null;
var win = null;
var nom = navigator.appName.toLowerCase();
var agt = navigator.userAgent.toLowerCase();
var is_major   = parseInt(navigator.appVersion);
var is_minor   = parseFloat(navigator.appVersion);
var is_ie      = (agt.indexOf("msie") != -1);
var is_ie4up   = (is_ie && (is_major >= 4));
var is_not_moz = (agt.indexOf('netscape')!=-1)
var is_nav     = (nom.indexOf('netscape')!=-1);
var is_nav4    = (is_nav && (is_major == 4));
var is_mac     = (agt.indexOf("mac")!=-1);
var is_gecko   = (agt.indexOf('gecko') != -1);
var is_opera   = (agt.indexOf("opera") != -1);function initbox(){
if (!dom&&!ie&&!ns4)
return
crossobj=(dom)?document.getElementById("dropin").style : ie? document.all.dropin : document.dropin
scroll_top=(ie)? truebody().scrollTop : window.pageYOffset
crossobj.top=scroll_top-250+calunits
crossobj.visibility=(dom||ie)? "visible" : "show"
dropstart=setInterval("dropin()",50)}function dropin(){
scroll_top=(ie)? truebody().scrollTop : window.pageYOffset
if (parseInt(crossobj.top)<100+scroll_top)
crossobj.top=parseInt(crossobj.top)+40+calunits
else{
clearInterval(dropstart)
bouncestart=setInterval("bouncein()",50)}}function bouncein(){
crossobj.top=parseInt(crossobj.top)-bouncelimit+calunits
if (bouncelimit<0)
bouncelimit+=8
bouncelimit=bouncelimit*-1
if (bouncelimit==0){
clearInterval(bouncestart)}}function dismissbox(){
if (window.bouncestart) clearInterval(bouncestart)
crossobj.visibility="hidden"}function truebody(){
return (document.compatMode && document.compatMode!="BackCompat")? document.documentElement : document.body
}function get_cookie(Name) {
var search = Name + "="
var returnvalue = ""
if (document.cookie.length > 0) {
offset = document.cookie.indexOf(search)
if (offset != -1) {
offset += search.length
end = document.cookie.indexOf(";", offset)
if (end == -1)
end = document.cookie.length;
returnvalue=unescape(document.cookie.substring(offset, end))}}return returnvalue;
}function dropornot(){
if (get_cookie("droppedin")==""){
window.onload=initbox
document.cookie="droppedin=yes"
}
}
dropornot();function iecompattest(){
return (document.compatMode && document.compatMode!="BackCompat")? document.documentElement : document.body}function get_cookie(Name) {
var search = Name + "="
var returnvalue = "";
if (document.cookie.length > 0) {
offset = document.cookie.indexOf(search)
if (offset != -1) {
offset += search.length
end = document.cookie.indexOf(";", offset);
if (end == -1) end = document.cookie.length;
returnvalue=unescape(document.cookie.substring(offset, end))
}
}
return returnvalue;
}function closebar(){
if (persistclose)
document.cookie="remainclosed=1"
document.getElementById("topbar").style.visibility="hidden"}function staticbar(){
	barheight=document.getElementById("topbar").offsetHeight
	var ns = (navigator.appName.indexOf("Netscape") != -1) || window.opera;
	var d = document;
	function ml(id){
		var el=d.getElementById(id);
		if (!persistclose || persistclose && get_cookie("remainclosed")=="")
		el.style.visibility="visible"
		if(d.layers)el.style=el;
		el.sP=function(x,y){this.style.left=x+"px";this.style.top=y+"px";};
		el.x = startX;
		if (verticalpos=="fromtop")
		el.y = startY;
		else{
		el.y = ns ? pageYOffset + innerHeight : iecompattest().scrollTop + iecompattest().clientHeight;
		el.y -= startY;
		}
		return el;
	}
	window.stayTopLeft=function(){
		if (verticalpos=="fromtop"){
		var pY = ns ? pageYOffset : iecompattest().scrollTop;
		ftlObj.y += (pY + startY - ftlObj.y)/8;
		}
		else{
		var pY = ns ? pageYOffset + innerHeight - barheight: iecompattest().scrollTop + iecompattest().clientHeight - barheight;
		ftlObj.y += (pY - startY - ftlObj.y)/8;
		}
		ftlObj.sP(ftlObj.x, ftlObj.y);
		setTimeout("stayTopLeft()", 10);
	}
	ftlObj = ml("topbar");
	stayTopLeft();}
if (window.addEventListener)
window.addEventListener("load", staticbar, false)
else if (window.attachEvent)
window.attachEvent("onload", staticbar)
else if (document.getElementById)
window.onload=staticbar

var is_rev=0;
if (is_gecko) {
temp = agt.split("rv:")
is_rev = parseFloat(temp[1])
}
var frametosearch = self;

function search(whichform, whichframe) {
if (is_ie4up && is_mac) return;
if (is_gecko && (is_rev <1)) return;
if (is_opera) return;
if(whichform.findthis.value!='') {
       str = whichform.findthis.value;
       win = whichframe;
       var frameval=false;if(win!=self){
       frameval=true;
       win = parent.frames[whichframe];}}
else return;
var strFound;if(is_nav4 && (is_minor < 5)) {strFound=win.find(str);}
if (is_gecko && (is_rev >= 1)) {if(frameval!=false) win.focus();
    strFound=win.find(str, false, false, true, false, frameval, false);
if (is_not_moz)  whichform.findthis.focus();}if (is_ie4up) {if (TRange!=null) {  
   TestRange=win.document.body.createTextRange();if (dupeRange.inRange(TestRange)) {
   TRange.collapse(false);
   strFound=TRange.findText(str);if (strFound) {
        win.document.body.scrollTop = win.document.body.scrollTop + TRange.offsetTop;
        TRange.select();
        }}else {
     TRange=win.document.body.createTextRange();
     TRange.collapse(false);
     strFound=TRange.findText(str);
     if (strFound) {
        win.document.body.scrollTop = TRange.offsetTop;
        TRange.select();}}}if (TRange==null || strFound==0) {
   TRange=win.document.body.createTextRange();
   dupeRange = TRange.duplicate();
   strFound=TRange.findText(str);if (strFound) {
        win.document.body.scrollTop = TRange.offsetTop;
        TRange.select();}}}if (!strFound) alert ("Cuvantul '"+str+"' nu a fost gasit!")}