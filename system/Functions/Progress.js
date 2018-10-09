// JavaScript Document
function progress(totalframe,nowframe,imagewidth,imagename,percentname)
{
var percent=nowframe/totalframe;
//imagename.width=Math.round(percent*imagewidth);
percentname.innerText=Math.round(percent*100)+"%";
}