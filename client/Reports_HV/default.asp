<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>Reports</title>
<script type="text/javascript" src="js/jquery-1.3.2.min.js"></script>
<script type=text/javascript>
$(function(){
	$('#webmenu li').hover(function(){
		$(this).children('ul').stop(true,true).show('slow');
	},function(){
		$(this).children('ul').stop(true,true).hide('slow');
	});
	
	$('#webmenu li').hover(function(){
		$(this).children('div').stop(true,true).show('slow');
	},function(){
		$(this).children('div').stop(true,true).hide('slow');
	});
});
</script>
<link href="css/Styles.css" rel="stylesheet" type="text/css" />
<style type="text/css">
* {
	margin:0;
	padding:0;
}
body {
	width:1180px;
	font-family:Arial, Verdana, Helvetica, sans-serif;
	font-size:1em;
	font-size:17px;
	color:#FFF;
	background-color:#339966;
	PADDING: 0px auto;
	MARGIN: 5px auto;
	text-align:left;
}
a {
	color:#FFF;
	text-decoration:none; 
}
ul {
	list-style:none;
}
#webmenu {
	height:37px;
	background:#333;
	font-size:1.3em;
	text-align:center;
	POSITION: absolute;
}
#webmenu a {
	font-size:0.65em;
}
#webmenu li ul {
	display:none;
}
#webmenu li ul li {
	float:none;
}
*html #webmenu li ul li {
	display:inline;
}
#webmenu li ul a {
	float:none;
	height:32px;
	line-height:32px;
	padding:0 10px;
	text-transform:capitalize;
}
#webmenu .height-auto {
	line-height:15px;
	padding:5px 10px;
}
.second-menu, .third-menu, .fourth-menu {
	position:absolute;
}
.first-menu li {
	float:left;
	position:relative;
}
.first-menu a {
	float:left;
	display:block;
	padding:0 20px;
	height:35px;
	line-height:35px;
	background:#333;
	border-top:1px solid #4a4a4a;
	border-left:1px solid #4a4a4a;
	border-bottom:1px solid #242424;
	border-right:1px solid #242424;
	font-size:.7em;
}
.first-menu a:hover {
	background:#4698ca;
	border-top:1px solid #5db1e0;
	border-left:1px solid #5db1e0;
}
.second-menu {
	top:37px;
	right:0;
}
*html .second-menu {
	right:-1px;
}
.second-menu a {
	font-size:11px;
}
.second-menu a.arrow {
	background:#3a3a3a url(images/arrow.gif) no-repeat right top;
}
.second-menu a.arrow:hover {
	background:#4698ca url(images/arrow.gif) no-repeat right -32px;
}
.second-menu a.arrow-02 {
	background:#3a3a3a url(images/arrow.gif) no-repeat right -64px;
}
.second-menu a.arrow-02:hover {
	background:#4698ca url(images/arrow.gif) no-repeat right -110px;
}
.third-menu, .fourth-menu {
	width:123px;
	top:0;
	left:123px;
}
.third-menu a {
	background:#4c4c4c;
	font-weight:normal;
	border-top:1px solid #595959;
	border-left:1px solid #595959;
	border-bottom:1px solid #333;
	border-right:1px solid #333;
}
#subMgm {
	width:123px;
}
#subMgm .third-menu {
	left:123px;
}
#subMgm .fourth-menu {
	left:123px;
}
#subMusic, #subNews {
	width:123px;
}
</style>
</head>
<body>
<div id="header"></div>
<div id="menu">
	<ul id="webmenu" class="first-menu">
	
	  <li><a href="#" target="_self">Line Reports</a>
		<ul style="display: none;" id="subMusic" class="second-menu">
		  <li><a href="DailyReportHv.asp" target="main">Daily</a></li>
		  <li><a href="WeeklyReportHv.asp" target="main">Weekly</a></li>
		  <li><a href="IntervalReportHv.asp" target="main">Interval</a></li>
		  <li><a href="DailyMaintain.asp" target="main">Maintain</a></li>
		  
		  <li><a href="#" class="arrow" target="_self">定制报告</a>
			<ul class="third-menu">
			  <li><a href="Maintains.asp" target="main">Qin Qi</a></li>
			  <li><a href="DailyOutput.asp" target="main">Wang Jianjun</a></li>
			</ul>
		  </li>
		</ul>
	  </li>
	  
	  <li><a href="#" target="_self">Periphery Reports</a>
		<ul id="subNews" class="second-menu">
		  <li><a href="DailyReportHvPer.asp" target="main">Daily</a></li>
		</ul>
	  </li>
	  
	</ul>
</div>
<div id="content1">
	<iframe src="DailyReportHv.asp" id="main" name="main" marginheight="0" frameborder="0" scrolling="no" vspace="0" hspace="0" marginwidth="0" width="100%" onLoad="iFrameHeight()">
	</iframe>
	<script type="text/javascript" language="javascript"> 
	function iFrameHeight() { 
	var ifm= document.getElementById("main"); 
	var subWeb = document.frames ? document.frames["main"].document : ifm.contentDocument; 
	if(ifm != null && subWeb != null) { 
	ifm.height = subWeb.body.scrollHeight; 
	} 
	} 
	</script>
</div>
<div id="footer"><p>&copy; KEB Production System (BPS)</p></div>
</body>
</html>
