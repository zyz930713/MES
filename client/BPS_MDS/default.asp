<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Marigold MDS</title>
<link href="css/default.css" rel="stylesheet" type="text/css" />
<script language="javascript" type="text/javascript" src="js/jquery.pack.js"></script>
</head>
<body>
<div id="header"><p>Machine Data System</p></div>
<div id="menu">
	<div id="MutiNav">
		<ul class="multiUl">
			<li>
				<a class="go" href="Output.asp" title="项目产出" target="main">项目产出</a>
			</li>
			
			<li class="s">|</li>
			
			<li>
				<a class="go" href="StFOR.asp" title="过程良率" target="main">过程良率</a>
			</li>

			<li class="s">|</li>
			
			<li>
				<a class="go" href="CarrierOut.asp" title="载体_产出统计" target="main">载体_产出统计</a>
			</li>
			
			<li class="s">|</li>
			
			<li>
				<a class="go" href="CarrierLoop.asp" title="载体_循环统计" target="main">载体_循环统计</a>
			</li>
			
			<li class="s">|</li>
			
			<li>
				<a class="go" href="Reports.asp" title="产出报告" target="main">产出报告</a>
			</li>
			
			<li class="s">|</li>
			
			<li>
				<a class="go" href="http://keb-prodapp:9901/Results/TopFailure.aspx" title="声学测试" target="main">声学测试</a>
			</li>
		</ul>
		<script type="text/javascript">
					var mst;
					jQuery(".multiUl li").hover(function(){
					var curItem = jQuery(this);
					mst = setTimeout(function(){
						curItem.find("blockquote").slideDown('_default');
						mst = null;
					});
					}, function(){
					if(mst!=null)clearTimeout(mst);
					jQuery(this).find("blockquote").slideUp('_default');
					});
		</script>
		<div class="clear"></div>
	</div>
</div>
<div id="content">
<iframe src="Output.asp" id="main" name="main" style="z-index:-1;" marginheight="0" frameborder="0" scrolling="no" vspace="0" hspace="0" marginwidth="0" width="100%" onLoad="iFrameHeight()">
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
<div id="footer">© KEB Production System (BPS)</div>
</body>
</html>