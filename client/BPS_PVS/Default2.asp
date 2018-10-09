<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<title>PVS Test</title>
<script type="text/javascript" language="javascript" src="../include/AnyChart/js/AnyChart.js"></script>
</head>
<body>
	<div id="container02"></div>
    <script type="text/javascript" language="javascript">
    var chart = new AnyChart('../include/AnyChart/swf/AnyChart.swf');
	chart.width = 1024;
    chart.height = 1200;
    chart.setXMLFile('data02.asp');
    chart.write('container02');
    </script>
</body>
</html>