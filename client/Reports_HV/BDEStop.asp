<!--#include file="conn.asp"-->
<!--#include virtual="/BpsEditor/include/function.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>BDE Stop</title>
<OBJECT ID="jatoolsPrinter" CLASSID="CLSID:B43D3361-D075-4BE2-87FE-057188254255" codebase="jatoolsPrinter.cab#version=8,3,0,0"></OBJECT>
<script language="javascript" type="text/javascript" src="../include/DatePicker/WdatePicker.js"></script>
<script language="javascript" type="text/javascript" src="../include/ThreeMenu.js"></script>
<script language="JavaScript" src="Charts/FusionCharts.js"></script>

<style type="text/css">
table tr{
	height:26px;
	line-height:26px;
	background:#FFFFFF;
}
table tr td{
	text-align:center;
}
@media print {#btn{display: none;}}
</style>
</head>
<body style="background-color:#FFFFFF;">

<%
ProductName = request("ProductName")
LineName = request("LineName")
Adate = request("Adate")
DName = request("DName")
SearchEdit = request("SearchEdit")

if Adate = "" then
	Adate = FormatTime((date()-1),2)
end if
%>
<br/>
<center>
开发中.............
<form method="post" action="BDEStop.asp">
	产品：<select name="ProductName" id="TmProduct"></select>
	　线别：<select name="LineName" id="TmLine"></select>
		<select name="CellName" id="TmCell" style="display:none" ></select>
			<script type="text/javascript">
				MenuInit('TmProduct','TmLine','TmCell','<%=ProductName%>','<%=LineName%>','1');
			</script>
	　日期：<input type="text" name="Adate" class="Wdate" id="Adate" onFocus="WdatePicker({isShowWeek:true,onpicked:function() {$dp.$('Adate_1').value=$dp.cal.getP('W','W');$dp.$('Adate_2').value=$dp.cal.getP('W','WW');}})" value="<%=Adate%>" size="10" />
	　<input type="submit" name="SearchEdit" value="刷新">
</form>
<center>
<div id='page1' style='border:0px solid green'>
<%
response.end
if Getdata = "刷新" then
	Response.buffer=true '开启缓存
	Response.Write "<div id='loading' style='width:960px;height:480px;background:#FFFFFF url(images/loading.gif) center no-repeat;'>" '先写入提示
	'Response.Write "<img src='images/loading.gif' style='vertical-align:middle;'/>" '先写入提示
	Response.Write "</div>"
	Response.flush() '输出缓存
	Response.clear()

	Startdate = Cdate(HDate)
	EndDate = Cdate(EndDate)
	
if SearchEdit = "刷新" then

	sql = "SELECT * FROM [dbo].[Day_Data_Machine] where ProductName = '"&ProductName&"' and LineName = '"&LineName&"' and DName = '"&DName&"' and Adate = '"&Adate&"' order by Dtime desc"
	'response.write sql
	strXMLds = ""
	rs.open sql,ConnSql,1,1
		do while not rs.eof
		HaveData = 1
		DlossName = rs("Dlossn")
		Adate = FormatTime(rs("Adate"),2)
		Dhtime = rs("Dhtime")
		Detime = rs("Detime")
		Dtime = rs("Dtime")
		Dequipment = rs("Dequipment")
		DPosition = rs("DPosition")
		Dproblem = rs("Dproblem")
		DReason = rs("DReason")
		Dsolution = rs("Dsolution")
		strXMLds = strXMLds &" <set label='" & DlossName & "' issliced='1' value='" & Dtime & "'/>"
		
		TableStr = TableStr & "<tr><td>" & DlossName & "</td><td>" & Dhtime & " - " & Detime & "</td><td>" & Dtime & "</td><td>" & Dequipment & "</td><td>" & DPosition & "</td><td align='left'>" & Dproblem & "</td><td>" & DReason & "</td><td>" & Dsolution & "</td></tr>"
		
		rs.movenext
		loop
	rs.close

	if HaveData = 1 then
		strXMLMSL = "<chart caption='Report' baseFontSize='12' outCnvBaseFontColor='000000' labelDisplay='Rotate' baseFontColor='0000CD' palette='2' animation='1' formatnumberscale='1' decimals='2' pieslicedepth='100' startingangle='-60' showPercentValues='1' showPercentInToolTip='0' labelSepChar='~'>"&strXMLds& "</chart>"
		response.write "<table width='100%'><tr><td>"
		Call renderChart("charts/Column2D.swf?ChartNoDataText=No data to display.&amp;#xA;", "", strXMLMSL, "myNext00", "550", "400", false,true)
		response.write "</td><td>"
		Call renderChart("charts/Pie2D.swf?ChartNoDataText=No data to display.&amp;#xA;", "", strXMLMSL, "myNext01", "550", "400", false,true)
		response.write "</tr></table><br/>"
		response.write "<table width='100%' border='1'>"
		response.write "<tr><th>维护项目</th><th>起止时间</th><th>时长</th><th>设备</th><th>工位</th><th>故障描述</th><th>故障原因</th><th>解决方法</th></tr>"
		response.write TableStr
		response.write "</table><br/>"
	else
		response.write "<div style='width:960px;height:480px;background:#FFFFFF;line-height:480px;font-size:32px;'><p>没有找到所需的数据！</p></div>"
	end if
		
end if


		response.write "<script>"
		response.write "var oLoading=document.getElementById('loading');"
		response.write "var parent=oLoading.parentNode;"
		response.write "parent.removeChild(oLoading);"
		response.write "</script>"
end if
%>
</div>
</body>
</html>
