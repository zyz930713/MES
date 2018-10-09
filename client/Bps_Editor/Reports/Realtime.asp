<!--#include file="conn.asp"-->
<!--#include virtual="/BpsEditor/include/function.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>RealTime</title>
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
Getdata = request("Getdata")
HDate = request("HDate")
EndDate = request("EndDate")
ProductName = request("ProductName")
LineName = request("LineName")
CellName = request("CellName")

	'取Cookies缓存数据'
	if ProductName = "" then
		ProductName = request.Cookies("Realtime")("ProductName")
	else
		response.Cookies("Realtime")("ProductName") = ProductName
	end if
	
	if LineName = "" then
		LineName = request.Cookies("Realtime")("LineName")
	else
		response.Cookies("Realtime")("LineName") = LineName
	end if

	if CellName = "" then
		CellName = request.Cookies("Realtime")("CellName")
	else
		response.Cookies("Realtime")("CellName") = CellName
	end if
	'取Cookies缓存数据'

	NowDate = FormatTime(now(),2)
	TemTime2 = cstr(NowDate) & " 07:15:00"
	TemTime1 = dateadd("h",-12,cdate(TemTime2))
	TemTime3 = dateadd("h",12,cdate(TemTime2))
	TemTime4 = dateadd("h",12,cdate(TemTime3))
	Nt = now()
	if datediff("s",TemTime1,Nt) > 0 and datediff("s",Nt,TemTime2) > 0 then 
		'response.write "A"
		Sdatetime = TemTime1
		Edatetime = TemTime2
	elseif datediff("s",TemTime2,Nt) > 0 and datediff("s",Nt,TemTime3) > 0 then 
		'response.write "B"
		Sdatetime = TemTime2
		Edatetime = TemTime3
	elseif datediff("s",TemTime3,Nt) > 0 and datediff("s",Nt,TemTime4) > 0 then 
		'response.write "C"
		Sdatetime = TemTime3
		Edatetime = TemTime4
	end if
	
	'response.write "<center>" & Sdatetime & " - " & Edatetime & "</center>"
if HDate = "" then
	HDate = Sdatetime
end if
if EndDate = "" then
	EndDate = Edatetime
end if
%>
<br/>
<center>
	<form method="post" action="Realtime.asp">
		产品：<select name="ProductName" id="TmProduct"></select>
		&nbsp;线别：<select name="LineName" id="TmLine"></select>
		&nbsp;Cell：<select name="CellName" id="TmCell"></select>
		<script type="text/javascript">
			MenuInit('TmProduct', 'TmLine', 'TmCell', '<%=ProductName%>', '<%=LineName%>', '<%=CellName%>');
		</script>
		起始时间：<input type="text" class="Wdate" onFocus="WdatePicker({startDate:'%y-%M-%d 07:15:00',dateFmt:'yyyy-MM-dd HH:mm:ss',alwaysUseStartDate:true})" name="HDate" value="<%=HDate%>" size="20" />&nbsp;
		终止时间：<input type="text" class="Wdate" onFocus="WdatePicker({startDate:'%y-%M-%d 19:15:00',dateFmt:'yyyy-MM-dd HH:mm:ss',alwaysUseStartDate:true})" name="EndDate" value="<%=EndDate%>" size="20" />&nbsp;<input type="submit" name="Getdata" value="刷新" />　<input type="submit" name="Getdata" value="产量" />
	</form>
<center>
<div id='page1' style='border:0px solid green'>
<%
if Getdata = "刷新" then
	Response.buffer=true '开启缓存
	Response.Write "<div id='loading' style='width:960px;height:480px;background:#FFFFFF url(images/loading.gif) center no-repeat;'>" '先写入提示
	'Response.Write "<img src='images/loading.gif' style='vertical-align:middle;'/>" '先写入提示
	Response.Write "</div>"
	Response.flush() '输出缓存
	Response.clear()

	Startdate = Cdate(HDate)
	EndDate = Cdate(EndDate)

	if ProductName = "RA" then
		TitleNmae = "RA 11x15 Speaker (Danubius)"
	else
		TitleNmae = "Petra"
	end if
	CellValue = 0
	OutputValue = 0
'----------------------------获取Target-------------------------------------------'	
	dim CellFORValue,Outpieces
	CellFORValue = 0
	Outpieces = 0
	set Rs = Server.CreateObject("adodb.recordset")
	Sql = "EXEC [dbo].[SP_GetOutCellTargetForRealtime]	@StartDate = N'"&FormatTime(Startdate,2)&"',@KSTAcronym = N'"&ProductName&"',@LineName = N'"&LineName&"',@CellName = N'"&CellName&"'"
	Rs.open sql,conn2,1,1
	if not rs.eof then
		CellFORValue = rs("CellFOR")
		Outpieces = csng(rs("Outpieces"))
		Outpiecesh = Outpieces/12
	end if
	Rs.close
	set Rs = nothing
	
	'response.write Outpieces
	'response.end
'---------------------------------------------------------------------------------'
	set Rs = Server.CreateObject("adodb.recordset")
	Sql = "EXEC [MIS].[dbo].[SP_GetOutPiecesForRealtime] @ProductName = N'"&ProductName&"',@LineName = N'"&LineName&"',@CellName = N'"&CellName&"',@BeginTime = N'"&Startdate&"',@EndTime = N'"&EndDate&"'"
	'response.write sql & "</br>"
	Rs.open sql,conn,1,1
	if not rs.eof then
		CellValue = 1
		Set fso = CreateObject("Scripting.FileSystemObject")
		FileNo = ProductName&"L"&LineName&"C"&CellName
		OutXml = "Data\"&FileNo&"Line2D.xml" 'Output数据文件输出路径
		GoodXml = "Data\"&FileNo&"GoodColumn2D.xml" 'Good数据文件输出路径
		BadXml = "Data\"&FileNo&"BadColumn2D.xml" 'Bad数据文件输出路径
		OutXmlPath = server.mappath(OutXml)
		if fso.fileexists(OutXmlPath) then fso.DeleteFile(OutXmlPath) '文件存在，就进行删除
		GoodXmlPath = server.mappath(GoodXml)
		if fso.fileexists(GoodXmlPath) then fso.DeleteFile(GoodXmlPath) '文件存在，就进行删除
		BadXmlPath = server.mappath(BadXml)
		if fso.fileexists(BadXmlPath) then fso.DeleteFile(BadXmlPath) '文件存在，就进行删除
		
		Set OutXmlFile = fso.Createtextfile(server.mappath(OutXml),true)
		Set GoodXmlFile = fso.Createtextfile(server.mappath(GoodXml),true)
		Set BadXmlFile = fso.Createtextfile(server.mappath(BadXml),true)
			
		GoodXmlFile.writeline "<graph caption='Pieces/h " & ProductName & " Line" & LineName & " Cell" & CellName & "' baseFontSize='10' xAxisName='Time' yAxisName='Pieces' decimalPrecision='0' formatNumberScale='0' rotateNames='0' showValues='1' yAxisMaxValue='5500'>"
		BadXmlFile.writeline "<graph caption='FOR " & ProductName & " Line" & LineName & " Cell" & CellName & "' baseFontSize='10' xAxisName='Time' yAxisName='FOR' numberSuffix='%25' decimalPrecision='2' formatNumberScale='0' rotateNames='0' showValues='1' yAxisMaxValue='1'>"
		do while not rs.eof
			Xtime = Cint(FormatTime(rs("ScanZeit"),8))+1
			if Xtime = 24 then Xtime = Xtime - 1
			GoodXmlFile.writeline "<set name='" & Xtime & "' value='" & rs("Ganzahl") & "' color='AFD8F8' />"
			BadXmlFile.writeline "<set name='" & Xtime & "' value='" & rs("CellFor") & "' color='AFD8F8' />"
			rs.movenext
		loop
		GoodXmlFile.writeline "<trendlines>"
		GoodXmlFile.writeline "<line startValue='"&Outpiecesh&"' color='ff0000' displayValue='"&Outpiecesh&"' showOnTop='1' thickness='2'/>"
		GoodXmlFile.writeline "</trendlines>"
		GoodXmlFile.writeline "</graph>"
		GoodXmlFile.close
		BadXmlFile.writeline "<trendlines><line startValue='"&CellFORValue&"' color='ff0000' displayValue='"&CellFORValue&"' showOnTop='1' thickness='2'/></trendlines>"
		BadXmlFile.writeline "</graph>"
		BadXmlFile.close
		Set fso = nothing
	end if
	rs.close
	set rs = nothing
'-----------------------------------------------------------------------------------'	
	set rs = Server.CreateObject("Adodb.recordset")
	sql = "SELECT RZaehlerID,Anzahl,ScanZeit FROM [MIS].dbo.MisTempTable WHERE Cid = '"&FileNo&"' AND '"&Startdate&"' < [ScanZeit] AND [ScanZeit] < '"&EndDate&"' ORDER BY ScanZeit"
	'response.write sql & "</br>"
	rs.open sql,conn,1,1
	if not rs.eof then
		OutputValue = 1
		OutXmlFile.writeline "<graph caption='Pieces " & Productid & " Line" & LineName & " Cell" & CellName & "' baseFontSize='10' xAxisName='Time' yAxisName='Pieces' decimalPrecision='0' formatNumberScale='0' rotateNames='1' showValues='0' yAxisMaxValue='65000'>"
		do while not rs.eof
			OutXmlFile.writeline "<set name='" & FormatTime(rs("ScanZeit"),7) & "' value='" & rs("Anzahl") & "' color='AFD8F8' />"
			DiffTime = datediff("s",Cdate(Startdate),Cdate(rs("ScanZeit")))
		rs.movenext
		loop
		RaTarget = Int(DiffTime * 1.446)
		OutXmlFile.writeline "<trendlines>"
		OutXmlFile.writeline "<line startValue='"&Outpieces&"' color='ff0000' displayValue='"&Outpieces&"' showOnTop='1' thickness='2'/>"
		OutXmlFile.writeline "<line startValue='0' endValue='" & RaTarget & "' color='ff0000' displayValue='' showOnTop='1' thickness='1'/>"
		OutXmlFile.writeline "</trendlines>"
		OutXmlFile.writeline "</graph>"
		OutXmlFile.close
	end if
	Set fso = nothing

	if OutputValue = 0 and CellValue = 0 then
		response.write "<div style='width:960px;height:480px;background:#FFFFFF;line-height:480px;font-size:32px;'><p>没有找到所需的数据！</p></div>"
	end if
end if

if Getdata = "产量" then
	Startdate = Cdate(HDate)
	EndDate = Cdate(EndDate)
	set Rs = Server.CreateObject("adodb.recordset")
	Sql = "EXEC [dbo].[SP_AutoOutputForBPS]	@ProductName = N'"&ProductName&"',@LineName = N'"&LineName&"',@StartTime = N'"&Startdate&"',@EndTime = N'"&EndDate&"'"
	'response.write Sql
	'response.end
	Rs.open sql,conn,1,1
	
	if not rs.eof then
		'response.write "有数据"
		%>
		<table width="480" border="1" cellspacing="0" cellpadding="0">
			<tr><th colspan="3"><%=ProductName%> - <%=LineName%></th></tr>
			<tr><th>Cell Nr</th><th>Good</th><th>Bad</th></tr>
		<%
		do while not rs.eof
			response.write "<tr><td>Cell "&Rs("CellName")&"</td><td>"&Rs("Good")&"</td><td>"&Rs("Bad")&"</td></tr>"
		rs.movenext
		loop
		response.write "</table>"
		else
		response.write "<div style='width:960px;height:480px;background:#FFFFFF;line-height:480px;font-size:32px;'><p>没有找到所需的数据！</p></div>"
	end if
	rs.close
end if

if OutputValue = "1" then
%>
<div id="chartdiv1" align="center">FusionCharts. </div>
		<script type="text/javascript">
			var chart = new FusionCharts("Charts/Line.swf", "ChartId1", "960", "360");
			chart.setDataURL("Data/<%=FileNo%>Line2D.xml");		   
			chart.render("chartdiv1");
		</script>
<%
end if
if CellValue = 1 then
%>	
<div id="chartdiv2" align="center">FusionCharts. </div>
		<script type="text/javascript">
			var chart = new FusionCharts("Charts/Column2D.swf", "ChartId2", "960", "360");
			chart.setDataURL("Data/<%=FileNo%>GoodColumn2D.xml");		   
			chart.render("chartdiv2");
		</script>

<div id="chartdiv3" align="center">FusionCharts. </div>
		<script type="text/javascript">
			var chart = new FusionCharts("Charts/Column2D.swf", "ChartId3", "960", "360");
			chart.setDataURL("Data/<%=FileNo%>BadColumn2D.xml");		   
			chart.render("chartdiv3");
		</script>
<%end if
		response.write "<script>"
		response.write "var oLoading=document.getElementById('loading');"
		response.write "var parent=oLoading.parentNode;"
		response.write "parent.removeChild(oLoading);"
		response.write "</script>"
%>
</div>
</body>
</html>
