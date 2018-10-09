<!--#include file= "include/function.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>RealTime</title>
<script language="javascript" type="text/javascript" src="include/DatePicker/WdatePicker.js"></script>
<script language="javascript" type="text/javascript" src="include/ThreeMenu.js"></script>
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
<body style="font-size:14px;background-color:#339966;">
<OBJECT id="WebBrowser" height="0" width="0" classid="CLSID:8856F961-340A-11D0-A96B-00C04FD705A2" style="display:none"></OBJECT>
<%
Set conn = Server.CreateObject("ADODB.Connection") 
conn.Open "driver={SQL Server};server=KEW-SQL2008;uid=KEWAMS;pwd=KEWAMS;database=TSDHV"

Getdata = request("Getdata")
HDate = request("HDate")
EndDate = request("EndDate")
ProductName = request("ProductName")
LineName = request("LineName")
CellName = request("CellName")


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
<center>
<form method="post" action="ams.asp">
	<table width="962" border="1" cellspacing="0" cellpadding="0">
		<tr style="height:30px;line-height:30px;">
			<td align="center">
				产品：测试机
			  起始时间：<input type="text" class="Wdate" onFocus="WdatePicker({startDate:'%y-%M-%d 07:15:00',dateFmt:'yyyy-MM-dd HH:mm:ss',alwaysUseStartDate:true})" name="HDate" value="<%=HDate%>" size="20" />&nbsp;
			  终止时间：<input type="text" class="Wdate" onFocus="WdatePicker({startDate:'%y-%M-%d 19:15:00',dateFmt:'yyyy-MM-dd HH:mm:ss',alwaysUseStartDate:true})" name="EndDate" value="<%=EndDate%>" size="20" />&nbsp;<input type="submit" name="Getdata" value="刷新" />　<input type="submit" name="Getdata" value="产量" />
			</td>
		</tr>
	</table>
</form>
<center>
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

	CellValue = 0
	OutputValue = 0
'----------------------------获取Target-------------------------------------------'	

	set rs = Server.CreateObject("Adodb.recordset")
	sql = "EXEC [dbo].[SP_GetDataTest] @TestAmsName = N'26000171BJGAS01B',@StartTime = N'"&Startdate&"',@EndTime = N'"&EndDate&"'"
	'response.write sql & "</br>"
	rs.open sql,conn,1,1
	if not rs.eof then
	
		Set fso = CreateObject("Scripting.FileSystemObject")
		FileNo = "26000171BJGAS01B"
		OutXml = "Data\26000171BJGAS01B.xml" 'Output数据文件输出路径
		OutXmlPath = server.mappath(OutXml)
		if fso.fileexists(OutXmlPath) then fso.DeleteFile(OutXmlPath) '文件存在，就进行删除
		
		Set OutXmlFile = fso.Createtextfile(server.mappath(OutXml),true)
		
		OutputValue = 1
		OutXmlFile.writeline "<graph caption='26000171BJGAS01B' baseFontSize='10' xAxisName='Time' yAxisName='FOR' numberSuffix='%25' decimalPrecision='2' formatNumberScale='0' rotateNames='0' showValues='1' yAxisMaxValue='1'>"
		do while not rs.eof
			OutXmlFile.writeline "<set name='" & rs("TestName") & "' value='" & rs("FOR") & "' color='AFD8F8' />"
		rs.movenext
		loop
		'RaTarget = Int(DiffTime * 1.446)
		' OutXmlFile.writeline "<trendlines>"
		' OutXmlFile.writeline "<line startValue='"&Outpieces&"' color='ff0000' displayValue='"&Outpieces&"' showOnTop='1' thickness='2'/>"
		' OutXmlFile.writeline "<line startValue='0' endValue='" & RaTarget & "' color='ff0000' displayValue='' showOnTop='1' thickness='1'/>"
		' OutXmlFile.writeline "</trendlines>"
		OutXmlFile.writeline "</graph>"
		OutXmlFile.close
	end if
	Set fso = nothing

	if OutputValue = 0 and CellValue = 0 then
		response.write "<div style='width:960px;height:480px;background:#FFFFFF;line-height:480px;font-size:32px;'><p>没有找到所需的数据！</p></div>"
	end if
end if

if OutputValue = 1 then
%>	
<div id="chartdiv2" align="center">FusionCharts. </div>
		<script type="text/javascript">
			var chart = new FusionCharts("Charts/Column2D.swf", "ChartId2", "960", "400");
			chart.setDataURL("Data/26000171BJGAS01B.xml");		   
			chart.render("chartdiv2");
		</script>
<%end if
		response.write "<script>"
		response.write "var oLoading=document.getElementById('loading');"
		response.write "var parent=oLoading.parentNode;"
		response.write "parent.removeChild(oLoading);"
		response.write "</script>"
%>
</body>
</html>
