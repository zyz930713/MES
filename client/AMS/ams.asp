<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>AMS Testing Data</title>
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
<%
Set conn = Server.CreateObject("ADODB.Connection") 
conn.Open "driver={SQL Server};server=KEW-SQL2008;uid=KEWAMS;pwd=KEWAMS;database=TSDHV"

Getdata = request("Getdata")
Sdatetime = request("Sdatetime")
Edatetime = request("Edatetime")
MCode = request("MCode")

if MCode = "" then MCode = "26000171BJGAS01B"
%>
<center>
<form method="post" action="ams.asp">
	<table width="962" border="1" cellspacing="0" cellpadding="0">
		<tr style="height:30px;line-height:30px;">
			<td align="center">
				测试机名：<input type="" name="MCode" value="<%=MCode%>"/>
				起始时间：<input type="text" class="Wdate" onFocus="WdatePicker({startDate:'%y-%M-%d 07:15:00',dateFmt:'yyyy-MM-dd HH:mm:ss',alwaysUseStartDate:true})" name="Sdatetime" value="<%=Sdatetime%>" size="20" />&nbsp;
				终止时间：<input type="text" class="Wdate" onFocus="WdatePicker({startDate:'%y-%M-%d 19:15:00',dateFmt:'yyyy-MM-dd HH:mm:ss',alwaysUseStartDate:true})" name="Edatetime" value="<%=Edatetime%>" size="20" />&nbsp;<input type="submit" name="Getdata" value="刷新" />
			</td>
		</tr>
	</table>
</form>
<center>
<%
if Getdata = "刷新" then
	Startdate = Cdate(Sdatetime)
	Edatetime = Cdate(Edatetime)

	CellValue = 0
	OutputValue = 0
'----------------------------获取Target-------------------------------------------'	

	set rs = Server.CreateObject("Adodb.recordset")
	sql = "EXEC [dbo].[SP_GetDataTest] @TestAmsName = N'"&MCode&"',@StartTime = N'"&Startdate&"',@EndTime = N'"&Edatetime&"'"
	rs.open sql,conn,1,1
	if not rs.eof then
	
		Set fso = CreateObject("Scripting.FileSystemObject")
		FileNo = MCode
		OutXml = "Data\"&MCode&".xml" 'Output数据文件输出路径
		OutXmlPath = server.mappath(OutXml)
		if fso.fileexists(OutXmlPath) then fso.DeleteFile(OutXmlPath) '文件存在，就进行删除
		Set OutXmlFile = fso.Createtextfile(server.mappath(OutXml),true)
		OutputValue = 1
		OutXmlFile.writeline "<graph caption='"&MCode&"' baseFontSize='10' xAxisName='TestName' yAxisName='FOR' numberSuffix='%25' decimalPrecision='2' formatNumberScale='0' rotateNames='0' showValues='1' yAxisMaxValue='1'>"
		do while not rs.eof
			OutXmlFile.writeline "<set name='" & rs("TestName") & "' value='" & rs("FOR") & "' color='0000CC' />"
		rs.movenext
		loop
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
			var chart = new FusionCharts("Charts/FCF_Column2D.swf", "ChartId2", "960", "400");
			chart.setDataURL("Data/<%=MCode%>.xml");		   
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
