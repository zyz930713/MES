<%session("Page_Role") = ",HV_Reports_Reader"%>
<!--#include virtual="/Functions/Login_Check.asp"-->
<!--#include virtual="/Functions/Role_Check.asp"-->
<!--#include virtual="/include/Functions.asp"-->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/include/Charts/FusionCharts.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>HV - Editor - Line</title>
<link rel="stylesheet" type="text/css" href="Styles/basic.css">
<script type="text/javascript" language="javascript" src="../include/ThreeMenu.js"></script>
<script type="text/javascript" language="javascript" src="../include/DatePicker/WdatePicker.js"></script>
<script type="text/javascript" language="javascript" src="../include/Charts/FusionCharts.js" ></script>
<script type="text/javascript" language="javascript" src="../include/Charts/FusionChartsExportComponent.js"></script>
<style type="text/css">
body{
font:normal 16px/16px "Arial";
}
table{
border-collapse:collapse;
border-spacing:0;
border:1px solid black;
}
table th{
border-color:black;
}
table td{
font:normal 14px/16px "Arial";
border-color:black;
}
a{
color:#fff;
text-decoration:none;
}
body,td,th {
	font-family: Arial;
}
</style>
</head>
<body>
<%
ProductName = request("ProductName")
LineName = request("LineName")
Adate = request("Adate")
DName = request("DName")
SearchEdit = request("SearchEdit")

if Adate = "" then
	Adate = FormatTime((date()-1),2)
end if

WeekFirstDay = DateAdd("d",0 - Weekday(Adate,2) + 1, CDate(Adate))
'response.end
%>
<center>
<br/>
<form method="post" action="DailyMaintain.asp">

	产品：<select name="ProductName" id="TmProduct"></select>
	　线别：<select name="LineName" id="TmLine"></select>
		<select name="CellName" id="TmCell" style="display:none" ></select>
			<script type="text/javascript">
				MenuInit('TmProduct','TmLine','TmCell','<%=ProductName%>','<%=LineName%>','1');
			</script>
	　日期：<input type="text" name="Adate" class="Wdate" id="Adate" onFocus="WdatePicker({isShowWeek:true,onpicked:function() {$dp.$('Adate_1').value=$dp.cal.getP('W','W');$dp.$('Adate_2').value=$dp.cal.getP('W','WW');}})" value="<%=Adate%>" size="10" />
	　D/N：<select name="DName" >
					<option value="D" <%if DName = "D" then response.write "selected"%> >Day</option>
					<option value="N" <%if DName = "N" then response.write "selected"%> >Night</option>
				</select>
	　<input type="submit" name="SearchEdit" value="刷新">

</form>
<center>
<table width="100%" border="0">
<%
if SearchEdit = "刷新" then
	set rs = Server.CreateObject("adodb.recordset")
	HaveData = 0
	sql = "SELECT [Rid],[QtySM],[QtyOC] ,[QtyTS],[QtyMT],[QABlock],[QARelease] FROM [dbo].[Day_Data_Online] where ProductName = '"&ProductName&"' and LineName = '"&LineName&"' and DName = '"&DName&"' and Adate = '"&Adate&"' order by Rid"
	rs.open sql,ConnSql,1,1
		QtySM = 0
		QtyOC = 0
		QtyTS = 0
		QtyMT = 0
		QABlock = 0
		QARelease = 0
	if not rs.eof then
		HaveData = 1
		OnlineId = rs("Rid")
		QtySM = rs("QtySM")
		QtyOC = rs("QtyOC")
		QtyTS = rs("QtyTS")
		QtyMT = rs("QtyMT")
		QABlock = rs("QABlock")
		QARelease = rs("QARelease")
	end if
	rs.close
	
	sql = "SELECT * FROM [dbo].[Day_Data_Machine] where ProductName = '"&ProductName&"' and LineName = '"&LineName&"' and DName = '"&DName&"' and Adate = '"&Adate&"' order by Dtime desc"
	'response.write sql
	strXMLds = ""
	rs.open sql,ConnSql,1,1
		do while not rs.eof
		HaveData = 1
		Dlossn = rs("Dlossn")
		DlossName = XlName(right(Dlossn,1))
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
		Call renderChart("../include/charts/Column2D.swf?ChartNoDataText=No data to display.&amp;#xA;", "", strXMLMSL, "myNext00", "550", "400", false,true)
		response.write "</td><td>"
		Call renderChart("../include/charts/Pie2D.swf?ChartNoDataText=No data to display.&amp;#xA;", "", strXMLMSL, "myNext01", "550", "400", false,true)
		response.write "</tr></table><br/>"
		response.write "<table width='100%' border='1'>"
		response.write "<tr><th>维护项目</th><th>起止时间</th><th>时长</th><th>设备</th><th>工位</th><th>故障描述</th><th>故障原因</th><th>解决方法</th></tr>"
		response.write TableStr
		response.write "</table><br/>"
	else
		response.write "<div style='width:960px;height:480px;background:#FFFFFF;line-height:480px;font-size:32px;'><p>没有找到所需的数据！</p></div>"
	end if

end if

function XlName(Xltype)
	select case Xltype
		case 1
			XlName = "更换零备件"
		case 2
			XlName = "维修"
		case 3
			XlName = "设备调整"
		case 4
			XlName = "计划性检测"
		case 5
			XlName = "单位加工时间"
		case 6
			XlName = "速度不匹配"
		case 7
			XlName = "小停顿"
		case 8
			XlName = "过程废品"
		case 9
			XlName = "初始不良品"
	end select
end function
%>
</table>
</body>
</html>