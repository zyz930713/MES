<!--#include file= "conn.asp"-->
<!--#include file= "../include/function.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<title>LPB</title>
<script language="javascript" type="text/javascript" src="../include/ThreeMenu.js"></script>
<script language="javascript" type="text/javascript" src="include/DatePicker/WdatePicker.js"></script>
<script language="JavaScript" src="../Charts/FusionCharts.js"></script>
<style type="text/css">
body{
font：normal 14px/16px "Arial";
}
#List {
	width: 644px;
	text-align: center;
	border:1px solid #E6E6E6;
}
#Nav1 {
	margin:5px 1px;
}

#Nav4 {
	width:644px;
	margin:5px 1px;
}
@media print {#btn{display: none;}}
</style>
</head>

<body>
<OBJECT id="WebBrowser" height="0" width="0" classid="CLSID:8856F961-340A-11D0-A96B-00C04FD705A2" VIEWASTEXT></OBJECT>
<%
	ProductName = request("ProductName")
	LineName = request("LineName")
	DDN = request("DDN")
	HShift = request("HShift")
	Ddate = request("Ddate")
	Pr = request("Pr")
	'取Cookies缓存数据'
	if ProductName = "" then
		ProductName = request.Cookies("Lpb")("ProductName")
	else
		response.Cookies("Lpb")("ProductName") = ProductName
	end if
	
	if LineName = "" then
		LineName = request.Cookies("Lpb")("LineName")
	else
		response.Cookies("Lpb")("LineName") = LineName
	end if
	
	if DDN = "" then
		DDN = request.Cookies("Lpb")("DDN")
	else
		response.Cookies("Lpb")("DDN") = DDN
	end if
	
	if HShift = "" then
		HShift = request.Cookies("Lpb")("HShift")
	else
		response.Cookies("Lpb")("HShift") = HShift
	end if
	
	if Ddate = "" then
		Ddate = FormatTime(request.Cookies("Lpb")("Ddate"),2)
	else
		response.Cookies("Lpb")("Ddate") = Ddate
	end if
	'取Cookies缓存数据'
	
	if DDN = "" then
		if 8 <= hour(now()) and hour(now()) <= 20 then
			DDN = "D"
			response.Cookies("Lpb")("DDN") = "D"
		else
			DDN = "N"
			response.Cookies("Lpb")("DDN") = "N"
		end if
	end if
	
	if Ddate = "" then
		Ddate = FormatTime(date(),2)
		response.Cookies("Lpb")("Ddate") = Ddate
	end if

	if DDN = "D" then
		Stime = cdate(Ddate & " 08:15:00")
		Etime = dateadd("h",12,Stime)
	else
		Stime = cdate(Ddate & " 20:15:00")
		Etime = dateadd("h",12,Stime)
	end if

'获取停机时间'
	Dlineno = ProductName & LineName
	SumDtime = 0
	HValue = 0
	set rs = Server.CreateObject("adodb.recordset")
	sql = "SELECT * FROM Ministat.dbo.BJLPBD2 where Dlineno = '"&Dlineno&"' and DD_N = '"&DDN&"' AND Hshift = '"&HShift&"' and Ddate = '"&Ddate&"'  ORDER BY Dtime DESC"
	rs.open sql,conn,1,1
	
		Set fso = CreateObject("Scripting.FileSystemObject")
		FileNo = ProductName&"L"&LineName&"DN"&DDN&"Sh"&HShift
		LpbXml = "Data\"&FileNo&"Column2D.xml" 'Output数据文件输出路径
		LpbXmlPath = server.mappath(LpbXml)
		if fso.fileexists(LpbXmlPath) then fso.DeleteFile(LpbXmlPath) '文件存在，就进行删除
	if not rs.eof then
		HValue = 1
		Set LpbXmlFile = fso.Createtextfile(server.mappath(LpbXml),true)
		LpbXmlFile.writeline "<graph caption='' baseFontSize='10' xAxisName='Lost type' yAxisName='Time lost(min)' decimalPrecision='0' formatNumberScale='0' rotateNames='1' showValues='1' yAxisMaxValue='50'>"
		TableHtml = "<table width='100%' border='1' cellspacing='0' cellpadding='0' style='font-size:12px;'>"
		TableHtml = TableHtml & "<tr><th>故障类型</th><th>设备</th><th>工位</th><th>问题</th><th>原因</th><th>解决方法</th><th>停机时间</th></tr>"
		
		do while not rs.eof
			id = rs("id")
			Dlossn = rs("Dlossn")
			DlossName = XlName(right(Dlossn,1))
			Ddate = FormatTime(rs("Ddate"),2)
			Dhtime = rs("Dhtime")
			Detime = rs("Detime")
			Dtime = rs("Dtime")
			Dequipment = rs("Dequipment")
			DPosition = rs("DPosition")
			Dproblem = rs("Dproblem")
			DReason = rs("DReason")
			Dsolution = rs("Dsolution")
			SumDtime = SumDtime + Dtime
			
				LpbXmlFile.writeline "<set name='" & Dlossn & "' value='" & Dtime & "' color='AFD8F8' />"
				TableHtml = TableHtml & "<tr><td>("&Dlossn&")"&DlossName&"</td><td>"&Dequipment&"</td><td>"&DPosition&"</td><td>"&DReason&"</td><td>"&Dproblem&"</td><td>"&Dsolution&"</td><td>"&Dtime&"</td></tr>"
			rs.movenext
		loop
		TableHtml = TableHtml & "</table>"
		LpbXmlFile.writeline "</graph>"
		LpbXmlFile.close
		
		'response.write TableHtml
	end if
	Set fso = nothing
	rs.close
	set rs = nothing
'获取停机时间'

'获取并确认是否有产出数据'
set rs = Server.CreateObject("adodb.recordset")
SqlStr = "SELECT TOP 1 ID,HID,Hlineno,HDate,HD_N,HShift,CONVERT(NVARCHAR(10),HSSV) HSSV,CONVERT(NVARCHAR(10),HLE) HLE,Hespeed,Hahour,Heqty,Haqty,HOEE ,Wima,Cell1,Cell2,Cell3,cell4,Cell21 FROM [Ministat].[dbo].[BJLPB] where Hlineno = '"&Dlineno&"' and HD_N = '"&DDN&"' AND Hshift = '"&HShift&"' and HDate = '"&Ddate&"'"
rs.open SqlStr,conn,1,1
if not rs.eof then
	HID = rs("HID")
	DDN = rs("HD_N")
	HShift = rs("HShift")
	HSSV = rs("HSSV")
	HLE = rs("HLE")
	Hahour = rs("Hahour")
	Haqty = rs("Haqty")
	HOEE = rs("HOEE")
	WimaFOR = rs("Wima")
	Cell1FOR = rs("Cell1")
	Cell2FOR = rs("Cell2")
	Cell3FOR = rs("Cell3")
	Cell4FOR = rs("Cell4")
	Cell5FOR = rs("Cell21")
else
	Hahour = 12
	Haqty = 0
	HOEE = 0
	WimaFOR = 0
	Cell1FOR = 0
	Cell2FOR = 0
	Cell3FOR = 0
	Cell4FOR = 0
	Cell5FOR = 0
	AddMenuName = "临时写入"
end if
rs.close
'获取并确认是否有产出数据'

'找OnlineData数量
SMCQTY = 0
OCCQTY = 0
TSCQTY = 0
MTCQTY = 0
SqlStr = "select sum(CASE when CType = 'MT' THEN CQTY END) MT,sum(CASE when CType = 'OC' THEN CQTY END) OC,sum(CASE when CType = 'SM' THEN CQTY END) SM,sum(CASE when CType = 'TS' THEN CQTY END) TS from [data-warehouse].[dbo].[OnlineData] where CDate = '"&Ddate&"' and CLine = '"&Dlineno&"' AND CShift = '"&HShift&"'-- and CShiftType = '"&DDN&"'"
rs.open SqlStr,conn2,1,1
if not rs.eof then
	SMCQTY = rs("SM")
	OCCQTY = rs("OC")
	TSCQTY = rs("TS")
	MTCQTY = rs("MT")
end if
rs.close

BackQty = 0
SendQty = 0
BQTY = 0
RQTY = 0
SqlStr = "SELECT TOP 1 * FROM Ministat.dbo.HGRecheck WHERE CType = 'Back' and CDate = '"&Ddate&"' and CLine = '"&Dlineno&"' AND CShift = '"&HShift&"'-- and CShiftType = '"&DDN&"'"
rs.open SqlStr,conn,1,1
if not rs.eof then BackQty = rs("CQTY")
rs.close

SqlStr = "SELECT TOP 1 * FROM Ministat.dbo.HGRecheck WHERE CType = 'Send' and CDate = '"&Ddate&"' and CLine = '"&Dlineno&"' AND CShift = '"&HShift&"'-- and CShiftType = '"&DDN&"'"
rs.open SqlStr,conn,1,1
if not rs.eof then SendQty = rs("CQTY")
rs.close

SqlStr = "SELECT TOP 1 * FROM Ministat.dbo.QABlock where CType = 'Block' and CDate = '"&Ddate&"' and CLine = '"&Dlineno&"' AND CShift = '"&HShift&"'-- and CShiftType = '"&DDN&"'"
rs.open SqlStr,conn,1,1
if not rs.eof then BQTY = rs("CQTY")
rs.close

SqlStr = "SELECT TOP 1 * FROM Ministat.dbo.QABlock where CType = 'Release' and CDate = '"&Ddate&"' and CLine = '"&Dlineno&"' AND CShift = '"&HShift&"'-- and CShiftType = '"&DDN&"'"
rs.open SqlStr,conn,1,1
if not rs.eof then RQTY = rs("CQTY")
rs.close

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
<div id="div0">
<center>
<%if Pr = "ok" then%>
<input onClick="document.all.WebBrowser.ExecWB(8,1)" type="button" value="页面设置" id="btn"/>
<input onClick="document.all.WebBrowser.ExecWB(7,1)" type="button" value="打印预览" id="btn"/>
<input type="button" value="关闭窗口" onClick="javascript:window.close()" id="btn"/>
<input onClick="document.all.WebBrowser.ExecWB(6,1)" type="button" value="打印" id="btn"/>
<%else%>
<form id="form1" name="form1" method="post" action="LpbReport.asp">

	<table width="640" border="0" cellspacing="0" cellpadding="0">
		<tr>
			<td>生产中心信息：</td>
		</tr>
		<tr>
			<td>
				<table width="100%" border="0" cellspacing="0" cellpadding="0">
					<tr>
						<td>产品：<select name="ProductName" id="TmProduct"></select></td>
						<td>线别：
						  <select name="LineName" id="TmLine"></select>
							<select name="CellName" id="TmCell" style="display:none;"></select>
							<script type="text/javascript">
								MenuInit('TmProduct', 'TmLine', 'TmCell', '<%=ProductName%>', '<%=LineName%>', '1');
							</script>
						</td>
						<td>日期：<input type="text" name="Ddate" class="Wdate" id="d122" onFocus="WdatePicker({isShowWeek:true,onpicked:function() {$dp.$('d122_1').value=$dp.cal.getP('W','W');$dp.$('d122_2').value=$dp.cal.getP('W','WW');}})" value="<%=Ddate%>" size="10" /></td>
						<td>D/N
						  <select name="DDN">
							<option value="D" <%if DDN = "D" then response.write "selected"%> >D</option>
							<option value="N" <%if DDN = "N" then response.write "selected"%> >N</option>
						</select></td>
						<td>班次
						  <select name="HShift" id="HShift">
							<option value="A" <%if HShift = "A" then response.write "selected"%> >A</option>
							<option value="B" <%if HShift = "B" then response.write "selected"%> >B</option>
							<option value="C" <%if HShift = "C" then response.write "selected"%> >C</option>
						</select></td>
						<td><input type="Submit" name="AddLpb" value="提交" /></td>
						<td><input type="button" value="打印" onClick="window.open('LpbReport.asp?ProductName=<%=ProductName%>&LineName=<%=LineName%>&Ddate=<%=Ddate%>&DDN=<%=DDN%>&HShift=<%=HShift%>&Pr=ok')"" /> </td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
</form>
<%end if%>
</center>
</div>

<div id="div1">
<center>
  <div id="List">
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
  <td width="25%" rowspan="2"><img src="../../LPB/images/Knowles.gif" width="119" height="34"
  alt="LPB"></td>
  <td width="75%" colspan="3">Everyone invovled in eliminating</td>
  <tr><td colspan="3">the loss of Overall Equipment Efficiency</td></tr>
  <tr><td colspan="4">　</td></tr>
  </table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="25%">Center：<%=ProductName%></td>
      <td width="25%">LineNo：<%=LineName%></td>
      <td width="25%">D/N：<%=DDN%></td>
      <td width="25%">Date：<%=Ddate%></td>
    </tr>
    <tr>
      <td>Shift：<%=HShift%></td>
      <td>SSV：<%=HSSV%></td>
      <td>LE：<%=HLE%></td>
      <td>生产时间：</td>
    </tr>
    <tr>
      <td>理论产量：</td>
      <td>实际产量：<%=Haqty%></td>
      <td>百分比1：</td>
      <td>百分比2：</td>
    </tr>
    <tr><td colspan="4">　</td></tr>
  </table>

<table width="100%" border="0" cellpadding="0" cellspacing="0">
<tr>
<td width="46%">
	<fieldset id="nav1">
	<legend>在线监测结果:</legend>
		<table width="100%" border="0" cellspacing="0" cellpadding="0">
		  <tr>
			<td width="33%" align="right">采样：</td>
			<td width="16%" align="left"><%=SMCQTY%></td>
			<td width="32%" align="right">目测废品：</td>
			<td width="19%" align="left"><%=OCCQTY%></td>
		  </tr>
		  <tr>
			<td align="right">清零废品：</td>
			<td align="left"><%=TSCQTY%></td>
			<td align="right">废膜/顶片：</td>
			<td align="left"><%=MTCQTY%></td>
		  </tr>
		</table>
	</fieldset>
</td>
<td width="24%">
	<fieldset id="nav1">
	<legend>质量封存与放行:</legend>
		<table width="100%" border="0" cellspacing="0" cellpadding="0">
		  <tr>
			<td width="50%" align="right">封存：</td>
			<td width="50%" align="left"><%=BQTY%></td>
		  </tr>
		  <tr>
			<td align="right">放行：</td>
			<td align="left"><%=RQTY%></td>
		  </tr>
		</table>
	</fieldset>
</td>
<td width="30%">
	<fieldset id="nav1">
	<legend>手工组复检:</legend>
		<table width="100%" border="0" cellspacing="0" cellpadding="0">
		  <tr>
			<td width="59%" align="right">送交手工组：</td>
			<td width="41%" align="left"><%=SendQty%></td>
		  </tr>
		  <tr>
			<td align="right">手工组返回：</td>
			<td align="left"><%=BackQty%></td>
		  </tr>
		</table>
	</fieldset>
</td>
</tr>
</table>

	<div id="nav4">
		<%if HValue = 1 then%>
		<%=TableHtml%>
		</br>
		<div id="chartdiv1" align="center">FusionCharts. </div>
		<script type="text/javascript">
			var chart = new FusionCharts("../Charts/FCF_Column2D.swf", "ChartId2", "640", "400");
			chart.setDataURL("Data/<%=FileNo%>Column2D.xml");		   
			chart.render("chartdiv1");
		</script>
		<%
		else
			response.write "没有设备维护记录"
		end if
		%>
	</div>
</div>
</center>
</div>

</body>
</html> 
