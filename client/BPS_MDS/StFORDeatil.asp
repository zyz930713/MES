<!--#include file="conn.asp"-->
<!--#include virtual="/include/function.asp" -->
<!--#include virtual="/include/Charts/FusionCharts.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>StationDetailFOR</title>
<script language="JavaScript" type="text/javascript" src="../include/Charts/FusionCharts.js"></script>
</head>
<body>
<center>
<%
	productName = request("productName")
	LineName = request("LineName")
	StID = request("StID")
	LR = request("LR")
	Edatetime = request("Edatetime")
	SetResteTime = request("SetResteTime")

	if SetResteTime = "复位" then
		set Rs1 = Server.CreateObject("adodb.recordset")
		Sql1 = "select * from [dbo].[PrResteTime] where productName = '"&productName&"' and LineName = '"&LineName&"' and StID = '"&StID&"' and PrPosition = '"&LR&"'"
		Rs1.open Sql1,conn,1,1
		if not Rs1.eof then
			sql = "update [dbo].[PrResteTime] set ResteTime = getdate() where productName = '"&productName&"' and LineName = '"&LineName&"' and StID = '"&StID&"' and PrPosition = '"&LR&"'"
		else
			sql = "insert into [dbo].[PrResteTime] values('"&productName&"','"&LineName&"','"&StID&"','"&LR&"',getdate())"
		end if
		Rs1.close
		set Rs1 = nothing
		' response.write sql
		' response.end
		' conn.execute(sql)
		response.write "<script type='text/javascript'>"
		response.write "window.opener=null;"
		response.write "window.open('','_self');"
		response.write "window.close();"
		response.write "</script>"
	end if
	
	set Rs = Server.CreateObject("adodb.recordset")
	Sql = "EXEC [dbo].[GetStationFORDetail] @productName = N'"&productName&"',@LineName = N'"&LineName&"',@StID = N'"&StID&"',@PrPosition = '"&LR&"',@EndTime = N'"&Edatetime&"'"
	'response.write Sql
	Rs.open sql,Conn,1,1
	if not rs.eof then
		do while not Rs.eof
			BadSum = rs("BadSum")
			PrSum = rs("PrSum")
			StFOR = rs("StFOR")
			PrBad1Xml = PrBad1Xml & "<set value='" & StFOR & "' name='" & BadSum & "/" & PrSum & "' color='FF0000'/>"
		Rs.movenext
		loop
		ChartXml = ChartXml & "<graph Caption='Deatil FOR' plotFillAlpha='100' baseFontSize='10' numberSuffix='%25' decimalPrecision='2' formatNumberScale='0' rotateNames='0' showValues='1'>"
		ChartXml = ChartXml & PrBad1Xml
		ChartXml = ChartXml & "</graph>"
		Call renderChart("../include/charts/Line.swf?ChartNoDataText=No data to display.&amp;#xA;", "", ChartXml, "myNext0"&i&"", "550", "400", false,true)
	else
		response.write "没有找到所需数据"
	end if
	Rs.close
	set Rs = nothing
%>
<form name="form1" method="post" action="StFORDeatil.asp" >
<input type="hidden" name="productName" value="<%=productName%>" />
<input type="hidden" name="LineName" value="<%=LineName%>" />
<input type="hidden" name="StID" value="<%=StID%>" />
<input type="hidden" name="LR" value="<%=LR%>" />
<input type="hidden" name="Edatetime" value="<%=Edatetime%>" />
<input type="submit" name="SetResteTime" value="复位"/>
<input type="button" name="CloseWinForm" value="关闭" onClick="window.close();"/>
</form>
</center>
</body>
</html>
