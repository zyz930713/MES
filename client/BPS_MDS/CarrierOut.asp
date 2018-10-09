<!--#include file="conn.asp"-->
<!--#include virtual="/include/function.asp" -->
<!--#include virtual="/include/Charts/FusionCharts.asp" -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="refresh" content="300" charset="utf-8" />
<title>StationFOR</title>
<script language="JavaScript" type="text/javascript" src="../include/Charts/FusionCharts.js"></script>
</head>
<body style="margin:20px;font-size:14px;background-color:#FFFFFF;">
<center>
<%
	productName = request("productName")
	LineName = request("LineName")
	'取Cookies缓存数据'
	if productName = "" then
		productName = request.Cookies("mds")("productName")
	else
		response.Cookies("mds")("productName") = productName
	end if
	productName = "Marigold"
	
	if LineName = "" then
		LineName = request.Cookies("mds")("LineName")
	else
		response.Cookies("mds")("LineName") = LineName
	end if
	
	if Itime = "" then
		Itime = request.Cookies("mds")("Itime")
	else
		response.Cookies("mds")("Itime") = Itime
	end if
	'取Cookies缓存数据'
	

	Itime = request("Itime")
	if Itime = "" then Itime = 1
	Edatetime = now()
	Sdatetime = dateadd("h",(0-Itime),Edatetime)
	'response.write Sdatetime & " - " & Edatetime

%>
<form name="form1" method="post" action="CarrierOut.asp" >
线号：<select name="LineName">
				<option value="AF01" <%if LineName = "AF01" then response.write "selected"%> >AF01</option>
				<option value="AF02" <%if LineName = "AF02" then response.write "selected"%> >AF02</option>
				<option value="AF03" <%if LineName = "AF03" then response.write "selected"%> >AF03</option>
				<option value="AF04" <%if LineName = "AF04" then response.write "selected"%> >AF04</option>
				<option value="AF05" <%if LineName = "AF05" then response.write "selected"%> >AF05</option>
			</select>
　间隔：<select name="Itime">
        <option value="1" <%if Itime = "1" then response.write "selected"%>> 1 </option>
        <option value="2" <%if Itime = "2" then response.write "selected"%>> 2 </option>
        <option value="3" <%if Itime = "3" then response.write "selected"%>> 3 </option>
        <option value="4" <%if Itime = "4" then response.write "selected"%>> 4 </option>
        <option value="5" <%if Itime = "5" then response.write "selected"%>> 5 </option>
        <option value="6" <%if Itime = "6" then response.write "selected"%>> 6 </option>
        <option value="7" <%if Itime = "7" then response.write "selected"%>> 7 </option>
        <option value="8" <%if Itime = "8" then response.write "selected"%>> 8 </option>
        <option value="9" <%if Itime = "9" then response.write "selected"%>> 9 </option>
        <option value="10" <%if Itime = "10" then response.write "selected"%>> 10 </option>
        <option value="11" <%if Itime = "11" then response.write "selected"%>> 11 </option>
        <option value="12" <%if Itime = "12" then response.write "selected"%>> 12 </option>
    </select>小时
　<input type="submit" name="refresh" value="刷新" />
</form>
<%
Response.buffer=true '开启缓存
Response.Write "<div id='loading' style='width:960px;height:480px;background:#FFFFFF url(images/loading.gif) center no-repeat;'>" '先写入提示
Response.Write "</div>"
Response.flush()
Response.clear()
set Rs = Server.CreateObject("adodb.recordset")
for C = 1 to 2
ChartXml = ""
categoriesXml = ""
datasetXml = ""
	Sql = "EXEC [dbo].[GetCarrierFOR] @productName = N'"&productName&"',@LineName = N'"&LineName&"',@StartTime = N'"&Sdatetime&"',@EndTime = N'"&Edatetime&"',@CarrierID = "&C&""
	'response.write Sql
	Rs.open sql,Conn,1,1
	if not rs.eof then
		ChartXml = ChartXml & "<graph useRoundEdges='0' formatNumberScale='0' numberPrefix='' numDivLines='3' rotateNames='0' numberSuffix='%25' decimalPrecision='2' caption='Carrier - "&C&"' yAxisName='FOR' xAxisName='CarrierNr'>"
		i = 1
		while not rs.eof
			CarrierFOR = rs("CarrierFOR")
			if CarrierFOR = 0 then CarrierFOR = ""
			
			if rs("ColName") = "0" then
				categoriesXml = categoriesXml & "<category showValues='0' name='" & rs("Value") & " - " & rs("FailSum") & "/" & rs("StFailSum") & " FOR:" & CarrierFOR & "%'/>"
			else
				StColor = rs("StColor")
				if i = 1 then datasetXml = datasetXml & "<dataset showValues='1' seriesName='"&rs("StName")&"' color='"&StColor&"'>"
					datasetXml = datasetXml & "<set value='"&CarrierFOR&"'/>"
				if i = 8 then datasetXml = datasetXml & "</dataset>"
				i = i + 1
			end if
			if i = 9 then i = 1
		rs.movenext
		wend
		ChartXml = ChartXml & "<categories>"
		ChartXml = ChartXml & categoriesXml
		ChartXml = ChartXml & "</categories>"
		ChartXml = ChartXml & datasetXml
		ChartXml = ChartXml & "</graph>"
		Call renderChart("../include/charts/StackedColumn2D.swf?ChartNoDataText=No data to display.&amp;#xA;", "", ChartXml, "myNext0"&C&"", "1024", "360", false,true)
		response.write "<br/>"
	end if
	Rs.close
next
set rs = nothing
response.write "<script>"
response.write "var oLoading=document.getElementById('loading');"
response.write "var parent=oLoading.parentNode;"
response.write "parent.removeChild(oLoading);"
response.write "</script>"
%>
</center>
</body>
</html>
