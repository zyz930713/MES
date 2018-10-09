<!--#include file= "conn.asp"-->
<!--#include file= "../include/function.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link rel="stylesheet" type="text/css" href="../../styles/basic.css" />
<script language="javascript" type="text/javascript" src="../include/DatePicker/WdatePicker.js"></script>
<title>Daily Reporting</title>
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
font:normal 16px/16px "Arial";
}
table td{
font:normal 14px/16px "Arial";
border-color:black;
}
a{
color:#fff;
text-decoration:none;
}
</style>
</head>
<body>
<%
if HDate = "" then
	HDate = FormatTime((date()-1),2)
end if
%>
</br>
<center>
<form method="post" action="">
项目:<select name="ProductName">
		<option value="RA" <%if ProductName = "RA" then response.write "selected"%> >Ra</option>
		<option value="Donau Slim" <%if ProductName = "Donau Slim" then response.write "selected"%> >Donau Slim</option>
		<option value="PETRA" <%if ProductName = "PETRA" then response.write "selected"%> >Petra</option>
	</select>&nbsp;
日期:<input type="text" name="HDate" class="Wdate" id="d122" onFocus="WdatePicker({isShowWeek:true,onpicked:function() {$dp.$('d122_1').value=$dp.cal.getP('W','W');$dp.$('d122_2').value=$dp.cal.getP('W','WW');}})" value="<%=HDate%>" size="10" />&nbsp;
<input type="submit" name="Getdata" value="提交" />
</form>
<center>

<%
HDate = request("HDate")
ProductName = request("ProductName")
Getdata = request("Getdata")
if Getdata = "提交" or Getdata = "ok" then
	Startdate = Cdate(HDate)
	if ProductName = "RA" then
		TitleNmae = "RA 11x15 Speaker (Danubius)"
	else
		TitleNmae = "Petra"
	end if
	EndDate = dateadd("d",1,Startdate)
	set Rs = Server.CreateObject("adodb.recordset")
	Sql = "EXEC [DWH].[dbo].[Create_Report_Data_"&ProductName&"] @StartDate = N'"&Startdate&"',@EndDate = N'"&EndDate&"'"
	Rs.open sql,conn,1,1
	%>
	
	<table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#7BB1DB">
	  <tr>
		<td width="20%">&nbsp;</td>
		<td width="60%" rowspan="2"><p style="font-size:20px;color:#FFF;text-align:center;"><%=TitleNmae%></p></td>
		<td width="20%" colspan="3">&nbsp;</td>
	  </tr>
	  <tr>
		<td style="color:#fff;text-align:center;"><a href="DailyReport.asp?Getdata=ok&ProductName=<%=ProductName%>&HDate=<%=dateadd("d",-1,Startdate)%>"><<</a>　<%=Startdate%>　<a href="DailyReport.asp?Getdata=ok&ProductName=<%=ProductName%>&HDate=<%=dateadd("d",1,Startdate)%>">>></a></td>
		<td width="3%" rowspan="2">&nbsp;</td>
		<td width="14%" rowspan="2"><img src="Logo.jpg" width="124" height="52" align="middle"></td>
		<td width="3%" rowspan="2">&nbsp;</td>
	  </tr>
	  <tr>
		<td style="color:#fff;text-align:center;">&nbsp;</td>
		<td><p style="font-size:14px;color:#FFF;text-align:center;">Daily Report</p></td>
	  </tr>
	  <tr>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td colspan="3">&nbsp;</td>
	  </tr>
	</table>
	</br>
	
	<table width="100%" border="1" cellpadding="0" cellspacing="0">
	<tr bgcolor='#B9E9FF'>
	<th>Prod</br>Line</th><th>Shift</th><th>Prod</br>Time</th><th>Down</br>Time</th><th>Target</br>CC</th>
	<%if ProductName = "RA" then%>
		<th>Output</br>Cell 5</th>
		<th>FOR</br>Cell 5</th>
		<th>FOR</br>Cell 4</th>
	<%else%>
		<th>Output</br>Cell 3</th>
	<%end if%>
	<th>FOR</br>Cell 3</th>
	<th>FOR</br>Cell 2</th>
	<th>FOR</br>Cell 1</th>
	<th>FOR</br>Transfer</th>
	<%if ProductName = "RA" then%>
		<th>Deviat.</br>Cell 5</th>
		<th>Weekly</br>Quantity</br>Cell 5</th>
		<th>Average</br>Output</br>Cell 5</th>
		<th>Preview</br>Week</br>Cell 5</th>
	<%ELSE%>
		<th>Deviat.</br>Cell 3</th>
		<th>Weekly</br>Quantity</br>Cell 3</th>
		<th>Average</br>Output</br>Cell 3</th>
		<th>Preview</br>Week</br>Cell 3</th>
	<%end if%>
	<th>OEE</th><th>TEEP</th></tr>
	<%
	remarks = ""
	set RsL = Server.CreateObject("adodb.recordset")
	SqlL = "select productname from report_hv_data where adate = '"&Startdate&"' and ProductName like '"&ProductName&"' GROUP BY Line_Num"
	RsL.open SqlL,conn,1,1
	do while not RsL.eof
		ProdTimeT = 0
		DownTimeT = 0
		TargetCCT = 0
		COGT = 0
		C1fT = 0
		C2fT = 0
		C3fT = 0
		C4fT = 0
		C5fT = 0
		FsfT = 0
		LcwT = 0
		CowT = 0
		OeeT = 0
		TeepT = 0
		LoopNum = 0
		'--------------------------------------------------------'
		set rs = Server.CreateObject("adodb.recordset")
		sql = "SELECT PRODUCT,Line_Num,ShiftAcronym,ProductionTime/60 ProdTime,isnull(ScheduledStopTime/60,0) DownTime,Target_OutputCC/720*ProductionTime TargetCC,isnull(Cell3_PiecesGood,0) C3G,isnull(Cell5_PiecesGood,0) C5G,[Target_FOR_Cell1] C1T,[Target_FOR_Cell2] C2T,[Target_FOR_Cell3] C3T,isnull([Target_FOR_Cell4],0) C4T,isnull([Target_FOR_Cell5],0) C5T,FOR_Transfer1_Target FsT,isnull(Cell5_PiecesBad/(Cell5_PiecesGood+Cell5_PiecesBad)*100,0) C5F,isnull(Cell4_PiecesBad/(Cell4_PiecesGood+Cell4_PiecesBad)*100,0) C4F,Cell3_PiecesBad/(Cell3_PiecesGood+Cell3_PiecesBad)*100 C3F,Cell2_PiecesBad/(Cell2_PiecesGood+Cell2_PiecesBad)*100 C2F,Cell1_PiecesBad/(Cell1_PiecesGood+Cell1_PiecesBad)*100 C1F,Transfer1_PiecesBad/(Transfer1_PiecesGood+Transfer1_PiecesBad)*100 FsF,LastCellWeekSum,C5w = isnull((SELECT TOP 1 SUM([Cell5_PiecesGood]) FROM [DWH].[dbo].[Report_Day_RA] AS RDR WHERE RDR.[Day] > = dateadd(d,-6,'"&Startdate&"') AND RDR.[day] < = '"&Startdate&"' AND RDR.ShiftAcronym = RD.ShiftAcronym AND RDR.ProdlineNr = RD.Line_Num),0),C3w = isnull((SELECT TOP 1 SUM([Cell3_PiecesGood]) FROM [DWH].[dbo].[Report_Day_PETRA] AS RDP WHERE RDP.[Day] > = dateadd(d,-6,'"&Startdate&"') AND RDP.[day] < = '"&Startdate&"' AND RDP.ShiftAcronym = RD.ShiftAcronym AND RDP.ProdlineNr = RD.Line_Num),0) FROM [DWH].[dbo].[Reportdata] as Rd WHERE PRODUCT = '"&ProductName&"' and [Day] = '"&Startdate&"' and Line_Num = '"&RsL("Line_num")&"' ORDER BY Line_num,ShiftAcronym desc"
		rs.open sql,conn,1,1
		do while not rs.eof
		PRODUCT = rs("PRODUCT")
		Line_num = rs("Line_num")
		ShiftAcronym = rs("ShiftAcronym")
		
			'-------------输出备注-------------'
			set rsa = Server.CreateObject("adodb.recordset")
			sqla = "SELECT dbo.DWRemark1.Remark FROM dbo.DWRemark1 INNER JOIN dbo.R_Prodline ON dbo.DWRemark1.RProdlineID = dbo.R_Prodline.RProdlineID INNER JOIN dbo.MD_KST ON dbo.R_Prodline.MDKSTID = dbo.MD_KST.MDKSTID INNER JOIN dbo.MD_Shift ON dbo.MD_Shift.MDShiftID = dbo.DWRemark1.MDShiftID WHERE (dbo.MD_KST.KSTAcronym = '"&PRODUCT&"') AND (dbo.R_Prodline.ProdlineNr = '"&Line_num&"') AND (dbo.MD_Shift.ShiftAcronym = '"&ShiftAcronym&"') AND ('"&Startdate&"' <= dbo.DWRemark1.ShiftBegin) AND (dbo.DWRemark1.ShiftBegin <= DATEADD(day, 1, '"&Startdate&"'))"
			rsa.open sqla,conn,1,1
			if not rsa.eof then
				Rm = rsa("Remark")
				Rm = replace(Rm,"Shift","Shift　")
				Rm = replace(Rm,"shift","Shift　")
				Rm = replace(replace(Rm,"（",""),"）","")
				Rm = replace(replace(Rm,"(",""),")","")
				Rm = replace(replace(Rm,"D","Day"),"Dayay","Day</br>")
				Rm = replace(replace(Rm,"N","Night"),"Nightight","Night</br>")
				Rm = replace(Rm,"小时","小时</br>")
				Rm = replace(Rm,"分钟.","分钟.</br></br>")
				Rm = replace(replace(Rm,"mon)","min"),"min","min</br></br>")
				Rm = replace(Rm,"计划","</br>计划")
				Rm = replace(Rm,"问题","</br>问题")
				Rm = replace(Rm,"原因","</br>原因")
				Rm = replace(Rm,"解决","</br>解决")
				remarks = remarks & "<tr><td align='left'>" & Rm & "</td></tr>"
			end if
			rsa.close
			set rsa = nothing
			'-------------输出备注-------------'
		
		Shift = replace(ShiftAcronym,"T","Day")
		Shift = replace(Shift,"N","Night")
		ProdTime = Formatnumber(rs("ProdTime"),1,-1)
		DownTime = Formatnumber(rs("DownTime"),1,-1)
		TargetCC = Formatnumber(rs("TargetCC"),0,-1)
		if Shift = "Night" then
			TrStyle = "style='background:#EBEBEB;text-align:center;'"
		else
			TrStyle = "style='background:#ffffff;text-align:center;'"
		end if
		C1F = csng(rs("C1F"))
		if C1F > csng(rs("C1T")) then
			C1FStr = "<font color='red'>" & Formatnumber(C1F,2) & "%</font>"
		else
			C1FStr = "<font color='green'>" & Formatnumber(C1F,2) & "%</font>"
		end if
		
		C2F = csng(rs("C2F"))
		if C2F > csng(rs("C2T")) then
			C2FStr = "<font color='red'>" & Formatnumber(C2F,2) & "%</font>"
		else
			C2FStr = "<font color='green'>" & Formatnumber(C2F,2) & "%</font>"
		end if
		
		C3F = csng(rs("C3F"))
		if C3F > csng(rs("C3T")) then
			C3FStr = "<font color='red'>" & Formatnumber(C3F,2) & "%</font>"
		else
			C3FStr = "<font color='green'>" & Formatnumber(C3F,2) & "%</font>"
		end if
		
		FsF = csng(rs("FsF"))
			FsFStr = "" & Formatnumber(FsF,2) & "%"
		
		LastCellWeekSum = Formatnumber(rs("LastCellWeekSum"),0)
		DayNo = Weekday(Startdate,vbMonday) 
		LastCellWeekAvg = Formatnumber(LastCellWeekSum / DayNo,0)

		if ProductName = "RA" then
			COG = Formatnumber(rs("C5G"),0,-1)
			Cow = Formatnumber(rs("C5w"),0)
			Oee = Formatnumber(COG / (ProdTime * 60),2,-1)
			Teep = Formatnumber(COG / 720,2,-1)
			C4F = csng(rs("C4F"))
			C4fT = C4fT + C4F
			if C4F > csng(rs("C4T")) then
				C4FStr = "<font color='red'>" & Formatnumber(C4F,2) & "%</font>"
			else
				C4FStr = "<font color='green'>" & Formatnumber(C4F,2) & "%</font>"
			end if
		
			C5F = csng(rs("C5F"))
			C5fT = C5fT + C5F
			if C5F > csng(rs("C5T")) then
				C5FStr = "<font color='red'>" & Formatnumber(C5F,2) & "%</font>"
			else
				C5FStr = "<font color='green'>" & Formatnumber(C5F,2) & "%</font>"
			end if
		else
			COG = Formatnumber(rs("C3G"),0,-1)
			Cow = Formatnumber(rs("C3w"),0)
			Oee = Formatnumber(COG / (ProdTime * 60),2,-1)
			Teep = Formatnumber(COG / 720,2,-1)
		end if
		ProdTimeT = ProdTimeT + ProdTime
		DownTimeT = DownTimeT + DownTime
		TargetCCT = TargetCCT + TargetCC
		LcwT = LcwT + LastCellWeekSum
		COGT = COGT + COG
		CowT = CowT + Cow
		C1fT = C1fT + C1F
		C2fT = C2fT + C2F
		C3fT = C3fT + C3F
		FsfT = FsfT + FsF
		OeeT = OeeT + Oee
		TeepT = TeepT + Teep
		%>
		<tr <%=TrStyle%>><td>Line <%=Line_num%></td><td><%=Shift%></td><td><%=ProdTime%>h</td><td><%=DownTime%>h</td><td><%=TargetCC%></td><td><%=COG%></td>
		<%if ProductName = "RA" then%>
		<td><%=C5FStr%>%</td>
		<td><%=C4FStr%>%</td>
		<%end if%>
		<td><%=C3FStr%>%</td><td><%=C2FStr%></td><td><%=C1FStr%></td><td><%=FsFStr%></td><td>第<%=DayNo%>天</td><td><%=LastCellWeekSum%></td><td><%=LastCellWeekAvg%></td><td><%=Cow%></td><td><%=Oee%>%</td><td><%=Teep%>%</td></tr>
		<%
		rs.movenext
		LoopNum = LoopNum + 1
		loop
		rs.close
		set rs = nothing
		'--------------------------------------------------------'
	RsL.movenext
	%>
		<tr style="background:#FAD97D;font-size:15px;text-align:center;"><td colspan="2">Line <%=Line_num%> Total</td><td><%=ProdTimeT%>h</td><td><%=DownTimeT%>h</td><td><%=Formatnumber(TargetCCT,0)%></td><td><%=Formatnumber(COGT,0)%></td>
		<%if ProductName = "RA" then%>
		<td><%=Formatnumber(C5fT/LoopNum,2)%>%</td>
		<td><%=Formatnumber(C4fT/LoopNum,2)%>%</td>
		<%end if%>
		<td><%=Formatnumber(C3fT/LoopNum,2)%>%</td><td><%=Formatnumber(C2fT/LoopNum,2)%>%</td><td><%=Formatnumber(C1fT/LoopNum,2)%>%</td><td><%=Formatnumber(FsfT/LoopNum,2)%>%</td><td>第<%=DayNo%>天</td><td><%=Formatnumber(LcwT,0)%></td><td><%=Formatnumber(LastCellWeekAvg,0)%></td><td><%=Formatnumber(CowT,0)%></td><td><%=Formatnumber(OeeT/LoopNum,2)%>%</td><td><%=Formatnumber(TeepT/LoopNum,2)%>%</td></tr>
		<%
	loop
	
	response.write "</table></br>"
	response.write "<table width='100%' border='1' cellpadding='0' cellspacing='0'>"
	response.write "<tr style='background:#C9C9C9;font-size:15px;'><th>Remarks:</th></tr>"
	response.write "<tr><td>"&remarks&"</td></tr>"
	response.write "</table></br>"
end if
%>

</table>
<!--#include file="../footer.asp"-->
</body>
</html>
