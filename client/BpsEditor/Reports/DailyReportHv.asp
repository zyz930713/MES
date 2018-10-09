<!--#include virtual="/Functions/function.asp" -->
<!--#include virtual="/BpsEditor/Conn_open.asp" -->
<!--#include virtual="/Functions/ChangeLineName.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link rel="stylesheet" type="text/css" href="../../styles/basic.css" />
<script language="javascript" type="text/javascript" src="../include/ThreeMenu.js"></script>
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
a{
color:#fff;
text-decoration:none;
}
</style>
</head>
<body>
<%
Adate = request("Adate")
ProductName = request("ProductName")
Getdata = request("Getdata")

if Adate = "" then
	Adate = FormatTime((date()-1),2)
end if

WeekFirstDay = DateAdd("d",0 - Weekday(Adate,2) + 1, CDate(Adate))
'response.end
%>
<br>
<center>
<form method="post" action="DailyReportHv.asp">
产品：<select name="ProductName" id="TmProduct"></select>
		<select name="LineName" id="TmLine" style="display:none" ></select>
		<select name="CellName" id="TmCell" style="display:none" ></select>
		<script type="text/javascript">
			MenuInit('TmProduct','TmLine','TmCell','<%=ProductName%>','<%=LineName%>','1');
		</script>
　日期：<input type="text" name="Adate" class="Wdate" id="d122" onFocus="WdatePicker({isShowWeek:true,onpicked:function() {$dp.$('d122_1').value=$dp.cal.getP('W','W');$dp.$('d122_2').value=$dp.cal.getP('W','WW');}})" value="<%=Adate%>" size="10" />
　<input type="submit" name="Getdata" value="刷新" />
</form>
<center>
<%
if Getdata = "刷新" or Getdata = "ok" then
	Startdate = Cdate(Adate)
	EndDate = Startdate
%>
<div id='page1'>
	<table width="1020" border="0" cellspacing="0" cellpadding="0" bgcolor="#7BB1DB">
	  <tr>
		<td width="20%">&nbsp;</td>
		<td width="60%" rowspan="2"><p style="font-size:20px;color:#FFF;text-align:center;"><strong><%=ProductName%></strong></p></td>
		<td width="20%" colspan="3">&nbsp;</td>
	  </tr>
	  <tr>
		<td style="color:#fff;text-align:center;"><a href="DailyReportHv.asp?Getdata=ok&ProductName=<%=ProductName%>&Adate=<%=dateadd("d",-1,Startdate)%>"><<</a>　<%=Startdate%>　<a href="DailyReportHv.asp?Getdata=ok&ProductName=<%=ProductName%>&Adate=<%=dateadd("d",1,Startdate)%>">>></a></td>
		<td width="3%" rowspan="2">&nbsp;</td>
		<td width="14%" rowspan="2"><img src="../Images/Logo.jpg" width="124" height="52" align="middle"></td>
		<td width="3%" rowspan="2">&nbsp;</td>
	  </tr>
	  <tr>
		<td style="color:#fff;text-align:center;"><%=FormatTime(Startdate,9)%></td>
		<td><p style="font-size:20px;color:#FFF;text-align:center;">Daily Report</p></td>
	  </tr>
	  <tr>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td colspan="3">&nbsp;</td>
	  </tr>
	</table>
	
	<table width="1020" border="1" cellpadding="0" cellspacing="0">
	<tr style="background:#B9E9FF;font-width:bold;">
	<th>Prod<br>Line</th>
	<th>DAY<br>Night</th>
	<th>Prod<br>Time</th>
	<th>Down<br>Time</th>
	<th>Target<br>CC</th>
	<%if ProductName = "RA" or ProductName = "Slim" then%>
		<th>Weekly<br>Quantity<br>Cell 5</th>
		<th>Output<br>Cell 5</th>
		<th>FOR<br>Cell 5</th>
		<th>FOR<br>Cell 4</th>
	<%else%>
		<th>Weekly<br>Quantity<br>Cell 3</th>
		<th>Output<br>Cell 3</th>
	<%end if%>
	<th>FOR<br>Cell 3</th>
	<th>FOR<br>Cell 2</th>
	<th>FOR<br>Cell 1</th>
	<th>FOR<br>Transfer</th>
	<th>OEE</th>
	<th>TEEP</th></tr>
	<%
	dim RemarkStr,RemarkStrSum
	RemarkStr = ""
	RemarkStrSum = ""
	
	set RsL = Server.CreateObject("adodb.recordset")
	SqlL = "select Linename from report_hv_data where '"&Startdate&"' < = adate and adate < = '"&EndDate&"' and ProductName like '"&ProductName&"' GROUP BY Linename order by Linename"
	'response.write SqlL & "<br>"
	RsL.open SqlL,conn,1,1
	do while not RsL.eof
		ProdTimeT = 0
		DownTimeT = 0
		TargetCCT = 0
		COWT = 0
		COGT = 0
		T1FT = 0
		C1FT = 0
		C2FT = 0
		C3FT = 0
		C4FT = 0
		C5FT = 0
		OeeT = 0
		TeepT = 0
		LoopNum = 0
		'--------------------------------------------------------'
		LineName = RsL("Linename")
		'取得Target
		set TarRs = server.createobject("adodb.recordset")
			OraTarSql = "select OutPCS,nvl(T1,0) T1,NVL(C1,0) C1,NVL(C2,0) C2,NVL(C3,0) C3,NVL(C4,0) C4,NVL(C5,0) C5,OEE from report_hv_target"
			OraTarSql = OraTarSql & " WHERE ProductName = '"&ProductName&"' AND LineName = '"&LineName&"' and Tdate < = '"&Startdate&"' ORDER BY Tdate desc"
			'response.write PRODUCTNAME & " - " & LINENAME & "<br>"
			'response.write OraTarSql & "<br>"
			TarRs.open OraTarSql,conn,1,1
			TargetCC = 0
			T1T = 0
			C1T = 0
			C2T = 0
			C3T = 0
			C4T = 0
			C5T = 0
			OEE = 0
			if not TarRs.eof then
				TargetCC = csng(TarRs("OutPCS"))
				T1T = CSNG(TarRs("T1"))
				C1T = CSNG(TarRs("C1"))
				C2T = CSNG(TarRs("C2"))
				C3T = CSNG(TarRs("C3"))
				C4T = CSNG(TarRs("C4"))
				C5T = CSNG(TarRs("C5"))
				OEE = TarRs("OEE")
			end if
			TarRs.close
			Set TarRs = nothing
		SET LoopRs = server.createobject("adodb.recordset")
		OraLoopSql = "SELECT ADATE,PRODUCTNAME,LINENAME,DNAME,ShiftName,NVL(OPENTIME,0) OPENTIME,NVL(STOPTIME,0) STOPTIME,NVL(T1G,0) T1G,NVL(T1B,0) T1B,NVL(C1G,0) C1G,NVL(C1B,0) C1B,NVL(C2G,0) C2G,NVL(C2B,0) C2B,NVL(C3G,0) C3G,NVL(C3B,0) C3B,NVL(C4G,0) C4G,NVL(C4B,0) C4B,NVL(C5G,0) C5G,NVL(C5B,0) C5B"
		OraLoopSql = OraLoopSql & ",NVL((select sum(c3g) from report_hv_data where '"&WeekFirstDay&"' < = adate and adate < = '"&EndDate&"' AND ProductName = '"&ProductName&"' AND LINENAME = '"&LineName&"' AND dname = rhd.dname),0) C3GW"
		OraLoopSql = OraLoopSql & ",NVL((select sum(c5g) from report_hv_data where '"&WeekFirstDay&"' < = adate and adate < = '"&EndDate&"' AND ProductName = '"&ProductName&"' AND LINENAME = '"&LineName&"' AND dname = rhd.dname),0) C5GW FROM report_hv_data RHD"
		OraLoopSql = OraLoopSql & " where '"&Startdate&"' < = adate and adate < = '"&EndDate&"' AND ProductName = '"&ProductName&"' AND LINENAME = '"&LineName&"' ORDER BY LINENAME,DNAME"
		'response.write OraLoopSql & "<br>"
		LoopRs.open OraLoopSql,conn,1,1
		do while not LoopRs.eof
			'开始处理数据
			ADATE = LoopRs("ADATE")
			DNAME = LoopRs("DNAME")
			ShiftName = LoopRs("ShiftName")
			OPENTIME = CSNG(LoopRs("OPENTIME")) / 60
			STOPTIME = CSNG(LoopRs("STOPTIME")) / 60
			
			if DNAME = "D" then
				Shift = "Day"
				TrStyle = "style='background:#FFFFFF;text-align:center;'"
			else
				Shift = "Night"
				TrStyle = "style='background:#EBEBEB;text-align:center;'"
			end if
			
			T1G = CSNG(LoopRs("T1G"))
			T1B = CSNG(LoopRs("T1B"))
			T1A = T1G + T1B
			IF T1A = 0 THEN
				T1F = 0
			ELSE
				T1F = T1B / T1A * 100
			END IF
			
			C1G = CSNG(LoopRs("C1G"))
			C1B = CSNG(LoopRs("C1B"))
			C1A = C1G + C1B
			IF C1A = 0 THEN
				C1F = 0
			ELSE
				C1F = C1B / C1A * 100
			END IF
			
			C2G = CSNG(LoopRs("C2G"))
			C2B = CSNG(LoopRs("C2B"))
			C2A = C2G + C2B
			IF C2A = 0 THEN
				C2F = 0
			ELSE
				C2F = C2B / C2A * 100
			END IF
			
			C3G = CSNG(LoopRs("C3G"))
			C3B = CSNG(LoopRs("C3B"))
			C3A = C3G + C3B
			IF C3A = 0 THEN
				C3F = 0
			ELSE
				C3F = C3B / C3A * 100
			END IF
			
			C4G = CSNG(LoopRs("C4G"))
			C4B = CSNG(LoopRs("C4B"))
			C4A = C4G + C4B
			IF C4A = 0 THEN
				C4F = 0
			ELSE
				C4F = C4B / C4A * 100
			END IF
			'response.write C4A
			
			C5G = CSNG(LoopRs("C5G"))
			C5B = CSNG(LoopRs("C5B"))
			C5A = C5G + C5B
			IF C5A = 0 THEN
				C5F = 0
			ELSE
				C5F = C5B / C5A * 100
			END IF
			
			C3GW = CSNG(LoopRs("C3GW"))
			C5GW = CSNG(LoopRs("C5GW"))
			
			if T1F > T1T then
				T1FStr = "<font color='red'>" & Formatnumber(T1F,2,-1) & "%</font>"
			else
				T1FStr = "<font color='green'>" & Formatnumber(T1F,2,-1) & "%</font>"
			end if
			
			if C1F > C1T then
				C1FStr = "<font color='red'>" & Formatnumber(C1F,2,-1) & "%</font>"
			else
				C1FStr = "<font color='green'>" & Formatnumber(C1F,2,-1) & "%</font>"
			end if
			
			if C2F > C2T then
				C2FStr = "<font color='red'>" & Formatnumber(C2F,2,-1) & "%</font>"
			else
				C2FStr = "<font color='green'>" & Formatnumber(C2F,2,-1) & "%</font>"
			end if

			if C3F > C3T then
				C3FStr = "<font color='red'>" & Formatnumber(C3F,2,-1) & "%</font>"
			else
				C3FStr = "<font color='green'>" & Formatnumber(C3F,2,-1) & "%</font>"
			end if
			
			if C4F > C4T then
				C4FStr = "<font color='red'>" & Formatnumber(C4F,2,-1) & "%</font>"
			else
				C4FStr = "<font color='green'>" & Formatnumber(C4F,2,-1) & "%</font>"
			end if
			
			if C5F > C5T then
				C5FStr = "<font color='red'>" & Formatnumber(C5F,2,-1) & "%</font>"
			else
				C5FStr = "<font color='green'>" & Formatnumber(C5F,2,-1) & "%</font>"
			end if

			if ProductName = "RA" or ProductName = "Slim" then
				COG = C5G
				COW = C5GW
			else
				COG = C3G
				COW = C3GW
			end if
			
			Oee = Formatnumber(COG / (OPENTIME * 60),2,-1)
			Teep = Formatnumber(COG / 720,2,-1)
			ProdTimeT = ProdTimeT + OPENTIME
			DownTimeT = DownTimeT + STOPTIME
			TargetCCT = TargetCCT + TargetCC
			COGT = COGT + COG
			COWT = COWT + COW
			T1FT = T1FT + T1F
			C1FT = C1FT + C1F
			C2FT = C2FT + C2F
			C3FT = C3FT + C3F
			C4FT = C4FT + C4F
			C5FT = C5FT + C5F
			OeeT = OeeT + Oee
			TeepT = TeepT + Teep
			
			ListStr = "<tr " & TrStyle & "><td>Line "&ReplaceCcl(LineName)&"</td><td>"& Shift & "</td><td>" & OPENTIME & "h</td><td>" & STOPTIME & "h</td><td>" & TargetCC & "</td><td>" & COW & "</td><td>" & COG & "</td>"
			if ProductName = "RA" or ProductName = "Slim" then ListStr = ListStr & "<td>" & C5FStr & "</td><td>" & C4FStr & "</td>"
			ListStr = ListStr & "<td>" & C3FStr & "</td><td>" & C2FStr & "</td><td>" & C1FStr & "</td><td>" & T1FStr & "</td><td>" & Oee & "%</td><td>" & Teep & "%</td></tr>"
			response.write ListStr

			'提取Reamrk开始
				RemarkStrSum = RemarkStrSum & "<tr><td><table width='1020' border='1' cellpadding='0' cellspacing='0'>" & ListStr & "</table></td></tr><tr style='background:#FFFFFF;'><td align='left'>"

				set rs = Server.CreateObject("adodb.recordset")
				sql = "SELECT * FROM [dbo].[Day_Data_PC] where ProductName = '"&ProductName&"' and LineName = '"&LineName&"' and DName = '"&DName&"' and Adate = '"&ADATE&"'"
				'response.write sql & "<br>"
				rs.open sql,ConnSql,1,1
				StopStr = ""
				PCStr = ""
				if not rs.eof then
					PcSum = rs("PcSum")
					PcQty = rs("PcQty")
					PlanStop = rs("PlanStop")
					if PlanStop = "0" OR PlanStop = "" then PlanStop = "0"
					StopStr = "计划停机: " & PlanStop & " 小时<br><br>"
					PCStr = "在线主载体数量: " & PcSum & "　待修主载体数量: " & PcQty & "<br>"
				end if
				rs.close
				set rs = nothing

				set rs = Server.CreateObject("adodb.recordset")
				sql = "SELECT * FROM [dbo].[Day_Data_Machine] where ProductName = '"&ProductName&"' and LineName = '"&LineName&"' and DName = '"&DName&"' and Adate = '"&ADATE&"' order by Rid"
				'response.write sql & "<br>"
				rs.open sql,ConnSql,1,1
				if not rs.eof then
					RemarkStrSum = RemarkStrSum & ADATE & "　" & ProductName & " " & ReplaceCcl(LineName) & "　D/N: " & rs("DName") & "　Shift: " & ShiftName & " </BR></BR>"
					RemarkStrSum = RemarkStrSum & StopStr & PCStr
					do while not rs.eof
						Dlossn = rs("Dlossn")
						ADATE = FormatTime(ADATE,2)
						Dhtime = rs("Dhtime")
						Detime = rs("Detime")
						Dtime = rs("Dtime")
						Dequipment = rs("Dequipment")
						DPosition = rs("DPosition")
						Dproblem = rs("Dproblem")
						DReason = rs("DReason")
						Dsolution = rs("Dsolution")
							RemarkStr = Dhtime & " - " & Detime & "　停机：" & Dtime & " 分钟<br>"
							RemarkStr = RemarkStr & "位置：" & Dequipment & " " & DPosition & "<br>"
							RemarkStr = RemarkStr & "问题：" & Dproblem & "<br>"
							RemarkStr = RemarkStr & "原因：" & DReason & "<br>"
							RemarkStr = RemarkStr & "解决：" & Dsolution & "<br>"
							RemarkStr = RemarkStr & "<br>"
						RemarkStrSum = RemarkStrSum & RemarkStr
					rs.movenext
					loop
				else
					RemarkStrSum = RemarkStrSum & "没有维护记录<br><br>"
				end if
				rs.close
				set rs = nothing
				'RemarkStrSum = RemarkStrSum & RemarkStr
			'提取Reamrk结束

			LoopRs.movenext
			LoopNum = LoopNum + 1
			loop
			LoopRs.close
			set LoopRs = nothing
			'--------------------------------------------------------'
			ListTotalStr = "<tr style='background:#FAD97D;font-size:15px;font-weight:bold;text-align:center;'>"
			ListTotalStr = ListTotalStr & "<td colspan='2'>Line" & ReplaceCcl(LineName) & " Total</td>"
			ListTotalStr = ListTotalStr & "<td>" & ProdTimeT & "h</td><td>" & DownTimeT & "h</td>"
			ListTotalStr = ListTotalStr & "<td>" & Formatnumber(TargetCCT,0,-1) & "</td>"
			ListTotalStr = ListTotalStr & "<td>" & Formatnumber(COWT,0,-1) & "</td>"
			ListTotalStr = ListTotalStr & "<td>" & Formatnumber(COGT,0,-1) & "</td>"
			if ProductName = "RA" or ProductName = "Slim" then
				ListTotalStr = ListTotalStr & "<td>" & Formatnumber(C5FT/LoopNum,2,-1) & "%</td>"
				ListTotalStr = ListTotalStr & "<td>" & Formatnumber(C4FT/LoopNum,2,-1) & "%</td>"
			end if
			ListTotalStr = ListTotalStr & "<td>" & Formatnumber(C3FT/LoopNum,2,-1) & "%</td>"
			ListTotalStr = ListTotalStr & "<td>" & Formatnumber(C2FT/LoopNum,2,-1) & "%</td>"
			ListTotalStr = ListTotalStr & "<td>" & Formatnumber(C1FT/LoopNum,2,-1) & "%</td>"
			ListTotalStr = ListTotalStr & "<td>" & Formatnumber(T1FT/LoopNum,2,-1) & "%</td>"
			ListTotalStr = ListTotalStr & "<td>" & Formatnumber(OeeT/LoopNum,2,-1) & "%</td>"
			ListTotalStr = ListTotalStr & "<td>" & Formatnumber(TeepT/LoopNum,2,-1) & "%</td></tr>"
			response.write ListTotalStr
	RsL.movenext
	loop
	
	response.write "</table>"
	response.write "<table width='1018' border='1' cellpadding='0' cellspacing='0'>"
	response.write "<tr style='background:#C9C9C9;height:26px;font-size:20px;'><th><strong>Remarks:</strong></th></tr>"
	response.write ""&RemarkStrSum&"</td></tr>"
	response.write "</table></div><br>"
end if

function TODB (TODBvalue)
	if isnull(TODBvalue) then
		TODB=""
	else
		TODB = TODBvalue
	end if
end function
%>

</table>
</body>
</html>
