<%session("Page_Role") = ",HV_Reports_Reader"%>
<!--#include virtual="/Functions/Login_Check.asp"-->
<!--#include virtual="/Functions/Role_Check.asp"-->
<!--#include virtual="/include/Functions.asp"-->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" type="text/css" href="../../styles/basic.css" />
<script language="javascript" type="text/javascript" src="../include/ThreeMenuMa.js"></script>
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
body,td,th {
	font-family: Arial;
}
</style>
</head>
<body bgcolor="#FFFFFF">
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
</br>
<center>
<form method="post" action="DailyReportHvPer.asp">
产品：<select name="ProductName" id="TmProduct"></select>
		<select name="LineName" id="TmLine" style="display:none" ></select>
		<select name="CellName" id="TmCell" style="display:none" ></select>
		<script type="text/javascript">
			MenuInit('TmProduct','TmLine','TmCell','<%=ProductName%>','1','1');
		</script>
日期：<input type="text" name="Adate" class="Wdate" id="d122" onFocus="WdatePicker({isShowWeek:true,onpicked:function() {$dp.$('d122_1').value=$dp.cal.getP('W','W');$dp.$('d122_2').value=$dp.cal.getP('W','WW');}})" value="<%=Adate%>" size="10" />&nbsp;
<input type="submit" name="Getdata" value="提交" />
</form>
<center>
<%
if Getdata = "提交" or Getdata = "ok" then
	Startdate = Cdate(Adate)
	EndDate = Startdate
	
	if ProductName = "RA" then
		TitleNmae = "RA 11x15 Speaker (Danubius)"
	else
		TitleNmae = ProductName
	end if
	
	TitleNmae = TitleNmae & " Periphery"
	%>
<table width="960" border="0" cellspacing="0" cellpadding="0" bgcolor="#7BB1DB">
	  <tr>
		<td width="18%">&nbsp;</td>
		<td rowspan="2"><p style="font-size:20px;color:#FFF;text-align:center;"><strong><%=TitleNmae%></strong></p></td>
		<td colspan="2">&nbsp;</td>
	  </tr>
	  <tr>
		<td style="color:#fff;text-align:center;"><a href="DailyReportHvPer.asp?Getdata=ok&ProductName=<%=ProductName%>&Adate=<%=dateadd("d",-1,Startdate)%>"><<</a>　<%=Startdate%>　<a href="DailyReportHvPer.asp?Getdata=ok&ProductName=<%=ProductName%>&Adate=<%=dateadd("d",1,Startdate)%>">>></a></td>
		<td width="13%" rowspan="2"><img src="Images/Logo.jpg" width="124" height="52" align="middle"></td>
		<td width="5%" rowspan="2">&nbsp;</td>
	  </tr>
	  <tr>
		<td style="color:#fff;text-align:center;">&nbsp;</td>
		<td><p style="font-size:20px;color:#FFF;text-align:center;">Daily Report</p></td>
	  </tr>
	  <tr>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td colspan="2">&nbsp;</td>
	  </tr>
</table>
<table width='960' border='1' cellpadding='0' cellspacing='0'>
	<%
	TableTitle = "<tr bgcolor='#B9E9FF'><th>Machine</th><th>DAY</br>Night</th><th>Prod</br>Time</th><th>Down</br>Time</th><th>Output</br>Target</th><th>Output</th><th>FOR</th><th>Deviation</th><th>Output</br>Weekly</th><th>Average</th><th>OEE</th><th>TEEP</th></tr>"
	dim RemarkStr,RemarkStrSum
	RemarkStr = ""
	RemarkStrSum = ""
	
	set RsL = Server.CreateObject("adodb.recordset")
	SqlL = "SELECT [MaName] FROM [KEB_ReportData].[dbo].[Day_Data_Periphery] where '"&Startdate&"' < = adate and adate < = '"&EndDate&"' and ProductName like '"&ProductName&"' group by MaName order by MaName"
	'response.write SqlL & "</br></br>"
	'response.end
	RsL.open SqlL,ConnSql,1,1
	do while not RsL.eof
	
		response.write TableTitle
		
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
		MaName = RsL("MaName")
		'取得Target
		set TarRs = server.createobject("adodb.recordset")
		TarSql = "SELECT [MaSum],[PcsTarget],[FORTarget],ISNULL(OEE,100) OEE,ISNULL(TEEP,100) TEEP from [dbo].[Cfg_Periphery]"
		TarSql = TarSql & " WHERE ProductName = '"&ProductName&"' AND MaName = '"&MaName&"' and Tdate < = '"&Startdate&"' ORDER BY Tdate desc"
		'response.write TarSql & "</br>"
		'response.end
		TarRs.open TarSql,ConnSql,1,1
		MaSum = 1
		PcsTarget = 0
		FORTarget = 0
		OEE = 0
		TEEP = 0
		if not TarRs.eof then
			MaSum = CSNG(TarRs("MaSum"))
			PcsTarget = csng(TarRs("PcsTarget"))
			FORTarget = TarRs("OEE")
			'response.write FORTarget
			OEE = TarRs("OEE")
			TEEP = TarRs("TEEP")
		end if
		TarRs.close
		Set TarRs = nothing
		
		'循环输出开始
		'response.write MaName & "</br>"
		for MaLoop = 1 to MaSum
			'response.write MaLoop& "</br>"
			for DNloop = 1 to 2
				if DNloop = 1 then
					DName = "D"
				else
					DName = "N"
				end if
				
				set ScRs = server.createobject("adodb.recordset")
				SearchSql = "SELECT [DName],[MaName],[MaNr],[OpenTime],[StopTime],[MaGood]"
				SearchSql = SearchSql & ",(select sum([MaGood]) from [dbo].[Day_Data_Periphery] where '"&WeekFirstDay&"' < = adate and adate < = '"&EndDate&"' and DName = DDP.DName and ProductName = DDP.ProductName and MaName = DDP.MaName and MaNr = DDP.MaNr) [MaGoodW]"
				SearchSql = SearchSql & ",(select AVG([MaGood]) from [dbo].[Day_Data_Periphery] where '"&WeekFirstDay&"' < = adate and adate < = '"&EndDate&"' and DName = DDP.DName and ProductName = DDP.ProductName and MaName = DDP.MaName and MaNr = DDP.MaNr) [MaGoodA]"
				SearchSql = SearchSql & ",[MaBad],([MaGood]+[MaBad])[MaAll] FROM [dbo].[Day_Data_Periphery] DDP"
				SearchSql = SearchSql & " where '"&Startdate&"' < = adate and adate < = '"&EndDate&"' and DName = '"&DName&"' and ProductName like '"&ProductName&"' and MaName = '"&MaName&"' and MaNr = '"&MaLoop&"' order by MaNr,DName"
				'response.write SearchSql & "</br>"
				ScRs.open SearchSql,ConnSql,1,1
				if not ScRs.eof then
					OpenTime = ScRs("OpenTime") / 60
					StopTime = ScRs("StopTime")
					MaGood = ScRs("MaGood")
					MaGoodW = ScRs("MaGoodW")
					MaGoodA = ScRs("MaGoodA")
					MaBad = ScRs("MaBad")
					MaAll = ScRs("MaAll")
					if MaAll = 0 then
						MaFOR = 0
					ELSE
						MaFOR = Formatnumber(MaBad / MaAll * 100,2,-1)
					END IF
					MaFORStr = "<font color='red'>" & MaFOR & "%</font>"
					Deviation = Formatnumber(FORTarget - MaFOR,2,-1)
					' OEE=（实际运转时间/12）*（实际产出总数/理论产能（目标））*（实际产出好品/实际产出总数）* 100
					' 例（12/12）*（16600/18000）*（16000/16600）=88.9%
					if PcsTarget = 0 then
						if MaGood = 0 then
							Oee = 0
						else
							Oee = MaGood / MaGood
						end if
					else
						Oee = MaGood / PcsTarget
					end if
					
					if MaAll = 0 then
						Oee = Oee * MaGood * 100
					ELSE
						Oee = Oee * (MaGood / MaAll) * 100
					END IF
					TEEP = Formatnumber(Oee,2,-1)

					Oee = Oee * (OPENTIME / 12)
					Oee = Formatnumber(Oee,2,-1)
					
				end if
				
				if DName = "D" then
					if not ScRs.eof then
						response.write "<tr bgcolor='#FFFFFF'><td>" & MaName & " - " & MaLoop & "</td><td>DAY</td><td>" & OpenTime & "h</td><td>" & StopTime & "h</td><td>" & PcsTarget & "</td><td>" & MaGood & "</td><td>" & MaFORStr & "</td><td>" & Deviation & "</td><td>" & MaGoodW & "</td><td>" & MaGoodA & "</td><td>" & OEE & "%</td><td>" & TEEP & "%</td></tr>"
					else
						response.write "<tr bgcolor='#FFFFFF'><td>" & MaName & " - " & MaLoop & "</td><td>DAY</td><td> - </td><td> - </td><td>" & PcsTarget & "</td><td> - </td><td> - </td><td> - </td><td> - </td><td> - </td><td> - </td><td> - </td></tr>"
					end if
				else
					if not ScRs.eof then
						response.write "<tr bgcolor='#EAEAEA'><td>" & MaName & " - " & MaLoop & "</td><td>Night</td><td>" & OpenTime & "h</td><td>" & StopTime & "h</td><td>" & PcsTarget & "</td><td>" & MaGood & "</td><td>" & MaFORStr & "</td><td>" & Deviation & "</td><td>" & MaGoodW & "</td><td>" & MaGoodA & "</td><td>" & OEE & "%</td><td>" & TEEP & "%</td></tr>"
					else
						response.write "<tr bgcolor='#EAEAEA'><td>" & MaName & " - " & MaLoop & "</td><td>Night</td><td> - </td><td> - </td><td>" & PcsTarget & "</td><td> - </td><td> - </td><td> - </td><td> - </td><td> - </td><td> - </td><td> - </td></tr>"
					end if
				end if
				
				ScRs.close
				set ScRs = nothing
			next
	next
	response.write "<tr><td colspan='10'>　</td></tr>"
	'循环输出结束
	
	set ReRs = server.createobject("adodb.recordset")
	ReSql = "SELECT [DName],[Dsolution] FROM [dbo].[Day_Data_MachinePer]"
	ReSql = ReSql & " where '"&Startdate&"' < = adate and adate < = '"&EndDate&"' and ProductName like '"&ProductName&"' and MaName = '"&MaName&"' ORDER BY DName"
	'response.write ReSql & "</br>"
	ReRs.open ReSql,ConnSql,1,1
	if not ReRs.eof then
		do while not ReRs.eof
			RemarkStrSum = RemarkStrSum & EndDate & " " & ProductName & " " & MaName & " " & ReRs("DName") & "</br></br>"
			RemarkStrSum = RemarkStrSum & ReRs("Dsolution") & "</br>"
		ReRs.movenext
		loop
	end if

	RsL.movenext
	loop
	'response.end
	
	response.write "</table>"
	response.write "<table width='960' border='1' cellpadding='0' cellspacing='0'>"
	response.write "<tr style='background:#C9C9C9;font-size:20px;'><th><strong>Remarks:</strong></th></tr>"
	response.write "<tr><td align='left'>"&RemarkStrSum&"</td></tr>"
	response.write "</table></br>"
end if

function TODB (TODBvalue)
	if isnull(TODBvalue) then
		TODB=""
	else
		TODB = TODBvalue
	end if
end function
%>

</body>
</html>
