<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/SendJMail.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<%
Response.buffer=true 
StartDate = formatdate(dateadd("d",-1,now()),"yyyymmdd")
EndDate = formatdate(now(),"yyyymmdd")

reportTime=replace(dateadd("d",-1,now()),"/","-")
fromTime=formatdate(reportTime,"dd.mm.yyyy")&" 07:15"
toTime=formatdate(now(),"dd.mm.yyyy")&" 07:14"
'response.write replace(StartDate & " 07:15:00","-","/")
'response.write "</br>"
'response.write replace(EndDate & " 07:15:00","-","/")

mailTitleFromDate = formatdate(reportTime,"dd.mm.yyyy")
reportDay=longweekdayconvert(weekday(reportTime))&","&formatdate(now(),"dd.mm.yyyy")

htmls="<html><head><meta http-equiv='Content-Type' content='text/html; charset=utf-8'>"
htmls=htmls&"<style type='text/css'>"
htmls=htmls&"body {font:normal 12px/13px 'Arial';}"
htmls=htmls&"table {border-collapse:collapse;border-spacing:0;border:1px solid black;}"
htmls=htmls&"table th{border-color:black;font-size:14px;}"
htmls=htmls&"table td{border-color:black;}"
htmls=htmls&"img{margin:5px;}"
htmls=htmls&"</style>"
htmls=htmls&"</head><body>"
'绘制页面上方页头
htmls=htmls&"<table width='100%' border='1' cellpadding='0' cellspacing='0' bgcolor='#7BB1DB'>"
htmls=htmls&"<tr><td><table width='100%' border='0' cellpadding='0' cellspacing='0' bgcolor='#7BB1DB'>"
htmls=htmls&"<tr><td width='26%' rowspan='4'>&nbsp;</td>"
htmls=htmls&"<td colspan='2' align='center'><span style='height:40px;line-height:40px;font-size:24px;'><b>Daily Production Report - Asia</b></span></td>"
htmls=htmls&"<td width='24%' rowspan='4' align='right'><img src='http://"&application("HostServer")&"/Images/rpt-logo1.jpg' /></td></tr>"
htmls=htmls&"<tr><td width='24%'><b>Report Day: &nbsp;</b>"&reportDay&"</td>"
htmls=htmls&"<td width='26%'><b>Time:&nbsp;</b> "&fromTime&" - "&toTime&"</td></tr>"
htmls=htmls&"<tr><td><b>Author:&nbsp;</b>Zhang Hui, Wang Binglin</td>"
htmls=htmls&"<td><b>Print Date:&nbsp;</b>"&formatdate(now(),"dd.mm.yyyy")&"</td></tr>"
htmls=htmls&"<tr><td colspan='2'>&nbsp;</td></tr></table></td></tr></table></br>"

'绘制页面下方描述
htmlDesc="<table width='100%' border='1' cellpadding='0'>"
htmlDesc=htmlDesc&"<tr><td colspan='2' align='center' bgcolor='#7BB1DB' ></b>Description</b></td></tr>"
htmlDesc=htmlDesc&"<tr><td width='120'>Target[pcs]</td><td>Set by production</td></tr>"
htmlDesc=htmlDesc&"<tr><td>Actual[pcs]</td><td>Pre-Store. last day 24H output time form 7:15 to 7:15 </td></tr>"
htmlDesc=htmlDesc&"<tr><td>Real to W/H</td><td>Warehouse Actual Received Qty,  last day 24H send to WH QTY . time from 7:15 to 7:15</td></tr>"
htmlDesc=htmlDesc&"<tr><td>Process Yield</td><td> Hektor : Main line , (Finished Assy. Good Qty / Start Qty) last day 24H output time form 7:15 to 7:15 </td></tr>"
htmlDesc=htmlDesc&"<tr><td>Acoustic Yield</td><td>Acoustic Test (Acoustic Pass / Acoustic Input) last day 24H output time form 7:15 to 7:15</td></tr>"
htmlDesc=htmlDesc&"<tr><td>Cosmetic Yield</td><td>Cosmetic Yield, (Cosmetic Pass / Cosmetic Input) last day 24H output time form 7:15 to 7:15</td></tr>"
htmlDesc=htmlDesc&"<tr><td>FPY</td><td>Total Yield, (Job Close Quantity / Job Open Quantity) last day 24H completed all JO yield</td></tr>"
htmlDesc=htmlDesc&"<tr><td>Job Qty</td><td>Job Qty, Job closed quantity and the relevant open quantity</td></tr>"
htmlDesc=htmlDesc&"<tr><td>Remarks</td><td>This reprot not included rework JO and QYT </td></tr></table>"

TangoAll = "'Maple','Hektor','Marigold','Pansy','Garland Highflex'"
AutoAll = "'PETRA','RA 11x15 Speaker (Danubius)','Donau Slim'"
OemAll = "'16mm MFD (KEB)','16mm MFD (KEW)','16mm MFD (797)','Julia-HAC','Julia-SC','Nith','Donau Slim (797)','Donau Slim (KEB)','Tesla'"
BeijingAll = TangoAll & "," & AutoAll

TangoHtmlStr = WriteHtml("Details Beijing (Tango):",TangoAll)		'TangoLines
AutoLineHtmlStr = WriteHtml("Details Beijing (Autolines):",AutoAll)		'AutoLines
OemLineHtmlStr =  WriteHtml("Details Beijing (OEM):",OemAll)		'OemLines
TitelTableHtml = TotalLineHtml("1","Total",BeijingAll,OemAll)		'Total

mailContents = htmls & TitelTableHtml & TangoHtmlStr & AutoLineHtmlStr & OemLineHtmlStr & htmlDesc & "</body></html>"

MailSwitch = 3			'邮件发送/测试用显示开关
if MailSwitch > 0 then
	select case MailSwitch
			case 1
				mailto="chao.yao@knowles.com"
			case 2
				mailto="bl.wang@knowles.com;hui.zhang@knowles.com;Xuehui.Chen@knowles.com;mark.fan@knowles.com;fanny.zhang@knowles.com;kunlun.liu@knowles.com;Vivian.Huang@knowles.com;Ivan.li@knowles.com;Young.Li@knowles.com;bo.wang@knowles.com;chao.yao@knowles.com"
			case 3
				mailto="_KEB_BPS_Report@knowles.com"
			case else
				response.Write(mailContents)
	end select
%>
<script type="text/javascript">
	window.opener=null;
	window.open("","_self");
	window.close();
</script>
<%
	SendJMail application("MailSender"),mailTo,"BPS Daily Report (Production Details) - " & mailTitleFromDate,mailContents
else
	response.Write(mailContents)
end if

'------------------------------------------------------功能函数部分-----------------------------------------------------------------'
'绘制表头函数
function TitelTable(TitelName)
	TotalTitel="<table width='100%' border='1' cellpadding='0' cellspacing='0' >"
	TotalTitel=TotalTitel&"<tr align='center' bgcolor='#DCE6F1'><th rowspan='2' width='16%'>Location</th><th rowspan='2' width='7%'>Lines</th><th rowspan='2' width='7%'>Target<br>[pcs]</th>"
	TotalTitel=TotalTitel&"<th rowspan='2' width='7%'>Actual<br>[pcs]</th><th rowspan='2' width='7%'>Real to<br>W/H</th>"
	TotalTitel=TotalTitel&"<th colspan='2'>Process Yield</th><th colspan='2'>Acoustic Yield</th><th colspan='2'>Cosmetic Yield</th><th colspan='2'>FPY</th></tr>"
	TotalTitel=TotalTitel&"<tr align='center' bgcolor='#DCE6F1'><th width='7%'>Target</th><th width='7%'>Real</th><th width='7%'>Target</th><th width='7%'>Real</th><th width='7%'>Target</th>"
	TitelTable=TotalTitel&"<th width='7%'>Real</th><th width='7%'>Target</th><th width='7%'>Real</th></tr>"
end function
'绘制分类表头函数
function ProdTypeTable(ProdType)
	TypeTitel="<table width='100%' border='1' cellpadding='0' cellspacing='0' ><tr align='center' bgcolor='#DCE6F1'><th colspan='21'>"&ProdType&"</th></tr>"
	TypeTitel=TypeTitel&"<tr align='center' bgcolor='#DCE6F1'><th rowspan='2' width='13%'>Product</th><th rowspan='2' width='4%'>Line</th><th rowspan='2' width='4%'>Target<br>[pcs]</th>"
	TypeTitel=TypeTitel&"<th rowspan='2' width='4%'>Actual<br>[pcs]</th><th rowspan='2' width='4%'>Real to<br>W/H</th>"
	TypeTitel=TypeTitel&"<th colspan='4'>Process Yield</th><th colspan='4'>Acoustic Yield</th><th colspan='4'>Cosmetic Yield</th>"
	TypeTitel=TypeTitel&"<th colspan='2'>Job Quantity</th><th colspan='2'>FPY</th></tr>"
	TypeTitel=TypeTitel&"<tr align='center' bgcolor='#DCE6F1'><th  width='4%'>Input</th><th  width='4%'>Output</th><th  width='4%'>Target</th><th  width='4%'>Real</th><th  width='4%'>Input</th><th  width='4%'>Output</th><th  width='4%'>Target</th><th  width='4%'>Real</th><th  width='4%'>Input</th><th  width='4%'>Output</th><th  width='4%'>Target</th><th width='4%'>Real</th>"
	TypeTitel=TypeTitel&"<th  width='4%'>Input</th><th  width='4%'>Output</th><th width='4%'>Target</th><th width='4%'>Real</th></tr>"
	ProdTypeTable = TypeTitel
end function
'绘制函数
function LoopWrite(SqlStr)
	SqlStr = SqlStr
	set RsL = Server.CreateObject("adodb.recordset")
	RsL.open SqlStr,conn,1,1
	do while not RsL.eof
		
		ProductName = RsL("PRODUCT")
		LineId = RsL("LINE_NUM")
		SumToatl = ""

		if LineId = "Total" then
			LineIdStr = LineId
			TrStyle = "style='background:#E6F1FC;font-size:14px;font-weight:bolder;'"
		elseif ProductName = "Beijing" then
			Set RsTemp = Server.CreateObject("adodb.recordset")
			SqlTemp = "select count(*) LineId,'SumTotal' ST from (select distinct product,line_num from vw_bps_dr_all where product in ("&BeijingAll&") and LINE_NUM not in ('Total','8411','8412'))"
			RsTemp.open SqlTemp,conn,1,1
			if not RsTemp.eof then
				LineIdStr = RsTemp("LineId")
				SumToatl = RsTemp("ST")
			end if
			RsTemp.close
			set RsTemp = nothing
			TrStyle = "style='background:#ffffff;font-size:13px;text-align:center;'"
		elseif ProductName = "Beijing Oem" then
			Set RsTemp = Server.CreateObject("adodb.recordset")
			SqlTemp = "select count(*) LineId,'SumTotal' ST from (select distinct product,line_num from vw_bps_dr_all where product in ("&OemAll&") and LINE_NUM not in ('Total','8411','8412'))"
			RsTemp.open SqlTemp,conn,1,1
			if not RsTemp.eof then
				LineIdStr = RsTemp("LineId")
				SumToatl = RsTemp("ST")
			end if
			RsTemp.close
			set RsTemp = nothing
			TrStyle = "style='background:#ffffff;font-size:13px;text-align:center;'"
		else
			LineIdStr = "Line#" & ReplaceCcl(LineId)
			TrStyle = "style='background:#ffffff;font-size:13px;'"
		end if

		TargetPcs = Csng(RsL("TARGET"))
		ActualPcs = Csng(RsL("ACTUAL"))
			if ActualPcs = 0 then
				ActualPcsStr = "-"
			elseif ActualPcs < TargetPcs then
				ActualPcsStr = "<font color='red'>" & Formatnumber(ActualPcs,0,-1) & "</font>"
			else
				ActualPcsStr = "<font color='green'>" & Formatnumber(ActualPcs,0,-1) & "</font>"
			end if
						
		RealToWh = Csng(RsL("REAL_TO_WH"))
			if RealToWh = 0 then
				RealToWhStr = "-"
			else
				RealToWhStr = Formatnumber(RealToWh,0,-1)
			end if
			
		ProcessTarget = Csng(RsL("PQ_TARGET"))
		ProcessReal = Csng(RsL("PQ_YIELD"))
			if ProcessReal = 0 then
				ProcessRealStr = "-"
			elseif ProcessReal < ProcessTarget then
				ProcessRealStr = "<font color='red'>" & Formatnumber(ProcessReal,1) & "%</font>"
			else
				ProcessRealStr = "<font color='green'>" & Formatnumber(ProcessReal,1) & "%</font>"
			end if
			
		AcousticTarget = Csng(RsL("AMS_TARGET"))
		AcousticReal = Csng(RsL("AMS_YIELD"))
			if AcousticReal = 0 then
				AcousticRealStr = "-"
			elseif AcousticReal < AcousticTarget then
				AcousticRealStr = "<font color='red'>" & Formatnumber(AcousticReal,1) & "%</font>"
			else
				AcousticRealStr = "<font color='green'>" & Formatnumber(AcousticReal,1) & "%</font>"
			end if
				
		CosmeticTarget = Csng(RsL("COSMETIC_TARGET"))
		CosmeticReal = Csng(RsL("COSMETIC_YIELD"))
			if CosmeticReal = 0 then
				CosmeticRealStr = "-"
			elseif CosmeticReal < CosmeticTarget then
				CosmeticRealStr = "<font color='red'>" & Formatnumber(CosmeticReal,1) & "%</font>"
			else
				CosmeticRealStr = "<font color='green'>" & Formatnumber(CosmeticReal,1) & "%</font>"
			end if
				
		FpyTarget = Csng(RsL("TTY_TARGET"))
		FpyReal = Csng(RsL("TTY_YIELD"))
			if FpyReal = 0 then
				FpyRealStr = "-"
			elseif FpyReal < FpyTarget then
				FpyRealStr = "<font color='red'>" & Formatnumber(FpyReal,1) & "%</font>"
			else
				FpyRealStr = "<font color='green'>" & Formatnumber(FpyReal,1) & "%</font>"
			end if
			
			if SumToatl = "" then 
				PQ_START = csng(RsL("PQ_START"))
					if PQ_START = 0 then
						PQ_STARTstr = "-"
					else
						PQ_STARTstr = PQ_START
					end if
					
				PQ_OUT = csng(RsL("PQ_OUT"))
					if PQ_OUT = 0 then
						PQ_OUTstr = "-"
					else
						PQ_OUTstr = PQ_OUT
					end if
					
				AMS_START = csng(RsL("AMS_START"))
					if AMS_START = 0 then
						AMS_STARTstr = "-"
					else
						AMS_STARTstr = AMS_START
					end if
					
				AMS_GOOD = csng(RsL("AMS_GOOD"))
					if AMS_GOOD = 0 then
						AMS_GOODstr = "-"
					else
						AMS_GOODstr = AMS_GOOD
					end if
				COSMETIC_START = csng(RsL("COSMETIC_START"))
					if COSMETIC_START = 0 then
						COSMETIC_STARTstr = "-"
					else
						COSMETIC_STARTstr = COSMETIC_START
					end if
					
				COSMETIC_GOOD = csng(RsL("COSMETIC_GOOD"))
					if COSMETIC_GOOD = 0 then
						COSMETIC_GOODstr = "-"
					else
						COSMETIC_GOODstr = COSMETIC_GOOD
					end if
					
				TTY_OPEN = csng(RsL("TTY_OPEN"))
					if TTY_OPEN = 0 then
						TTY_OPENstr = "-"
					else
						TTY_OPENstr = TTY_OPEN
					end if
					
				TTY_CLOSE = csng(RsL("TTY_CLOSE"))
					if TTY_CLOSE = 0 then
						TTY_CLOSEstr = "-"
					else
						TTY_CLOSEstr = TTY_CLOSE
					end if
					
			end if

		if LineId <> "0" then
			LoopWrite = LoopWrite & "<tr "&TrStyle&"><td>"&ProductName&"</td>"
			LoopWrite = LoopWrite & "<td align='center'>"&LineIdStr&"</td>"
			LoopWrite = LoopWrite & "<td align='center'>"&Formatnumber(TargetPcs,0,-1)&"</td>"
			LoopWrite = LoopWrite & "<td align='center'>"&ActualPcsStr&"</td>"
			LoopWrite = LoopWrite & "<td align='center'>"&RealToWhStr&"</td>"
			
			if SumToatl = "" then 
				LoopWrite = LoopWrite & "<td align='center'>"&PQ_STARTstr&"</td>"
				LoopWrite = LoopWrite & "<td align='center'>"&PQ_OUTstr&"</td>"
			end if
			
			LoopWrite = LoopWrite & "<td align='center'>"&Formatnumber(ProcessTarget,1)&"%</td>"
			LoopWrite = LoopWrite & "<td align='center'>"&ProcessRealStr&"</td>"
			
			if SumToatl = "" then 
				LoopWrite = LoopWrite & "<td align='center'>"&AMS_STARTstr&"</td>"
				LoopWrite = LoopWrite & "<td align='center'>"&AMS_GOODstr&"</td>"
			end if
			
			LoopWrite = LoopWrite & "<td align='center'>"&Formatnumber(AcousticTarget,1)&"%</td>"
			LoopWrite = LoopWrite & "<td align='center'>"&AcousticRealStr&"</td>"
			
			if SumToatl = "" then 
				LoopWrite = LoopWrite & "<td align='center'>"&COSMETIC_STARTstr&"</td>"
				LoopWrite = LoopWrite & "<td align='center'>"&COSMETIC_GOODstr&"</td>"
			end if
			
			LoopWrite = LoopWrite & "<td align='center'>"&Formatnumber(CosmeticTarget,1)&"%</td>"
			LoopWrite = LoopWrite & "<td align='center'>"&CosmeticRealStr&"</td>"
			
			if SumToatl = "" then 
				LoopWrite = LoopWrite & "<td align='center'>"&TTY_OPENstr&"</td>"
				LoopWrite = LoopWrite & "<td align='center'>"&TTY_CLOSEstr&"</td>"
			end if
			
			LoopWrite = LoopWrite & "<td align='center'>"&Formatnumber(FpyTarget,1)&"%</td>"
			LoopWrite = LoopWrite & "<td align='center'>"&FpyRealStr&"</td></tr>"
		end if			 
	RsL.movenext
	loop
	RsL.close
	set RsL = nothing
end function

'循环绘制
function WriteHtml(LineType,LineName)
		set RsRs = Server.CreateObject("adodb.recordset")
		SqlL = "select product,COUNT(product) Lines from vw_bps_dr_all where product in ("&LineName&") and LINE_NUM not in ('Total','8411','8412') group by product order by product"
		' response.write SqlL & "</br></br>"
		' Response.flush()
		RsRs.open SqlL,conn,1,1
		if not RsRs.eof then
			WriteHtml = ProdTypeTable(LineType)
			do while not RsRs.eof
				LineName1 = RsRs("product")
				Lines1 = RsRs("Lines")
				set Rs = Server.CreateObject("adodb.recordset")
				SqlStr = "select product,COUNT(product) Lines,line_num,line_type from vw_bps_dr_all where product in ('"&LineName1&"') and LINE_NUM not in ('Total','8411','8412') group by product,line_num,line_type"
				Rs.open SqlStr,conn,1,1
				if not Rs.eof then
					ProName = rs("product")
					Lines = rs("Lines")
					LineTy = rs("line_type")
						SqlStrL = "select product,line_num,nvl(sum(TARGET),0) TARGET,nvl(sum(ACTUAL),0) ACTUAL,nvl(sum(REAL_TO_WH),0) REAL_TO_WH,nvl(sum(PQ_OUT),0) PQ_OUT,sum(nvl(PQ_START,0)+nvl(PQ_OFFSET,0)) PQ_START,nvl(avg(PQ_TARGET),0) PQ_TARGET,nvl(avg(PQ_YIELD),0) PQ_YIELD,nvl(sum(AMS_GOOD),0) AMS_GOOD,nvl(sum(AMS_START),0) AMS_START,nvl(avg(AMS_TARGET),0) AMS_TARGET,nvl(avg(AMS_YIELD),0) AMS_YIELD,nvl(sum(COSMETIC_GOOD),0) COSMETIC_GOOD,nvl(sum(COSMETIC_START),0) COSMETIC_START,nvl(avg(COSMETIC_TARGET),0) COSMETIC_TARGET,nvl(avg(COSMETIC_YIELD),0) COSMETIC_YIELD,nvl(sum(TTY_OPEN),0) TTY_OPEN,nvl(sum(TTY_CLOSE),0) TTY_CLOSE,nvl(avg(TTY_TARGET),0) TTY_TARGET,nvl(avg(TTY_YIELD),0) TTY_YIELD from vw_bps_dr_all where product in ('"&ProName&"') and LINE_NUM not in ('Total','8411','8412') group by product,line_num order by line_num"
						'response.write ProName & " - " & LineTy & " - " & Lines1 & " - " &  Lines & "</br></br>"
						'Response.flush()
						WriteHtml = WriteHtml & LoopWrite(SqlStrL)
						if (Cint(Lines1)/cint(Lines)) > 1 then
							if LineTy = "AUTO" then
								'response.write ProName & " - " & LineTy & " - " & Lines1 & " - " &  Lines & "</br></br>"
								SqlStrT = "select product,COUNT(product) Lines,'Total' line_num,avg(pq_target) pq_target,decode(sum(nvl(pq_start,0)+nvl(PQ_OFFSET,0)),0,0,(sum(nvl(pq_out,0))/sum(nvl(pq_start,1)+nvl(PQ_OFFSET,0))))*100 pq_yield,avg(ams_target) ams_target,decode(sum(nvl(ams_start,1)),0,0,(sum(nvl(AMS_GOOD,0))/sum(nvl(ams_start,1))))*100 ams_yield,avg(cosmetic_target) cosmetic_target,decode(sum(nvl(cosmetic_start,1)),0,0,(sum(nvl(cosmetic_good,0))/sum(nvl(cosmetic_start,1))))*100 cosmetic_yield,AVG(tty_target) tty_target,decode(sum(nvl(TTY_OPEN,1)),0,0,nvl(sum(tty_close),0)/nvl(sum(TTY_OPEN),0))*100 TTy_yield,sum(nvl(target,0)) Target,sum(nvl(actual,0)) Actual,sum(nvl(real_to_wh,0)) real_to_wh,sum(nvl(pq_out,0)) pq_out,sum(nvl(pq_start,0)+nvl(PQ_OFFSET,0)) pq_start,sum(nvl(AMS_GOOD,0)) AMS_GOOD,sum(nvl(ams_start,0)) ams_start,sum(nvl(cosmetic_good,0)) cosmetic_good,SUM(nvl(cosmetic_start,0)) cosmetic_start,sum(nvl(tty_close,0)) tty_close,nvl(SUM(TTY_OPEN),0) TTY_OPEN from vw_bps_dr_all where product in ('"&ProName&"') and LINE_NUM not in ('Total','8411','8412') group by product"
							elseif LineTy = "OEM" then
								SqlStrT = "select product,COUNT(product) Lines,'Total' line_num,avg(pq_target) pq_target,decode(sum(nvl(pq_start,0)+nvl(PQ_OFFSET,0)),0,0,(sum(nvl(pq_out,0))/sum(nvl(pq_start,1)+nvl(PQ_OFFSET,0))))*100 pq_yield,avg(ams_target) ams_target,decode(sum(nvl(ams_start,1)),0,0,(sum(nvl(AMS_GOOD,0))/sum(nvl(ams_start,1))))*100 ams_yield,avg(cosmetic_target) cosmetic_target,decode(sum(nvl(cosmetic_start,1)),0,0,(sum(nvl(cosmetic_good,0))/sum(nvl(cosmetic_start,1))))*100 cosmetic_yield,AVG(tty_target) tty_target,decode(sum(nvl(TTY_OPEN,1)),0,0,nvl(sum(tty_close),0)/nvl(sum(TTY_OPEN),0))*100 TTy_yield,sum(nvl(target,0)) Target,sum(nvl(actual,0)) Actual,sum(nvl(real_to_wh,0)) real_to_wh,sum(nvl(pq_out,0)) pq_out,sum(nvl(pq_start,0)+nvl(PQ_OFFSET,0)) pq_start,sum(nvl(AMS_GOOD,0)) AMS_GOOD,sum(nvl(ams_start,0)) ams_start,sum(nvl(cosmetic_good,0)) cosmetic_good,SUM(nvl(cosmetic_start,0)) cosmetic_start,sum(nvl(tty_close,0)) tty_close,nvl(SUM(TTY_OPEN),0) TTY_OPEN from vw_bps_dr_all where product in ('"&ProName&"') and LINE_NUM not in ('Total','8411','8412') group by product"
							else
								SqlStrT = "select product,COUNT(product) Lines,'Total' line_num,avg(pq_target) pq_target,decode(sum(nvl(pq_start,0)+nvl(PQ_OFFSET,0)),0,0,(sum(nvl(pq_out,0))/sum(nvl(pq_start,1)+nvl(PQ_OFFSET,0))))*100 pq_yield,avg(ams_target) ams_target,decode(sum(nvl(ams_start,1)),0,0,((sum(nvl(AMS_GOOD,0)) + (sum(nvl(ams_start,1)) - sum(nvl(AMS_GOOD,0)))* (select AVG(acousitc_offset) from RPT_dAILY_TARGET where product in ('"&ProName&"') and line in (select LINE_NUM from vw_bps_dr_all where product in ('"&ProName&"'))) )/sum(nvl(ams_start,1))))*100 ams_yield,avg(cosmetic_target) cosmetic_target,decode(sum(nvl(cosmetic_start,1)),0,0,(sum(nvl(cosmetic_good,0))/sum(nvl(cosmetic_start,1))))*100 cosmetic_yield,AVG(tty_target) tty_target,decode(sum(nvl(TTY_OPEN,1)),0,0,(sum(nvl(tty_close,0))/sum(nvl(TTY_OPEN,1))))*100 TTy_yield,sum(nvl(target,0)) Target,sum(nvl(actual,0)) Actual,sum(nvl(real_to_wh,0)) real_to_wh,sum(nvl(pq_out,0)) pq_out,sum(nvl(pq_start,0)+nvl(PQ_OFFSET,0)) pq_start,sum(nvl(AMS_GOOD,0)) AMS_GOOD,sum(nvl(ams_start,0)) ams_start,sum(nvl(cosmetic_good,0)) cosmetic_good,SUM(nvl(cosmetic_start,0)) cosmetic_start,sum(nvl(tty_close,0)) tty_close,SUM(nvl(TTY_OPEN,0)) TTY_OPEN from vw_bps_dr_all where product in ('"&ProName&"') and  LINE_NUM not in ('Total','8411','8412') group by product"
							end IF
							'response.write "SqlStrT: - " & SqlStrT & "</br></br>"
							WriteHtml = WriteHtml & LoopWrite(SqlStrT)
						end if
					rs.close
					set rs = nothing
					
				end if
			RsRs.movenext
			loop
			WriteHtml = WriteHtml & "</table></br>"
		else
			WriteHtml = ""
		end if
end function

'Total绘制
Function TotalLineHtml(Switch,LineType,BjLineName,OemLineName)
	if Switch = 1 then
		TotalHtml = TitelTable("Total")
		for i = 1 to 2
			if i = 1 then
				SqlStr = "select 'Beijing' product,count(product) LINE_NUM,nvl(sum(target),0) target,nvl(sum(actual),0) actual,nvl(sum(real_to_wh),0) real_to_wh,nvl(sum(target*pq_target),0)/nvl(sum(target),1) pq_target,nvl(sum(target*AMS_target),0)/nvl(sum(target),1) AMS_target,nvl(sum(target*cosmetic_target)/sum(target),0) cosmetic_target,nvl(sum(target*TTY_target)/sum(target),0) TTY_target,decode(nvl(sum(pq_start),0)+nvl(sum(PQ_OFFSET),0),0,0,nvl(sum(pq_out),0)/(nvl(sum(pq_start),0)+nvl(sum(PQ_OFFSET),0))*100) pq_Yield,decode(nvl(sum(ams_start),0),0,0,nvl(sum(ams_good),0)/nvl(sum(ams_start),0)*100) AMS_Yield,decode(nvl(sum(cosmetic_start),0),0,0,nvl(sum(cosmetic_good),0)/nvl(sum(cosmetic_start),0)*100) cosmetic_Yield,decode(nvl(sum(tty_open),0),0,0,nvl(sum(tty_close),0)/sum(nvl(tty_open,0))*100) TTY_Yield from vw_bps_dr_all where product in ("&BjLineName&") and LINE_NUM not in ('Total','8411','8412')"
				'response.write SqlStr
			elseif i = 2 then
				SqlStr = "select 'Beijing Oem' product,count(product) LINE_NUM,nvl(sum(target),0) target,nvl(sum(actual),0) actual,nvl(sum(real_to_wh),0) real_to_wh,nvl(sum(target*pq_target),0)/nvl(sum(target),1) pq_target,nvl(sum(target*AMS_target),0)/nvl(sum(target),1) AMS_target,nvl(sum(target*cosmetic_target)/sum(target),0) cosmetic_target,nvl(sum(target*TTY_target)/sum(target),0) TTY_target,nvl(sum(target*pq_Yield)/sum(target),0) pq_Yield,nvl(sum(target*AMS_Yield)/sum(target),0) AMS_Yield,nvl(sum(target*cosmetic_Yield)/sum(target),0) cosmetic_Yield,nvl(sum(target*TTY_Yield)/sum(target),0) TTY_Yield from vw_bps_dr_all where product in ("&OemAll&")"
			end if
			TotalHtml = TotalHtml & LoopWrite(SqlStr)
		next
		TotalLineHtml = TotalHtml & "</table></br>"
	else
		TotalLineHtml = ""
	end if
end function
'线别改名
function ReplaceCcl(ProductN)
	ReplaceCcl = Replace(ProductN,"841","")
	ReplaceCcl = Replace(ReplaceCcl,"3714","0") 'Maple 4线单独改'
	ReplaceCcl = Replace(ReplaceCcl,"371","")
	ReplaceCcl = Replace(ReplaceCcl,"521","")
	ReplaceCcl = Replace(ReplaceCcl,"801","")
	ReplaceCcl = Replace(ReplaceCcl,"911","")
end function
%>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->