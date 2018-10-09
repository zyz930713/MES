<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/SendJMail.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<%
StartDate = formatdate(dateadd("d",-7,now()),"yyyymmdd")
EndDate = formatdate(now(),"yyyymmdd")

reportTime=replace(dateadd("d",-7,now()),"/","-")
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
htmls=htmls&"<td colspan='2' align='center'><span style='height:40px;line-height:40px;font-size:24px;'><b>Weekly Production Report - KEB IT Internal Demo</b></span></td>"
htmls=htmls&"<td width='24%' rowspan='4' align='right'><img src='http://"&application("HostServer")&"/Images/rpt-logo1.jpg' /></td></tr>"
htmls=htmls&"<tr><td width='24%'><b>Report Day: &nbsp;</b>"&reportDay&"</td>"
htmls=htmls&"<td width='26%'><b>Time:&nbsp;</b> "&fromTime&" - "&toTime&"</td></tr>"
htmls=htmls&"<tr><td><b>Author:&nbsp;</b>Jeff Xu, Zhang Hui, Wang Binglin</td>"
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
htmlDesc=htmlDesc&"<tr><td>Remarks</td><td>This reprot not included rework  JO and QYT </td></tr></table>"

TangoAll = "'Maple','Hektor','Marigold','Pansy','Garland Highflex'"
AutoAll = "'PETRA','RA 11x15 Speaker (Danubius)'"
OemAll = "'16mm MFD (KEB)','16mm MFD (KEW)','16mm MFD (797)','Julia-HAC','Julia-SC','Nith'"
BeijingAll = TangoAll & "," & AutoAll

TangoHtmlStr = WriteHtml("Details Beijing (Tango):",TangoAll)		'TangoLines
AutoLineHtmlStr = WriteHtml("Details Beijing (Autolines):",AutoAll)		'AutoLines
OemLineHtmlStr =  WriteHtml("Details Beijing (OEM):",OemAll)		'OemLines
TitelTableHtml = TotalLineHtml("1","Total",BeijingAll,OemAll)		'Total

mailContents = htmls & TitelTableHtml & TangoHtmlStr & AutoLineHtmlStr & OemLineHtmlStr & htmlDesc & "</body></html>"

MailSwitch = 0			'邮件发送/测试用显示开关
if MailSwitch > 0 then
	select case MailSwitch
			case 1
				mailto="chao.yao@knowles.com"
			case 2
				mailto="bl.wang@knowles.com;Jeff.Xu@knowles.com;hui.zhang@knowles.com;Xuehui.Chen@knowles.com;mark.fan@knowles.com;fanny.zhang@knowles.com;kunlun.liu@knowles.com;Vivian.Huang@knowles.com;Ivan.li@knowles.com;Young.Li@knowles.com;bo.wang@knowles.com;chao.yao@knowles.com"
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
	SendJMail application("MailSender"),mailTo,"Daily Report KEB) - " & mailTitleFromDate,mailContents
else
	response.Write(mailContents)
end if

'------------------------------------------------------功能函数部分-----------------------------------------------------------------'
'绘制表头函数
function TitelTable(TitelName)
	TotalTitel="<table width='100%' border='1' cellpadding='0' cellspacing='0' >"
	TotalTitel=TotalTitel&"<tr align='center' bgcolor='#DCE6F1'><td rowspan='2' width='16%'><b>Location</b></td><td rowspan='2' width='6%'><b>Lines</b></td><td rowspan='2' width='6%'><b>Target<br>[pcs]</b></td>"
	TotalTitel=TotalTitel&"<td rowspan='2' width='6%'><b>Actual<br>[pcs]</b></td><td rowspan='2' width='6%'><b>Real to<br>W/H</b></td>"
	TotalTitel=TotalTitel&"<td colspan='2'><b>Process Yield</b></td><td colspan='2'><b>Acoustic Yield</b></td><td colspan='2'><b>Cosmetic Yield</b></td>"
	TotalTitel=TotalTitel&"<td colspan='2'><b>FPY</b></td><td colspan='2'><b>Blocked[kpcs]</b></td></tr>"
	TotalTitel=TotalTitel&"<tr align='center' bgcolor='#DCE6F1'><td width='6%'><b>Target</b></td><td width='6%'><b>Real</b></td><td width='6%'><b>Target</b></td><td width='6%'><b>Real</b></td><td width='6%'><b>Target</b></td>"
	TitelTable=TotalTitel&"<td width='6%'><b>Real</b></td><td width='6%'><b>Target</b></td><td width='6%'><b>Real</b></td><td width='6%'><b>+/-24h</b></td><td width='6%'><b>total</b></td></tr>"
end function
'绘制分类表头函数
function ProdTypeTable(ProdType)
	TypeTitel="<table width='100%' border='1' cellpadding='0' cellspacing='0' ><tr align='center' bgcolor='#DCE6F1'><td colspan='17'><b>"&ProdType&"</b></td></tr>"
	TypeTitel=TypeTitel&"<tr align='center' bgcolor='#DCE6F1'><td rowspan='2' width='16%'><b>Product</b></td><td rowspan='2' width='6%'><b>Line</b></td><td rowspan='2' width='6%'><b>Target<br>[pcs]</b></td>"
	TypeTitel=TypeTitel&"<td rowspan='2' width='6%'><b>Actual<br>[pcs]</b></td><td rowspan='2' width='6%'><b>Real to<br>W/H</b></td>"
	TypeTitel=TypeTitel&"<td colspan='2'><b>Process Yield</b></td><td colspan='2'><b>Acoustic Yield</b></td><td colspan='2'><b>Cosmetic Yield</b></td>"
	TypeTitel=TypeTitel&"<td colspan='2'><b>FPY</b></td><td colspan='2'><b>Blocked[kpcs]</b></td></tr>"
	TypeTitel=TypeTitel&"<tr align='center' bgcolor='#DCE6F1'><td  width='6%'><b>Target</b></td><td  width='6%'><b>Real</b></td><td  width='6%'><b>Target</b></td><td  width='6%'><b>Real</b></td><td  width='6%'><b>Target</b></td>"
	TypeTitel=TypeTitel&"<td width='6%'><b>Real</b></td><td width='6%'><b>Target</b></td><td width='6%'><b>Real</b></td><td width='6%'><b>+/-24h</b></td><td width='6%'><b>total</b></td></tr>"

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

		
		if LineId = "Total" then
			LineIdStr = LineId
			TrStyle = "style='background:#E6F1FC;font-size:14px;font-weight:bolder;'"
		elseif ProductName = "Beijing" then
			Set RsTemp = Server.CreateObject("adodb.recordset")
			SqlTemp = "select count(*) LineId from (select distinct product,line_num from bps_dr_all where product in ("&BeijingAll&") and LINE_NUM not in ('Total','8411','8412') and '"&StartDate&"' < = create_time AND create_time < = '"&EndDate&"')"
			RsTemp.open SqlTemp,conn,1,1
			if not RsTemp.eof then
				LineIdStr = RsTemp("LineId")
				TrStyle = "style='background:#ffffff;font-size:13px;text-align:center;'"
			end if
			RsTemp.close
			set RsTemp = nothing
		elseif ProductName = "Beijing Oem" then
			Set RsTemp = Server.CreateObject("adodb.recordset")
			SqlTemp = "select count(*) LineId from (select distinct product,line_num from bps_dr_all where product in ("&OemAll&") and LINE_NUM not in ('Total','8411','8412') and '"&StartDate&"' < = create_time AND create_time < = '"&EndDate&"')"
			RsTemp.open SqlTemp,conn,1,1

			if not RsTemp.eof then
				LineIdStr = RsTemp("LineId")
				TrStyle = "style='background:#ffffff;font-size:13px;text-align:center;'"
			end if
			RsTemp.close
			set RsTemp = nothing
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
				ProcessRealStr = "<font color='red'>" & Formatnumber(ProcessReal,2) & "%</font>"
			else
				ProcessRealStr = "<font color='green'>" & Formatnumber(ProcessReal,2) & "%</font>"
			end if
			
		AcousticTarget = Csng(RsL("AMS_TARGET"))
		AcousticReal = Csng(RsL("AMS_YIELD"))
			if AcousticReal = 0 then
				AcousticRealStr = "-"
			elseif AcousticReal < AcousticTarget then
				AcousticRealStr = "<font color='red'>" & Formatnumber(AcousticReal,2) & "%</font>"
			else
				AcousticRealStr = "<font color='green'>" & Formatnumber(AcousticReal,2) & "%</font>"
			end if
				
		CosmeticTarget = Csng(RsL("COSMETIC_TARGET"))
		CosmeticReal = Csng(RsL("COSMETIC_YIELD"))
			if CosmeticReal = 0 then
				CosmeticRealStr = "-"
			elseif CosmeticReal < CosmeticTarget then
				CosmeticRealStr = "<font color='red'>" & Formatnumber(CosmeticReal,2) & "%</font>"
			else
				CosmeticRealStr = "<font color='green'>" & Formatnumber(CosmeticReal,2) & "%</font>"
			end if
				
		FpyTarget = Csng(RsL("TTY_TARGET"))
		FpyReal = Csng(RsL("TTY_YIELD"))
			if FpyReal = 0 then
				FpyRealStr = "-"
			elseif FpyReal < FpyTarget then
				FpyRealStr = "<font color='red'>" & Formatnumber(FpyReal,2) & "%</font>"
			else
				FpyRealStr = "<font color='green'>" & Formatnumber(FpyReal,2) & "%</font>"
			end if

		if LineId <> "0" then
			LoopWrite = LoopWrite & "<tr "&TrStyle&"><td>"&ProductName&"</td>"
			LoopWrite = LoopWrite & "<td align='center'>"&LineIdStr&"</td>"
			LoopWrite = LoopWrite & "<td align='center'>"&Formatnumber(TargetPcs,0,-1)&"</td>"
			LoopWrite = LoopWrite & "<td align='center'>"&ActualPcsStr&"</td>"
			LoopWrite = LoopWrite & "<td align='center'>"&RealToWhStr&"</td>"
			LoopWrite = LoopWrite & "<td align='center'>"&Formatnumber(ProcessTarget,2)&"%</td>"
			LoopWrite = LoopWrite & "<td align='center'>"&ProcessRealStr&"</td>"
			LoopWrite = LoopWrite & "<td align='center'>"&Formatnumber(AcousticTarget,2)&"%</td>"
			LoopWrite = LoopWrite & "<td align='center'>"&AcousticRealStr&"</td>"
			LoopWrite = LoopWrite & "<td align='center'>"&Formatnumber(CosmeticTarget,2)&"%</td>"
			LoopWrite = LoopWrite & "<td align='center'>"&CosmeticRealStr&"</td>"
			LoopWrite = LoopWrite & "<td align='center'>"&Formatnumber(FpyTarget,2)&"%</td>"
			LoopWrite = LoopWrite & "<td align='center'>"&FpyRealStr&"</td><td>&nbsp;</td><td>&nbsp;</td></tr>"
		end if			 
	RsL.movenext
	loop
	RsL.close
	set RsL = nothing
end function

'循环绘制
function WriteHtml(LineType,LineName)
		
		set RsRs = Server.CreateObject("adodb.recordset")
		SqlL = "select product,COUNT(product) Lines from bps_dr_all where product in ("&LineName&") and LINE_NUM not in ('Total','8411','8412') and '"&StartDate&"' < create_time AND create_time < = '"&EndDate&"' group by product"
		RsRs.open SqlL,conn,1,1
		if not RsRs.eof then
			WriteHtml = ProdTypeTable(LineType)
			do while not RsRs.eof
				LineName1 = RsRs("product")
				Lines1 = RsRs("Lines")
				set Rs = Server.CreateObject("adodb.recordset")
				SqlStr = "select product,COUNT(product) Lines,line_num,line_type from bps_dr_all where product in ('"&LineName1&"') and LINE_NUM not in ('Total','8411','8412') and '"&StartDate&"' < create_time AND create_time < = '"&EndDate&"' group by product,line_num,line_type"
				Rs.open SqlStr,conn,1,1
				if not Rs.eof then
					ProName = rs("product")
					Lines = rs("Lines")
					LineTy = rs("line_type")
						SqlStrL = "select product,line_num,nvl(sum(TARGET),0) TARGET,nvl(sum(ACTUAL),0) ACTUAL,nvl(sum(REAL_TO_WH),0) REAL_TO_WH,nvl(sum(PQ_OUT),0) PQ_OUT,nvl(sum(PQ_START+PQ_OFFSET),0) PQ_START,nvl(avg(PQ_TARGET),0) PQ_TARGET,nvl(avg(PQ_YIELD),0) PQ_YIELD,nvl(sum(AMS_GOOD),0) AMS_GOOD,nvl(sum(AMS_START),0) AMS_START,nvl(avg(AMS_TARGET),0) AMS_TARGET,nvl(avg(AMS_YIELD),0) AMS_YIELD,nvl(sum(COSMETIC_GOOD),0) COSMETIC_GOOD,nvl(sum(COSMETIC_START),0) COSMETIC_START,nvl(avg(COSMETIC_TARGET),0) COSMETIC_TARGET,nvl(avg(COSMETIC_YIELD),0) COSMETIC_YIELD,nvl(sum(TTY_OPEN),0) TTY_OPEN,nvl(sum(TTY_CLOSE),0) TTY_CLOSE,nvl(avg(TTY_TARGET),0) TTY_TARGET,nvl(avg(TTY_YIELD),0) TTY_YIELD from bps_dr_all where product in ('"&ProName&"') and LINE_NUM not in ('Total','8411','8412') and '"&StartDate&"' < create_time AND create_time < = '"&EndDate&"' group by product,line_num"
						WriteHtml = WriteHtml & LoopWrite(SqlStrL)
						if (Cint(Lines1)/cint(Lines)) > 1 then
							if LineTy = "AUTO" then
								SqlStrT = "select product,line_num,line_type,nvl(sum(TARGET),0) TARGET,nvl(sum(ACTUAL),0) ACTUAL,nvl(sum(REAL_TO_WH),0) REAL_TO_WH,nvl(sum(PQ_OUT),0) PQ_OUT,nvl(sum(PQ_START+PQ_OFFSET),0) PQ_START,nvl(avg(PQ_TARGET),0) PQ_TARGET,nvl(avg(PQ_YIELD),0) PQ_YIELD,nvl(sum(AMS_GOOD),0) AMS_GOOD,nvl(sum(AMS_START),0) AMS_START,nvl(avg(AMS_TARGET),0) AMS_TARGET,nvl(avg(AMS_YIELD),0) AMS_YIELD,nvl(sum(COSMETIC_GOOD),0) COSMETIC_GOOD,nvl(sum(COSMETIC_START),0) COSMETIC_START,nvl(avg(COSMETIC_TARGET),0) COSMETIC_TARGET,nvl(avg(COSMETIC_YIELD),0) COSMETIC_YIELD,nvl(sum(TTY_OPEN),0) TTY_OPEN,nvl(sum(TTY_CLOSE),0) TTY_CLOSE,nvl(avg(TTY_TARGET),0) TTY_TARGET,nvl(avg(TTY_YIELD),0) TTY_YIELD from bps_dr_all where product in ('"&ProName&"') and LINE_NUM = 'Total' and '"&StartDate&"' < create_time and create_time < = '"&EndDate&"' group by product,line_num,line_type"
							else
								SqlStrT = "select product,COUNT(product) Lines,'Total' line_num,avg(pq_target) pq_target,(nvl(sum(pq_out),0))/(nvl(sum(pq_start+PQ_OFFSET),1))*100 pq_yield,avg(ams_target) ams_target,(nvl(sum(AMS_GOOD),0))/(nvl(sum(ams_start),1))*100 ams_yield,avg(cosmetic_target) cosmetic_target,(nvl(sum(cosmetic_good),0))/(nvl(sum(cosmetic_start),1))*100 cosmetic_yield,AVG(tty_target) tty_target,(nvl(sum(tty_close),0))/(nvl(sum(TTY_OPEN),1))*100 TTy_yield,nvl(sum(target),0) Target,nvl(sum(actual),0) Actual,nvl(sum(real_to_wh),0) real_to_wh,nvl(sum(pq_out),0) pq_out,nvl(sum(pq_start+PQ_OFFSET),0) pq_start,nvl(sum(AMS_GOOD),0) AMS_GOOD,nvl(sum(ams_start),0) ams_start,nvl(sum(cosmetic_good),0) cosmetic_good,nvl(SUM(cosmetic_start),0) cosmetic_start,nvl(sum(tty_close),0) tty_close,nvl(SUM(TTY_OPEN),0) TTY_OPEN from bps_dr_all where product in ('"&ProName&"') and LINE_NUM not in ('Total','8411','8412') and '"&StartDate&"' < create_time AND create_time < = '"&EndDate&"' group by product"
							end if
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
				SqlStr = "select 'Beijing' product,count(product) LINE_NUM,nvl(sum(target),0) target,nvl(sum(actual),0) actual,nvl(sum(real_to_wh),0) real_to_wh,nvl(sum(target*pq_target),0)/nvl(sum(target),1) pq_target,nvl(sum(target*AMS_target),0)/nvl(sum(target),1) AMS_target,nvl(sum(target*cosmetic_target)/sum(target),0) cosmetic_target,nvl(sum(target*TTY_target)/sum(target),0) TTY_target,nvl(sum(pq_out)/sum(pq_start)*100,0) pq_Yield,nvl(sum(ams_good)/sum(ams_start)*100,0) AMS_Yield,nvl(sum(cosmetic_good)/sum(cosmetic_start)*100,0) cosmetic_Yield,nvl(sum(tty_close)/sum(tty_open)*100,0) TTY_Yield from bps_dr_all where product in ("&BjLineName&") and LINE_NUM not in ('Total','8411','8412') and '"&StartDate&"' < create_time AND create_time < = '"&EndDate&"'"
			elseif i = 2 then
				SqlStr = "select 'Beijing Oem' product,count(product) LINE_NUM,nvl(sum(target),0) target,nvl(sum(actual),0) actual,nvl(sum(real_to_wh),0) real_to_wh,nvl(sum(target*pq_target),0)/nvl(sum(target),1) pq_target,nvl(sum(target*AMS_target),0)/nvl(sum(target),1) AMS_target,nvl(sum(target*cosmetic_target)/sum(target),0) cosmetic_target,nvl(sum(target*TTY_target)/sum(target),0) TTY_target,nvl(sum(target*pq_Yield)/sum(target),0) pq_Yield,nvl(sum(target*AMS_Yield)/sum(target),0) AMS_Yield,nvl(sum(target*cosmetic_Yield)/sum(target),0) cosmetic_Yield,nvl(sum(target*TTY_Yield)/sum(target),0) TTY_Yield from bps_dr_all where product in ("&OemAll&") and '"&StartDate&"' < create_time AND create_time < = '"&EndDate&"'"
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
	ReplaceCcl = Replace(ReplaceCcl,"371","")
	ReplaceCcl = Replace(ReplaceCcl,"521","")
	ReplaceCcl = Replace(ReplaceCcl,"801","")
	ReplaceCcl = Replace(ReplaceCcl,"911","")
end function
%>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->