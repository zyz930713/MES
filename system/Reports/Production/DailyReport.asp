<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/SendJMail.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<script type="text/javascript">
	window.opener=null;   
  	window.open("","_self");  
	window.close();
</script>
<%
reportTime=replace(dateadd("d",-1,now()),"/","-")
fromTime=formatdate(reportTime,"dd.mm.yyyy")&" 07:15"
mailTitleFromDate=formatdate(reportTime,"dd.mm.yyyy")
toTime=formatdate(now(),"dd.mm.yyyy")&" 07:14"
'fromTime="28.07.2013 07:15"
'toTime="29.07.2013 07:15"


reportDay=longweekdayconvert(weekday(reportTime))&","&formatdate(reportTime,"dd.mm.yyyy")

'mailTo="Wenjie.Lu@knowles.com;Gilbert.Zhou@knowles.com;Klaus.Fiedler@knowles.com;yan.meng@knowles.com;hui.zhang@knowles.com;Jeff.xu@knowles.com;Jane.Fu@knowles.com;zhihua.yuan@knowles.com;Ted.Wang@knowles.com;gerald.li@knowles.com;Dennis.Feng@Knowles.com;rock.shi@knowles.com;deland.diao@knowles.com;benjamin.wu@knowles.com;fanny.zhang@knowles.com;michael.qiao@knowles.com;Libing.yang@knowles.com;brent.liu@knowles.com;tim.wang@knowles.com;alex.qiu@knowles.com;Dongling.Lao@knowles.com;Junjie.Zhang@knowles.com;kuigang.bi@knowles.com;guihua.jia@knowles.com;peter.zhao@knowles.com;Roger.Wei@knowles.com;haizheng.wu@knowles.com;catherine.xu@knowles.com;zhiliang.yang@knowles.com;yinghuan.qu@knowles.com;Sonny.Zheng@knowles.com;Loren.Lei@knowles.com;Xiaoyu.Chai@knowles.com;Steven.Zhu@knowles.com;zhenqiu.lu@knowles.com;Chunli.Yu@knowles.com;benjamin.wu@knowles.com;jacy.shen@knowles.com;Hendrely.Wang@knowles.com;guoli.zhang@knowles.com;ligang.yu@knowles.com;Davide.Liu@knowles.com;Eric.Du@knowles.com;kunlun.liu@knowles.com;Vivian.huang@knowles.com;lennie.wei@knowles.com;ivan.li@knowles.com;Young.li@knowles.com;beijing.sun@knowles.com"



'MailTo="ivan.li@knowles.com;Young.li@knowles.com;fanny.zhang@knowles.com;guihua.jia@knowles.com;kunlun.liu@knowles.com;Vivian.huang@knowles.com"
mailTo="Young.li@knowles.com"
'mailto="_KEB_BPS_DailyReport@knowles.com"

htmls="<html><head><meta http-equiv='Content-Type' content='text/html; charset=utf-8'></head>"
'htmls=htmls&"<style>.td_block {color:#FFFFFF;background-color:#73A2EE;font-size: 16px;height:20px;width:120px}</style>"
htmls=htmls&"<body><table width='100%' border='1' cellpadding='0' cellspacing='0' bgcolor='#0099FF'>"
'htmls=htmls&"<font face='Arial, Helvetica, sans-serif'>"
htmls=htmls&"<tr><td height='80px'><table width='100%' border='0' cellpadding='0' cellspacing='0'>"
htmls=htmls&"<tr align='center'><th style='font-size:24px;' >Daily Production Report - KEB IT Internal Test</th><td width='30px' rowspan='2' align='right'><img src='http://"&application("HostServer")&"/Images/rpt-logo.jpg'>&nbsp;</td></tr>"
htmls=htmls&"<tr align='center' ><td><table width='80%' border='0' cellpadding='0' cellspacing='0'>"
htmls=htmls&"<tr><td><b>Report Day: &nbsp;</b>"&reportDay&"</td>"
htmls=htmls&"<td><b>Page: &nbsp;</b>1/1</td>"
htmls=htmls&"<td><b>Time:&nbsp;</b> "&fromTime&" - "&toTime&"</td></tr>"
htmls=htmls&"<tr ><td><b>Author:&nbsp;</b>Jeff Xu, Zhang Hui</td><td><b>Print Date:&nbsp;</b>"&formatdate(now(),"dd.mm.yyyy")&"</td></tr></table>"
htmls=htmls&"</td></tr></table></td></tr></table><br>"

'total
htTotal="<table width='100%' border='1' cellpadding='0' cellspacing='0' >"
htTotal=htTotal&"<tr align='center' bgcolor='#B9E9FF'><td rowspan='2'><b>Location</b></td><td rowspan='2'><b>Lines</b></td><td rowspan='2'><b>Target<br>[pcs]</b></td>"
htTotal=htTotal&"<td rowspan='2'><b>Actual<br>[pcs]</b></td><td rowspan='2'><b>Real to<br>W/H</b></td>"
htTotal=htTotal&"<td colspan='2'><b>Process Yield</b></td><td colspan='2'><b>Acoustic Yield</b></td><td colspan='2'><b>Cosmetic Yield</b></td><td colspan='2'><b>Leakage Yield</b></td>"
htTotal=htTotal&"<td colspan='2'><b>TTY</b></td><td colspan='2'><b>Job Quantity</b></td></tr>"
htTotal=htTotal&"<tr align='center' bgcolor='#B9E9FF'><td ><b>Target</b></td><td ><b>Real</b></td><td ><b>Target</b></td><td ><b>Real</b></td><td ><b>Target</b></td>"
htTotal=htTotal&"<td ><b>Real</b></td><td ><b>Target</b></td><td ><b>Real</b></td><td ><b>Target</b></td><td ><b>Real</b></td><td ><b>Job Open</b></td><td ><b>Job Close</b></td></tr>"


'detail
htDetl="<table width='100%' border='1' cellpadding='0' cellspacing='0' ><tr align='center' bgcolor='#B9E9FF'><td colspan='17'><b>Details Beijing(Semi):</td></tr>"
htDetl=htDetl&"<tr align='center' bgcolor='#B9E9FF'><td rowspan='2'><b>Product</b></td><td rowspan='2'><b>Line</b></td><td rowspan='2'><b>Target<br>[pcs]</b></td>"
htDetl=htDetl&"<td rowspan='2'><b>Actual<br>[pcs]</b></td><td rowspan='2'><b>Real to<br>W/H</b></td>"
htDetl=htDetl&"<td colspan='2'><b>Process Yield</b></td><td colspan='2'><b>Acoustic Yield</b></td><td colspan='2'><b>Cosmetic Yield</b></td><td colspan='2'><b>Leakage Yield</b>"
htDetl=htDetl&"<td colspan='2'><b>TTY</b></td><td colspan='2'><b>Job Quantity</b></td></tr>"
htDetl=htDetl&"<tr align='center' bgcolor='#B9E9FF'><td ><b>Target</b></td><td ><b>Real</b></td><td ><b>Target</b></td><td ><b>Real</b></td><td ><b>Target</b></td>"
htDetl=htDetl&"<td ><b>Real</b></td><td ><b>Target</b></td><td ><b>Real</b></td><td ><b>Target</b></td><td ><b>Real</b></td><td ><b>Job Open</b></td><td ><b>Job Close</b></td></tr>"

'Description
htmlDesc="<br><table width='100%' border='1' cellpadding='0'>"
htmlDesc=htmlDesc&"<tr><td colspan='2' align='center' bgcolor='#B9E9FF' ></b>Description</b></td></tr>"
htmlDesc=htmlDesc&"<tr><td width='60px'>Target[pcs]</td><td style='font-size:14px;'>Set by production</td></tr>"
htmlDesc=htmlDesc&"<tr><td width='60px'>Actual[pcs]</td><td style='font-size:14px;'>Pre-Store. last day 24H output time form 7:15 to 7:15 </td></tr>"
htmlDesc=htmlDesc&"<tr><td width='60px'>Real to W/H</td><td style='font-size:14px;'>Warehouse Actual Received Qty,  last day 24H send to WH QTY . time from 7:15 to 7:15</td></tr>"
htmlDesc=htmlDesc&"<tr><td width='60px'>Process Yield</td><td style='font-size:14px;'> Hektor : Main line  , Winslet : HFL , (Finished Assy. Good Qty / Start Qty) last day 24H output time form 7:15 to 7:15 </td></tr>"
htmlDesc=htmlDesc&"<tr><td width='60px'>Acoustic Yield</td><td style='font-size:14px;'>Acoustic Test (Acoustic Pass / Acoustic Input) last day 24H output time form 7:15 to 7:15</td></tr>"
htmlDesc=htmlDesc&"<tr><td width='60px'>Cosmetic Yield</td><td style='font-size:14px;'>Cosmetic Yield, (Cosmetic Pass / Cosmetic Input) last day 24H output time form 7:15 to 7:15</td></tr>"
htmlDesc=htmlDesc&"<tr><td width='60px'>TTY</td><td style='font-size:14px;'>Total Yield, (Job Close Quantity / Job Open Quantity) last day 24H completed all JO yield</td></tr>"
htmlDesc=htmlDesc&"<tr><td width='60px'>Job Qty</td><td style='font-size:14px;'>Job Qty, Job closed quantity and the relevant open quantity</td></tr>"
htmlDesc=htmlDesc&"<tr><td width='60px'>Remarks</td><td style='font-size:14px;'>This reprot not included rework  JO and QYT </td></tr></table>"
set rsReal=server.createobject("adodb.recordset")
lines=0
tempProduct=""
totalTarget=0
subTotalTarget=0
totalActual=0
subTotalAct=0
totalWH=0
subWH=0
totalProcTargetQty=0
subProcTargetQty=0
totalProcJobQty=0
subProcJobQty=0
totalProcQty=0
subProcQty=0
totalAcousticTargetQty=0
subAcousticTargetQty=0
totalAcousticJobQty=0
subAcousticJobQty=0
totalAcousticQty=0
subAcousticQty=0
totalFoiTargetQty=0
subFoiTargetQty=0
totalFoiJobQty=0
subFoiJobQty=0
totalFoiQty=0
subFoiQty=0
totalFpyTargetQty=0
subFpyTargetQty=0
totalFpyJobQty=0
subFpyJobQty=0
totalFpyQty=0
subFpyQty=0
totalTholdQty=0
subTHoldQty=0
totalHoldQty=0
subHoldQty=0


'1.	Target set by Engineering
'2.	Actual - Pre-Store
'3.	Real-to W/H - Warehouse Receiving Qty
'4.	Process Yield - 2D Printing and Before
'5.	Acoustic Yield - Acoustic Test
'6.	Cosmetic Yield - FOI Yield
'7.	FPY - First Pass Yield
'8.	Blocked [kpcs] - Hold Qty

'get target setting

'------------------------------------------------------Details Beijing(Semi):Total------------------------------------
sql="SELECT PRODUCT,LINE,TARGET,TARGET*PROCESS_YIELD_TARGET AS PROCESS_TARGET,PROCESS_YIELD_TARGET,TARGET*ACOUSTIC_YIELD_TARGET AS ACOUSTIC_TARGET,ACOUSTIC_YIELD_TARGET,TARGET*FOI_YIELD_TARGET AS FOI_TARGET,FOI_YIELD_TARGET,TARGET*LEAKAGE_YIELD_TARGET AS LEAKAGE_TARGET,LEAKAGE_YIELD_TARGET,TARGET*FPY_TARGET AS FPY_QTY,FPY_TARGET FROM RPT_DAILY_TARGET  where  PRODUCT not in('Maple','Goldberg','Marigold','Pansy')  ORDER BY PRODUCT,LINE  "
rs.open sql,conn,1,3
if rs.eof then
	'does not set config data
	response.End()
else
for i=0 to rs.recordcount 
	if i=0 then
		tempProduct = rs("PRODUCT")
	end if
	if not rs.eof then
		product=rs("PRODUCT")
	end if
	'show sub total
	if rs.eof or tempProduct <> product then
		htDetl=htDetl&"<tr align='center'><td><b>"&tempProduct&"</b></td><td><b>Total</b></td><td><b>"&subTotalTarget&"</br></td><td>"
		if subTotalAct>0 then
			if subTotalAct < subTotalTarget then
				fontColor="red"
			else
				fontColor="green"
			end if
			htDetl=htDetl&"<font color='"&fontColor&"' ><b>"&subTotalAct&"</b></font>"	
		end if
		htDetl=htDetl&"&nbsp;</td><td>"
		if subWH>0 then
			htDetl=htDetl&"<B>"&subWH&"</b>"
		end if
		htDetl=htDetl&"&nbsp;</td>"
		
		subProcTargetYield=subProcTargetQty/subTotalTarget
		htDetl=htDetl&"<td><b>"&round(subProcTargetYield*100,2)&"%</b></td><td>"
		if subProcJobQty>0 then		
			subProcYield=subProcQty/subProcJobQty				
			if subProcYield < subProcTargetYield then
				fontColor="red"
			else
				fontColor="green"
			end if
			htDetl=htDetl&"<font color='"&fontColor&"' ><b>"&round(subProcYield*100,2)&"%</b></font>"
		end if		
		htDetl=htDetl&"&nbsp;</td>"
		
		subAcousticTargetYield=subAcousticTargetQty/subTotalTarget
		htDetl=htDetl&"<td><b>"&round(subAcousticTargetYield*100,2)&"%</b></td><td>"
		if subAcousticJobQty>0 then
			subAcousticYield=subAcousticQty/subAcousticJobQty				
			if subAcousticYield < subAcousticTargetYield then
				fontColor="red"
			else
				fontColor="green"
			end if
			htDetl=htDetl&"<font color='"&fontColor&"' ><b>"&round(subAcousticYield*100,2)&"%</b></font>"
		end if
		htDetl=htDetl&"&nbsp;</td>"
				
		subFoiTargetYield=subFoiTargetQty/subTotalTarget
		htDetl=htDetl&"<td><b>"&round(subFoiTargetYield*100,2)&"%</b></td><td>"
		if subFoiJobQty>0 then
			subFoiYield=subFoiQty/subFoiJobQty				
			if subFoiYield < subFoiTargetYield then
				fontColor="red"
			else
				fontColor="green"
			end if
			htDetl=htDetl&"<font color='"&fontColor&"' ><b>"&round(subFoiYield*100,2)&"%</b></font>"
		end if
		htDetl=htDetl&"&nbsp;</td>"
		
		
		
		'-----------------------Leakage Yield ---------------
		subLeakageTargetYield=subLeakageTargetQty/subTotalTarget
		htDetl=htDetl&"<td><b>"&round(subLeakageTargetYield*100,2)&"%</b></td><td>"
		if subLeakageJobQty>0 then
			subLeakageYield=subLeakageQty/subLeakageJobQty				
			if subLeakageYield < subLeakageTargetYield then
				fontColor="red"
			else
				fontColor="green"
			end if
			htDetl=htDetl&"<font color='"&fontColor&"' ><b>"&round(subLeakageYield*100,2)&"%</b></font>"
		end if
		htDetl=htDetl&"&nbsp;</td>"
		
	
		
		
		'-------------------------------------
				
		subFpyTargetYield=subFpyTargetQty/subTotalTarget
		htDetl=htDetl&"<td><b>"&round(subFpyTargetYield*100,2)&"%</b></td><td>"
		if subFpyJobQty >0 then
			subFpyYield=subFpyQty/subFpyJobQty				
			if subFpyYield < subFpyTargetYield then
				fontColor="red"
			else
				fontColor="green"
			end if
			htDetl=htDetl&"<font color='"&fontColor&"' ><b>"&round(subFpyYield*100,2)&"%</b></font>"
		end if
		htDetl=htDetl&"&nbsp;</td><td>"	
		if subHoldQty>0 then
			htDetl=htDetl&subHoldQty
		end if
		'htDetl=htDetl&"&nbsp;</td><td>"	
		htDetl=htDetl&"<font color='"&fontColor&"' ><b>"&subFpyJobQty&"</b></font>"   'Total Job Close
		htDetl=htDetl&"&nbsp;</td><td>"	
		if subTHoldQty>0 then
			htDetl=htDetl&subTHoldQty
		end if
		'htDetl=htDetl&"&nbsp;</td></tr>"
		htDetl=htDetl&"<font color='"&fontColor&"' ><b>"&subFpyQty&"</b></font>"      'Total Job Open
		htDetl=htDetl&"&nbsp;</td></tr>"	
		if rs.eof then
			exit for
		end if
		tempProduct=rs("PRODUCT")
		subTotalTarget=0
		subTotalAct=0
		subWH=0
		subProcTargetQty=0
		subProcJobQty=0
		subProcQty=0
		subAcousticTargetQty=0
		subAcousticJobQty=0
		subAcousticQty=0
		subFoiTargetQty=0
		subFoiJobQty=0
		subFoiQty=0
		subFpyTargetQty=0
		subFpyJobQty=0
		subFpyQty=0
		subHoldQty=0
		subTHoldQty=0				
	end if
	
	totalTarget=totalTarget+csng(rs("TARGET"))
	subTotalTarget=subTotalTarget+csng(rs("TARGET"))
	htDetl=htDetl&"<tr align='center'><td>"&product&"</td><td>"&"LINE"&"&nbsp;#"&right(rs("LINE"),1)&"</td><td>"&rs("TARGET")&"</td><td>"
	'get actual qty
	sql="select sum(store_quantity) as actualQty from job_master_store_pre where store_quantity >0 and store_time between to_date('"&fromTime&"','dd.mm.yyyy hh24:mi') and to_date('"&toTime&"','dd.mm.yyyy hh24:mi') and line_name ='"&rs("LINE")&"'  AND substr(job_number,1,3)<>'RWK'"

	rsReal.open sql,conn,1,3
	if not rsReal.eof and not isnull(rsReal("actualQty")) then
		totalActual=totalActual+csng(rsReal("actualQty"))
		subTotalAct=subTotalAct+csng(rsReal("actualQty"))
		if csng(rsReal("actualQty")) < csng(rs("TARGET")) then
			fontColor="red"
		else
			fontColor="green"
		end if
		htDetl=htDetl&"<font color='"&fontColor&"'>"&rsReal("actualQty")&"</font>"
	end if
	htDetl=htDetl&"&nbsp;</td><td>"
	rsReal.close
'------------------------------------------------------------------------------------------	
	'get real to W/H qty
	'sql="select sum(packed_qty) as toWHQty from job_pack_detail where packed_line ='"&rs("LINE")&"' and WHREC_time between to_date('"&fromTime&"','dd.mm.yyyy hh24:mi') and to_date('"&toTime&"','dd.mm.yyyy hh24:mi')"
	sql="select sum(packed_qty) as toWHQty from job_pack_detail where packed_line ='"&rs("LINE")&"' and WHREC_time between to_date('"&fromTime&"','dd.mm.yyyy hh24:mi') and to_date('"&toTime&"','dd.mm.yyyy hh24:mi') AND  substr(job_number,1,3)<>'RWK' and substr(box_id,1,2)='FG' and get_JSQ is null"
	rsReal.open sql,conn,1,3
	if not rsReal.eof and not isnull(rsReal("toWHQty")) then
		totalWH=totalWH+csng(rsReal("toWHQty"))			
		subWH=subWH+csng(rsReal("toWHQty"))
		htDetl=htDetl&rsReal("toWHQty")
	end if
	htDetl=htDetl&"&nbsp;</td>"
	rsReal.close
'------------------------------------------------------------------------------------------	
	'get process_yield 
	htDetl=htDetl&"<td>"&round(csng(rs("PROCESS_YIELD_TARGET"))*100,2)&"%</td><td>"
	totalProcTargetQty=totalProcTargetQty+csng(rs("PROCESS_TARGET"))
	subProcTargetQty=subProcTargetQty+csng(rs("PROCESS_TARGET"))	
	
	if tempProduct = "Winslet" then 
		mother_station_id_set="SA00001040"
		
	    sql="SELECT SUM(b.job_start_quantity) AS proc_job_qty  FROM job_stations a,job b WHERE a.job_number=b.job_number and a.sheet_number=b.sheet_number AND a.close_time BETWEEN to_date('"&fromTime&"','dd.mm.yyyy hh24:mi') and to_date('"&toTime&"','dd.mm.yyyy hh24:mi') and a.station_id in(select nid from station where mother_station_id='SA00000988') and b.line_name ='"&rs("LINE")&"' and a.good_quantity='0'	"
		rsReal.open sql,conn,1,3
		proc_job_qty690=rsReal("proc_job_qty")
		rsReal.close
		
	else
		mother_station_id_set="SA00000988"
	end if
	if  proc_job_qty690="" or isnull(proc_job_qty690) then
	proc_job_qty690=0
	end if
	
	
	sql="SELECT SUM(a.good_quantity) AS proc_qty,SUM(b.job_start_quantity) AS proc_job_qty FROM job_stations a,job b WHERE a.job_number=b.job_number and a.sheet_number=b.sheet_number AND a.close_time BETWEEN to_date('"&fromTime&"','dd.mm.yyyy hh24:mi') and to_date('"&toTime&"','dd.mm.yyyy hh24:mi') and a.station_id in(select nid from station where mother_station_id='"&mother_station_id_set&"') and b.line_name ='"&rs("LINE")&"'"
	rsReal.open sql,conn,1,3	
	if not rsReal.eof and not isnull(rsReal("proc_job_qty")) then		
		totalProcQty=totalProcQty+csng(rsReal("proc_qty"))
		subProcQty=subProcQty+csng(rsReal("proc_qty"))
		totalProcJobQty=totalProcJobQty+(csng(rsReal("proc_job_qty"))+csng(proc_job_qty690))
		subProcJobQty=subProcJobQty+(csng(rsReal("proc_job_qty"))+csng(proc_job_qty690))
		yield=csng(rsReal("proc_qty"))/(csng(rsReal("proc_job_qty"))+csng(proc_job_qty690))
		if yield < csng(rs("PROCESS_YIELD_TARGET")) then
			fontColor="red"
		else
			fontColor="green"
		end if
		htDetl=htDetl&"<font color='"&fontColor&"'>"&round(yield*100,2)&"%</font>"
	end if
	htDetl=htDetl&"&nbsp;</td>"
	rsReal.close
	'------------------------------------------------------------------------------------------
	'get Acoustic_Yield 
	htDetl=htDetl&"<td>"&round(csng(rs("ACOUSTIC_YIELD_TARGET"))*100,2)&"%</td><td>"
	totalAcousticTargetQty=totalAcousticTargetQty+csng(rs("ACOUSTIC_TARGET"))
	subAcousticTargetQty=subAcousticTargetQty+csng(rs("ACOUSTIC_TARGET"))
	'sql="select sum(a.good_quantity) as acoustic_qty,sum(a.station_start_quantity) as acoustic_job_qty from job_stations a,job_master b where a.job_number=b.job_number and a.close_time between to_date('"&fromTime&"','dd.mm.yyyy hh24:mi') and to_date('"&toTime&"','dd.mm.yyyy hh24:mi') and  a.station_id in (select nid from station where mother_station_id='SA00000989') and b.line_name ='"&rs("LINE")&"'"
	
	sql="select sum(a.good_quantity) as acoustic_qty,sum(a.station_start_quantity) as acoustic_job_qty from job_stations a,job_master b where a.job_number=b.job_number and substr(a.job_number,1,3)<>'RWK'  and a.sheet_number<500 and a.close_time between to_date('"&fromTime&"','dd.mm.yyyy hh24:mi') and to_date('"&toTime&"','dd.mm.yyyy hh24:mi') and  a.station_id in (select nid from station where mother_station_id='SA00000989') and b.line_name ='"&rs("LINE")&"'"
	rsReal.open sql,conn,1,3
	if not rsReal.eof and not isnull(rsReal("acoustic_job_qty")) then		
		totalAcousticQty=totalAcousticQty+csng(rsReal("acoustic_qty"))
		subAcousticQty=subAcousticQty+csng(rsReal("acoustic_qty"))
		totalAcousticJobQty=totalAcousticJobQty+csng(rsReal("acoustic_job_qty"))
		subAcousticJobQty=subAcousticJobQty+csng(rsReal("acoustic_job_qty"))
		yield=csng(rsReal("acoustic_qty"))/csng(rsReal("acoustic_job_qty"))
		if yield < csng(rs("ACOUSTIC_YIELD_TARGET")) then
			fontColor="red"
		else

			fontColor="green"
		end if
		htDetl=htDetl&"<font color='"&fontColor&"'>"&round(yield*100,2)&"%</font>"
	end if
	htDetl=htDetl&"&nbsp;</td>"
	rsReal.close
'------------------------------------------------------------------------------------------
	'get Foi_Yield 
	htDetl=htDetl&"<td>"&round(csng(rs("FOI_YIELD_TARGET"))*100,2)&"%</td><td>"
	totalFoiTargetQty=totalFoiTargetQty+csng(rs("FOI_TARGET"))
	subFoiTargetQty=subFoiTargetQty+csng(rs("FOI_TARGET"))
	'sql="select sum(a.good_quantity) as foi_qty,sum(a.station_start_quantity) as foi_job_qty from job_stations a,job_master b where a.job_number=b.job_number and a.close_time between to_date('"&fromTime&"','dd.mm.yyyy hh24:mi') and to_date('"&toTime&"','dd.mm.yyyy hh24:mi') and  a.station_id in (select nid from station where mother_station_id='SA00000990') and b.line_name ='"&rs("LINE")&"'"
	sql="select sum(a.good_quantity) as foi_qty,sum(a.station_start_quantity) as foi_job_qty from job_stations a,job_master b where a.job_number=b.job_number and substr(a.job_number,1,3)<>'RWK'  and a.sheet_number<500 and a.close_time between to_date('"&fromTime&"','dd.mm.yyyy hh24:mi') and to_date('"&toTime&"','dd.mm.yyyy hh24:mi') and  a.station_id in (select nid from station where mother_station_id='SA00000990') and b.line_name ='"&rs("LINE")&"'"
	rsReal.open sql,conn,1,3
	if not rsReal.eof and not isnull(rsReal("foi_job_qty")) then		
		totalFoiQty=totalFoiQty+csng(rsReal("foi_qty"))
		subFoiQty=subFoiQty+csng(rsReal("foi_qty"))
		totalFoiJobQty=totalFoiJobQty+csng(rsReal("foi_job_qty"))
		subFoiJobQty=subFoiJobQty+csng(rsReal("foi_job_qty"))
		yield=csng(rsReal("foi_qty"))/csng(rsReal("foi_job_qty"))
		if yield < csng(rs("FOI_YIELD_TARGET")) then
			fontColor="red"
		else
			fontColor="green"
		end if
		htDetl=htDetl&"<font color='"&fontColor&"'>"&round(yield*100,2)&"%</font>"
	end if
	htDetl=htDetl&"&nbsp;</td>"
	rsReal.close		
	
	'------------------------------------------------------------------------------------------
	'get Leakage
	
	htDetl=htDetl&"<td>"&round(csng(rs("LEAKAGE_YIELD_TARGET"))*100,2)&"%</td><td>"
	totalLeakageTargetQty=totalLeakageTargetQty+csng(rs("LEAKAGE_TARGET"))
	subLeakageTargetQty=subLeakageTargetQty+csng(rs("LEAKAGE_TARGET"))
	'sql="select sum(a.good_quantity) as foi_qty,sum(a.station_start_quantity) as foi_job_qty from job_stations a,job_master b where a.job_number=b.job_number and a.close_time between to_date('"&fromTime&"','dd.mm.yyyy hh24:mi') and to_date('"&toTime&"','dd.mm.yyyy hh24:mi') and  a.station_id in (select nid from station where mother_station_id='SA00000990') and b.line_name ='"&rs("LINE")&"'"
	sql="select sum(a.good_quantity) as LEAKAGE_qty,sum(a.station_start_quantity) as LEAKAGE_job_qty from job_stations a,job_master b where a.job_number=b.job_number and substr(a.job_number,1,3)<>'RWK'  and a.sheet_number<500 and a.close_time between to_date('"&fromTime&"','dd.mm.yyyy hh24:mi') and to_date('"&toTime&"','dd.mm.yyyy hh24:mi') and  a.station_id in (select nid from station where mother_station_id='SA00001060') and b.line_name ='"&rs("LINE")&"'"
	
	rsReal.open sql,conn,1,3
	if not rsReal.eof and not isnull(rsReal("LEAKAGE_job_qty")) then		
		totalLeakageQty=totalLeakageQty+csng(rsReal("LEAKAGE_qty"))
		subLeakageQty=subLeakageQty+csng(rsReal("LEAKAGE_qty"))
		totalLeakageJobQty=totalLeakageJobQty+csng(rsReal("LEAKAGE_job_qty"))
		subLeakageJobQty=subLeakageJobQty+csng(rsReal("LEAKAGE_job_qty"))
		yield=csng(rsReal("LEAKAGE_qty"))/csng(rsReal("LEAKAGE_job_qty"))
		if yield < csng(rs("LEAKAGE_YIELD_TARGET")) then
			fontColor="red"
		else
			fontColor="green"
		end if
		htDetl=htDetl&"<font color='"&fontColor&"'>"&round(yield*100,2)&"%</font>"
	end if
	htDetl=htDetl&"&nbsp;</td>"
	rsReal.close 

	
	
	
'------------------------------------------------------------------------------------------
	'get fpy 
	htDetl=htDetl&"<td>"&round(csng(rs("FPY_TARGET"))*100,2)&"%</td><td>"
	totalFpyTargetQty=totalFpyTargetQty+csng(rs("FPY_QTY"))
	subFpyTargetQty=subFpyTargetQty+csng(rs("FPY_QTY"))
	'sql="select sum(job_good_quantity) as fpy_qty,sum(job_start_quantity) as fpy_job_qty,round(sum(job_good_quantity)/sum(job_start_quantity),4) as fpy from job  where   close_time between to_date('"&fromTime&"','dd.mm.yyyy hh24:mi') and to_date('"&toTime&"','dd.mm.yyyy hh24:mi') and line_name ='"&rs("LINE")&"'"
	sql="select sum(job_good_quantity) as fpy_qty,sum(job_start_quantity) as fpy_job_qty,round(sum(job_good_quantity)/sum(job_start_quantity),4) as fpy from job  where substr(job_number,1,3)<>'RWK'  and sheet_number<500 and   close_time between to_date('"&fromTime&"','dd.mm.yyyy hh24:mi') and to_date('"&toTime&"','dd.mm.yyyy hh24:mi') and line_name ='"&rs("LINE")&"'"
	rsReal.open sql,conn,1,3
	if not rsReal.eof and not isnull(rsReal("fpy_job_qty")) then		
		totalFpyQty=totalFpyQty+csng(rsReal("fpy_qty"))
		subFpyQty=subFpyQty+csng(rsReal("fpy_qty"))
		uniqFpyQty=csng(rsReal("fpy_qty"))
		totalFpyJobQty=totalFpyJobQty+csng(rsReal("fpy_job_qty"))
		subFpyJobQty=subFpyJobQty+csng(rsReal("fpy_job_qty"))
		uniqFpyJobQty=csng(rsReal("fpy_job_qty"))
		yield=csng(rsReal("fpy_qty"))/csng(rsReal("fpy_job_qty"))
		if yield < csng(rs("FPY_TARGET")) then
			fontColor="red"
		else
			fontColor="green"
		end if
		htDetl=htDetl&"<font color='"&fontColor&"'>"&round(yield*100,2)&"%</font>"
	else
		uniqFpyQty=""
		uniqFpyJobQty=""
	end if
	htDetl=htDetl&"&nbsp;</td><td>"
	rsReal.close	
	
	'get &/- 24h hold qty
	'Change to Open Quantity
	sql="select sum(job_good_quantity) as holdQty from job a,job_hold_release_history b where a.job_number=b.job_number and a.sheet_number =b.sheet_number and b.transaction_time between to_date('"&fromTime&"','dd.mm.yyyy hh24:mi')-1 and to_date('"&toTime&"','dd.mm.yyyy hh24:mi')+1 and status=2 and line_name='"&rs("LINE")&"'"
	rsReal.open sql,conn,1,3
	if not rsReal.eof and not isnull(rsReal("holdQty")) then
		totalHoldQty=totalHoldQty+csng(rsReal("holdQty"))
		subHoldQty=subHoldQty+csng(rsReal("holdQty"))
		htDetl=htDetl&rsReal("holdQty")
	end if
	rsReal.close
	htDetl=htDetl&"<font color='"&fontColor&"'>"&uniqFpyJobQty&"</font>"  'semi job close
	'htDetl=htDetl&"&"subFpyJobQty&"</td><td>"
	htDetl=htDetl&"&nbsp;</td><td>"
	'get total hold qty
	'Change to Close Quantity
	sql="select sum(job_good_quantity)/1000 as totalHoldQty from job where status=2 and line_name='"&rs("LINE")&"'"
	rsReal.open sql,conn,1,3
	if not rsReal.eof and not isnull(rsReal("totalHoldQty")) then
		totalTHoldQty=totalHoldQty+csng(rsReal("totalHoldQty"))
		subTHoldQty=subTHoldQty+csng(rsReal("totalHoldQty"))
		htDetl=htDetl&rsReal("totalHoldQty")
	end if
	rsReal.close
	'htDetl=htDetl&"&nbsp;fffff</td></tr>" 
	htDetl=htDetl&"<font color='"&fontColor&"'>"&uniqFpyQty&"</font>"  'semi job open
	htDetl=htDetl&"&nbsp;</td></tr>"
	rs.movenext
	lines=lines+1
next
end if
rs.close

htDetl=htDetl&"</table>"



'------------------------------

subHTTotal="<td>"&lines&"</td><td>"&totalTarget&"</td><td>"
if totalActual>0 then
	if totalActual < totalTarget then
		fontColor="red"
	else
		fontColor="green"
	end if
	subHTTotal=subHTTotal&"<font color='"&fontColor&"' >"&totalActual&"</font>"
end if
subHTTotal=subHTTotal&"&nbsp;</td><td>"
if totalWH>0 then
	subHTTotal=subHTTotal&totalWH
end if
subHTTotal=subHTTotal&"&nbsp;</td>"
targetYield=totalProcTargetQty/totalTarget
subHTTotal=subHTTotal&"<td>"&round(targetYield*100,2)&"%</td><td>"
if totalProcJobQty >0 then
	realYield=totalProcQty/totalProcJobQty				
	if realYield < targetYield then
		fontColor="red"
	else
		fontColor="green"
	end if
	subHTTotal=subHTTotal&"<font color='"&fontColor&"' >"&round(realYield*100,2)&"%</font>"	
end if
subHTTotal=subHTTotal&"&nbsp;</td>"

targetYield=totalAcousticTargetQty/totalTarget
subHTTotal=subHTTotal&"<td>"&round(targetYield*100,2)&"%</td><td>"
if totalAcousticJobQty>0 then
	realYield=totalAcousticQty/totalAcousticJobQty		
	if realYield < targetYield then
		fontColor="red"
	else
		fontColor="green"
	end if
	subHTTotal=subHTTotal&"<font color='"&fontColor&"' >"&round(realYield*100,2)&"%</font>"	
end if
subHTTotal=subHTTotal&"&nbsp;</td>"

targetYield=totalFoiTargetQty/totalTarget
subHTTotal=subHTTotal&"<td>"&round(targetYield*100,2)&"%</td><td>"
if totalFoiJobQty >0 then
	realYield=totalFoiQty/totalFoiJobQty	
	if realYield < targetYield then
		fontColor="red"
	else
		fontColor="green"
	end if
	subHTTotal=subHTTotal&"<font color='"&fontColor&"' >"&round(realYield*100,2)&"%</font>"	
end if
subHTTotal=subHTTotal&"&nbsp;</td>"

'------------------------------Leakage  ---------
targetYield=totalLeakageTargetQty/totalTarget
subHTTotal=subHTTotal&"<td>"&round(targetYield*100,2)&"%</td><td>"
if totalLeakageJobQty >0 then
	realYield=totalLeakageQty/totalLeakageJobQty	
	if realYield < targetYield then
		fontColor="red"
	else
		fontColor="green"
	end if
	subHTTotal=subHTTotal&"<font color='"&fontColor&"' >"&round(realYield*100,2)&"%</font>"	
end if
subHTTotal=subHTTotal&"&nbsp;</td>"



'---------------------------------------







targetYield=totalFpyTargetQty/totalTarget
subHTTotal=subHTTotal&"<td>"&round(targetYield*100,2)&"%</td><td>"
if totalFpyJobQty>0 then
	realYield=totalFpyQty/totalFpyJobQty				
	if realYield < targetYield then
		fontColor="red"
	else
		fontColor="green"
	end if
	subHTTotal=subHTTotal&"<font color='"&fontColor&"' >"&round(realYield*100,2)&"%</font>"	
end if
subHTTotal=subHTTotal&"&nbsp;</td><td>"












if totalHoldQty>0 then
	subHTTotal=subHTTotal&totalHoldQty
end if
'subHTTotal=subHTTotal&"&nbsp;a</td><td>"
subHTTotal=subHTTotal&"<font color='"&fontColor&"' >"&totalFpyJobQty&"</font>"	'beijing total  job close
subHTTotal=subHTTotal&"&nbsp;</td><td>"

if totalTHoldQty>0 then
	subHTTotal=subHTTotal&totalTHoldQty
end if


'subHTTotal=subHTTotal&"&nbsp;b</td></tr>"
subHTTotal=subHTTotal&"<font color='"&fontColor&"' >"&totalFpyQty&"</font>"	   'beijing total  job open
subHTTotal=subHTTotal&"&nbsp;</td></tr>"



'-------------------------------------------------------'

subHTTotalNew="<td><B>"&lines&"</b></td><td><b>"&totalTarget&"</b></td><td>"
if totalActual>0 then
	if totalActual < totalTarget then
		fontColor="red"
	else
		fontColor="green"
	end if
	subHTTotalNew=subHTTotalNew&"<font color='"&fontColor&"' ><B>"&totalActual&"</b></font>"
end if
subHTTotalNew=subHTTotalNew&"&nbsp;</td><td>"
if totalWH>0 then
	subHTTotalNew=subHTTotalNew&"<B>"&totalWH&"</B>"
end if
subHTTotalNew=subHTTotalNew&"&nbsp;</td>"
targetYield=totalProcTargetQty/totalTarget
subHTTotalNew=subHTTotalNew&"<td><B>"&round(targetYield*100,2)&"%</b></td><td>"
if totalProcJobQty >0 then
	realYield=totalProcQty/totalProcJobQty				
	if realYield < targetYield then
		fontColor="red"
	else
		fontColor="green"
	end if
	subHTTotalNew=subHTTotalNew&"<font color='"&fontColor&"' >"&round(realYield*100,2)&"%</font>"	
end if
subHTTotalNew=subHTTotalNew&"&nbsp;</td>"

targetYield=totalAcousticTargetQty/totalTarget
subHTTotalNew=subHTTotalNew&"<td><B>"&round(targetYield*100,2)&"%</b></td><td>"
if totalAcousticJobQty>0 then
	realYield=totalAcousticQty/totalAcousticJobQty		
	if realYield < targetYield then
		fontColor="red"
	else
		fontColor="green"
	end if
	subHTTotalNew=subHTTotalNew&"<font color='"&fontColor&"' ><B>"&round(realYield*100,2)&"%</b></font>"	
end if
subHTTotalNew=subHTTotalNew&"&nbsp;</td>"

targetYield=totalFoiTargetQty/totalTarget
subHTTotalNew=subHTTotalNew&"<td><b>"&round(targetYield*100,2)&"%</b></td><td>"
if totalFoiJobQty >0 then
	realYield=totalFoiQty/totalFoiJobQty	
	if realYield < targetYield then
		fontColor="red"
	else
		fontColor="green"
	end if
	subHTTotalNew=subHTTotalNew&"<font color='"&fontColor&"' ><b>"&round(realYield*100,2)&"%</b></font>"	
end if
subHTTotalNew=subHTTotalNew&"&nbsp;</td>"


'----------------------------------------

targetYield=totalLeakageTargetQty/totalTarget
subHTTotalNew=subHTTotalNew&"<td><b>"&round(targetYield*100,2)&"%</b></td><td>"
if totalFoiJobQty >0 then
	realYield=totalLeakageQty/totalLeakageJobQty	
	if realYield < targetYield then
		fontColor="red"
	else
		fontColor="green"
	end if
	subHTTotalNew=subHTTotalNew&"<font color='"&fontColor&"' ><b>"&round(realYield*100,2)&"%</b></font>"	
end if
subHTTotalNew=subHTTotalNew&"&nbsp;</td>"

'----------------------------------------








targetYield=totalFpyTargetQty/totalTarget
subHTTotalNew=subHTTotalNew&"<td><b>"&round(targetYield*100,2)&"%</b></td><td>"
if totalFpyJobQty>0 then
	realYield=totalFpyQty/totalFpyJobQty				
	if realYield < targetYield then
		fontColor="red"
	else
		fontColor="green"
	end if
	subHTTotalNew=subHTTotalNew&"<font color='"&fontColor&"' ><b>"&round(realYield*100,2)&"%</b></font>"	
end if

subHTTotalNew=subHTTotalNew&"&nbsp;</td><td>"

if totalHoldQty>0 then
	subHTTotalNew=subHTTotalNew&totalHoldQty
end if
'subHTTotalNew=subHTTotalNew&"&nbsp;a</td><td>"
subHTTotalNew=subHTTotalNew&"<font color='"&fontColor&"' ><b>"&totalFpyJobQty&"</b></font>"	'beijing total  job close
subHTTotalNew=subHTTotalNew&"&nbsp;</td><td>"

if totalTHoldQty>0 then
	subHTTotalNew=subHTTotalNew&totalTHoldQty
end if


'subHTTotalNew=subHTTotalNew&"&nbsp;b</td></tr>"
subHTTotalNew=subHTTotalNew&"<font color='"&fontColor&"' ><b>"&totalFpyQty&"</b></font>"	   'beijing total  job open
subHTTotalNew=subHTTotalNew&"&nbsp;</td></tr>"

'--------------------------------------------------------'











htTotal=htTotal&"<tr align='center'><td>Beijing</td>"&subHTTotal&"<tr align='center'><td><b>Total</b></td>"&subHTTotalNew&"</table><br>"

mailContents=htmls&htTotal&htDetl&htmlDesc&"</body></html>"
'response.Write(mailContents)

SendJMail application("MailSender"),mailTo,"Daily Production Internal Test Report (Process Yield Cut Point Winslet 691, Other 690) "&mailTitleFromDate,mailContents
%>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->