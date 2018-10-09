<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->

<!--#include virtual="/Functions/SendJMail.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->

<%
reportTime=replace(dateadd("d",-1,now()),"/","-")
fromTime=formatdate(reportTime,"dd.mm.yyyy")&" 07:15"
mailTitleFromDate=formatdate(reportTime,"dd.mm.yyyy")
toTime=formatdate(now(),"dd.mm.yyyy")&" 07:14"
'fromTime="28.07.2013 07:15"
'toTime="29.07.2013 07:15"

reportDay=longweekdayconvert(weekday(reportTime))&","&formatdate(reportTime,"dd.mm.yyyy")


'mailto="Ivan.li@knowles.com;Vivian.Huang@knowles.com;Yao.chao@knowles.com;Young.li@knowles.com;bo.wang@knowles.com"
'mailto="Ivan.li@knowles.com;bo.wang@knowles.com;chao.yao@knowles.com"
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
htTotal=htTotal&"<td colspan='2'><b>Process Yield</b></td><td colspan='2'><b>Acoustic Yield</b></td><td colspan='2'><b>Cosmetic Yield</b></td>"
htTotal=htTotal&"<td colspan='2'><b>FPY</b></td><td colspan='2'><b>Blocked[kpcs]</b></td></tr>"
htTotal=htTotal&"<tr align='center' bgcolor='#B9E9FF'><td ><b>Target</b></td><td ><b>Real</b></td><td ><b>Target</b></td><td ><b>Real</b></td><td ><b>Target</b></td>"
htTotal=htTotal&"<td ><b>Real</b></td><td ><b>Target</b></td><td ><b>Real</b></td><td ><b>+/-24h</b></td><td ><b>total</b></td></tr>"


'detail

'function detailhead (ProdType)
'htDetl="<table width='100%' border='1' cellpadding='0' cellspacing='0' ><tr align='center' bgcolor='#B9E9FF'><td colspan='17'><b>Details Beijing("&ProdType&"):</td></tr>"
'htDetl=htDetl&"<tr align='center' bgcolor='#B9E9FF'><td rowspan='2'><b>Product</b></td><td rowspan='2'><b>Line</b></td><td rowspan='2'><b>Target<br>[pcs]</b></td>"
'htDetl=htDetl&"<td rowspan='2'><b>Actual<br>[pcs]</b></td><td rowspan='2'><b>Real to<br>W/H</b></td>"
'htDetl=htDetl&"<td colspan='2'><b>Process Yield</b></td><td colspan='2'><b>Acoustic Yield</b></td><td colspan='2'><b>Cosmetic Yield</b></td>"
'htDetl=htDetl&"<td colspan='2'><b>FPY</b></td><td colspan='2'><b>Blocked[kpcs]</b></td></tr>"
'htDetl=htDetl&"<tr align='center' bgcolor='#B9E9FF'><td ><b>Target</b></td><td ><b>Real</b></td><td ><b>Target</b></td><td ><b>Real</b></td><td ><b>Target</b></td>"
'htDetl=htDetl&"<td ><b>Real</b></td><td ><b>Target</b></td><td ><b>Real</b></td><td ><b>+/-24h</b></td><td ><b>total</b></td></tr>"
'detailhead=htDetl
'end function
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
OEMlines=0
OEMtotalTarget=0
OEMtotalActual=0
OEMtargetYield=0
OEMtotalTarget=0
OEMtotalProcTargetQty=0
OEMrealYield=0
OEMtotalProcQty=0
OEMtotalProcJobQty=0



OEMtotalAcousticTargetQty=0
OEMtotalTarget=0


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

'------------------------------------------------------Details Beijing(Semi):Total----------------------------------
DetailsAutolines= DetailsA("PETRA,RA","Details Beijing(Autolines):")
DetailsHighFiex=  DetailsA("Winslet,Garland","Details Beijing(HighFiex):")
DetailsSemi=  DetailsA("Hektor","Details Beijing(Semi):")
DetailsOEM= DetailsB("Julia-HAC,Nith,16mm MFD (KEW),16mm MFD (KEB)","OEM")
	
function DetailsA(ProductName,ProdType)
htDetl="<table width='100%' border='1' cellpadding='0' cellspacing='0' ><tr align='center' bgcolor='#B9E9FF'><td colspan='17'><b>"&ProdType&"</b></td></tr>"
htDetl=htDetl&"<tr align='center' bgcolor='#B9E9FF'><td rowspan='2'><b>Product</b></td><td rowspan='2'><b>Line</b></td><td rowspan='2'><b>Target<br>[pcs]</b></td>"
htDetl=htDetl&"<td rowspan='2'><b>Actual<br>[pcs]</b></td><td rowspan='2'><b>Real to<br>W/H</b></td>"
htDetl=htDetl&"<td colspan='2'><b>Process Yield</b></td><td colspan='2'><b>Acoustic Yield</b></td><td colspan='2'><b>Cosmetic Yield</b></td>"
htDetl=htDetl&"<td colspan='2'><b>FPY</b></td><td colspan='2'><b>Blocked[kpcs]</b></td></tr>"
htDetl=htDetl&"<tr align='center' bgcolor='#B9E9FF'><td ><b>Target</b></td><td ><b>Real</b></td><td ><b>Target</b></td><td ><b>Real</b></td><td ><b>Target</b></td>"
htDetl=htDetl&"<td ><b>Real</b></td><td ><b>Target</b></td><td ><b>Real</b></td><td ><b>+/-24h</b></td><td ><b>total</b></td></tr>"
set rsC=server.createobject("adodb.recordset")




sqlC="SELECT count(*) Num,PRODUCT  FROM VW_BPS_DR_ALL where PRODUCT  in('"&replace(ProductName,",","','")&"') and TARGET<>0 group by PRODUCT order by product"

rsC.open sqlC,conn,1,3
   do while not rsC.eof
   
   product = rsC("product")
   ProductCount = cint(rsC("Num"))
   
   subTotalTarget=0
subTotalAct=0

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
subTHoldQty=0
subHoldQty=0

sql="SELECT PRODUCT,LINE_NUM,CREATE_TIME,LINE_TYPE,TARGET,nvl(ACTUAL,'0') ACTUAL,nvl(REAL_TO_WH,'0') REAL_TO_WH ,nvl(PQ_OUT,0) PQ_OUT, nvl(PQ_START,0) PQ_START,PQ_TARGET, PQ_YIELD,TARGET*(PQ_TARGET/100) AS PROCESS_TARGET,AMS_GOOD,AMS_START,AMS_TARGET,AMS_YIELD,TARGET*(AMS_TARGET/100) AS ACOUSTIC_TARGET,COSMETIC_GOOD,COSMETIC_START,COSMETIC_TARGET,COSMETIC_YIELD,TARGET*(COSMETIC_TARGET/100) AS FOI_TARGET,TTY_OPEN,TTY_CLOSE,TTY_TARGET,TTY_YIELD,TARGET*(TTY_TARGET/100) AS FPY_QTY FROM VW_BPS_DR_ALL  where  PRODUCT  ='"&product&"' and TARGET<>0  ORDER BY PRODUCT,LINE_NUM  "
rs.open sql,conn,1,3


do while not rs.eof


   '  if i=0 then
		tempProduct = rs("PRODUCT")
	
	
	
	
'----------------------------------------------------------------------------------------------------------------------------		
	totalTarget=totalTarget+csng(rs("TARGET"))
	subTotalTarget=subTotalTarget+csng(rs("TARGET"))
	htDetl=htDetl&"<tr align='center'><td>"&product&"</td><td>"&"LINE"&"&nbsp;#"&right(rs("LINE_Num"),1)&"</td><td>"&Formatnumber(rs("TARGET"),0)&"</td><td>"
	'get actual qty
	
		totalActual=totalActual+clng(rs("actual"))
		subTotalAct=subTotalAct+clng(rs("actual"))
		if clng(rs("actual")) < clng(rs("TARGET")) then
			fontColor="red"
		else
			fontColor="green"
		end if
		htDetl=htDetl&"<font color='"&fontColor&"'>"&Formatnumber(rs("actual"),0)&"</font>"
	'end if
	htDetl=htDetl&"&nbsp;</td><td>"
	'rsReal.close
'------------------------------------------------------------------------------------------	
	'get real to W/H qty
	
	
		totalWH=totalWH+clng(rs("REAL_TO_WH"))			
		subWH=subWH+clng(rs("REAL_TO_WH"))
		htDetl=htDetl&Formatnumber(rs("REAL_TO_WH"),0)
	'end if
	htDetl=htDetl&"&nbsp;</td>"
	'rsReal.close
'------------------------------------------------------------------------------------------	
	'get process_yield 
	htDetl=htDetl&"<td>"&Formatnumber(rs("PQ_TARGET"))&"%</td><td>"
	totalProcTargetQty=totalProcTargetQty+clng(rs("PROCESS_TARGET"))
	subProcTargetQty=subProcTargetQty+clng(rs("PROCESS_TARGET")	)  'PQ_YIELD,TARGET*(PQ_TARGET/100)
	
		totalProcQty=totalProcQty+csng(rs("PQ_OUT"))
		subProcQty=subProcQty+csng(rs("PQ_OUT"))
		totalProcJobQty=totalProcJobQty+csng(rs("PQ_START"))
		subProcJobQty=subProcJobQty+clng(rs("PQ_START"))
		if isnull(rs("PQ_YIELD") )then
		yield=0
		else
		yield=csng(rs("PQ_YIELD"))/100
		end if
		
		if yield < clng(rs("PQ_TARGET"))/100 then
			fontColor="red"
		else
			fontColor="green"
		end if
		htDetl=htDetl&"<font color='"&fontColor&"'>"&Formatnumber(yield*100,2)&"%</font>"
	
	htDetl=htDetl&"&nbsp;</td>"
	
'------------------------------------------------------------------------------------------
	'get Acoustic_Yield 
	htDetl=htDetl&"<td>"&Formatnumber(rs("AMS_TARGET"),2)&"%</td><td>"
	totalAcousticTargetQty=totalAcousticTargetQty+csng(rs("ACOUSTIC_TARGET"))
	subAcousticTargetQty=subAcousticTargetQty+csng(rs("ACOUSTIC_TARGET"))
	    
		if  isnull(rs("AMS_GOOD")) then
	    AMS_GOOD=0
		else
		AMS_GOOD=csng(rs("AMS_GOOD"))
		end if
	
	    if isnull(rs("AMS_START")) then
		AMS_START=0
		else
		AMS_START=csng(rs("AMS_START"))
		end if
	
		totalAcousticQty=totalAcousticQty+AMS_GOOD
		subAcousticQty=subAcousticQty+AMS_GOOD
		totalAcousticJobQty=totalAcousticJobQty+AMS_START
		subAcousticJobQty=subAcousticJobQty+AMS_START
		
		if isnull(rs("AMS_YIELD")) then
		yield=0
		else
		yield=csng(rs("AMS_YIELD"))/100
		end if
		if yield < csng(rs("AMS_TARGET"))/100 then
			fontColor="red"
		else

			fontColor="green"
		end if
		htDetl=htDetl&"<font color='"&fontColor&"'>"&Formatnumber(yield*100,2)&"%</font>"

	htDetl=htDetl&"&nbsp;</td>"
	
'------------------------------------------------------------------------------------------
	'get Foi_Yield 
	htDetl=htDetl&"<td>"&Formatnumber(rs("COSMETIC_TARGET"),2)&"%</td><td>"
	totalFoiTargetQty=totalFoiTargetQty+csng(rs("FOI_TARGET"))
	subFoiTargetQty=subFoiTargetQty+csng(rs("FOI_TARGET"))
	
	    if isnull(rs("COSMETIC_GOOD")) then
		COSMETIC_GOOD=0
		else
		COSMETIC_GOOD=csng(rs("COSMETIC_GOOD"))
		end if
		
		if  isnull(rs("COSMETIC_START")) then
		COSMETIC_START=0
		else
		COSMETIC_START=csng(rs("COSMETIC_START"))
		end if
		
		totalFoiQty=totalFoiQty+COSMETIC_GOOD
		subFoiQty=subFoiQty+COSMETIC_GOOD
		totalFoiJobQty=totalFoiJobQty+COSMETIC_START
		subFoiJobQty=subFoiJobQty+COSMETIC_START
		
		if  isnull(rs("COSMETIC_YIELD")) then
		yield=0
		else
		yield=csng(rs("COSMETIC_YIELD"))/100
		end if
		
	
		if yield < csng(rs("COSMETIC_TARGET"))/100 then
			fontColor="red"
		else
			fontColor="green"
		end if
		htDetl=htDetl&"<font color='"&fontColor&"'>"&Formatnumber(yield*100,2)&"%</font>"
	
	htDetl=htDetl&"&nbsp;</td>"
			
	

	
'------------------------------------------------------------------------------------------
	'get fpy 
	htDetl=htDetl&"<td>"&Formatnumber(rs("TTY_TARGET"))&"%</td><td>"
	totalFpyTargetQty=totalFpyTargetQty+csng(rs("FPY_QTY"))
	subFpyTargetQty=subFpyTargetQty+csng(rs("FPY_QTY"))
		
		if  isnull(rs("TTY_CLOSE")) then
		TTY_CLOSE=0
		else
		TTY_CLOSE=csng(rs("TTY_CLOSE"))
		end if
		if  isnull(rs("TTY_OPEN")) then
		TTY_OPEN=0
		else
		TTY_OPEN=csng(rs("TTY_OPEN"))
		end if
		
		
		totalFpyQty=totalFpyQty+TTY_CLOSE
		subFpyQty=subFpyQty+TTY_CLOSE
		uniqFpyQty=TTY_CLOSE
		totalFpyJobQty=totalFpyJobQty+TTY_OPEN
		subFpyJobQty=subFpyJobQty+TTY_OPEN
		uniqFpyJobQty=TTY_OPEN
		
		if  isnull(rs("TTY_YIELD")) then
		yield=0
		else
		yield=csng(rs("TTY_YIELD"))/100
		end if
		
		
		if yield < csng(rs("TTY_TARGET"))/100 then
			fontColor="red"
		else
			fontColor="green"
		end if
		htDetl=htDetl&"<font color='"&fontColor&"'>"&Formatnumber(yield*100,2)&"%</font>"
	
	'	uniqFpyQty=""
		'uniqFpyJobQty=""
	
	htDetl=htDetl&"&nbsp;</td><td>&nbsp;"
	

    htDetl=htDetl&"</td><td>&nbsp;</td></tr>"
	
	lines=lines+1
rs.movenext
Loop
rs.close

if ProductCount >1 then  
		htDetl=htDetl&"<tr align='center'><td><b>"&tempProduct&"</b></td><td><b>Total</b></td><td><b>"&Formatnumber(subTotalTarget,0)&"</br></td><td>"
'----------------------------------------------------------------------------------------------------------------------------		
		
		if subTotalAct>0 then
			if subTotalAct < subTotalTarget then
				fontColor="red"
			else
				fontColor="green"
			end if
			htDetl=htDetl&"<font color='"&fontColor&"' ><b>"&Formatnumber(subTotalAct,0)&"</b></font>"	
		else	
		htDetl=htDetl&"-"	
		end if
		htDetl=htDetl&"&nbsp;</td><td>"
		
'----------------------------------------------------------------------------------------------------------------------------		
		
		if subWH>0 then
			htDetl=htDetl&"<B>"&Formatnumber(subWH,0)&"</b>"
		else
		htDetl=htDetl&"-"	
		end if
		htDetl=htDetl&"&nbsp;</td>"
		
'----------------------------------------------------------------------------------------------------------------------------		
		subProcTargetYield=subProcTargetQty/subTotalTarget        '  TARGET*(PQ_TARGET/100) 叫加权平均算法  所有线总和/TARGET总和
		htDetl=htDetl&"<td><b>"&Formatnumber(subProcTargetYield*100,2)&"%</b></td><td>"   
		if subProcJobQty > 0 then
			subProcYield=subProcQty/subProcJobQty	'所有的良品总和/所有线投入总和
			if subProcYield < subProcTargetYield then
				fontColor="red"
			else
				fontColor="green"
			end if
			htDetl=htDetl&"<font color='"&fontColor&"' ><b>"&Formatnumber(subProcYield*100,2)&"%</b></font>"
		else
		htDetl=htDetl&"-"	
		end if		
		htDetl=htDetl&"&nbsp;</td>"
		
'----------------------------------------------------------------------------------------------------------------------------	
		subAcousticTargetYield=subAcousticTargetQty/subTotalTarget
		htDetl=htDetl&"<td><b>"&Formatnumber(subAcousticTargetYield*100,2)&"%</b></td><td>"
		if subAcousticJobQty>0 then
			subAcousticYield=subAcousticQty/subAcousticJobQty				
			if subAcousticYield < subAcousticTargetYield then
				fontColor="red"
			else
				fontColor="green"
			end if
			htDetl=htDetl&"<font color='"&fontColor&"' ><b>"&Formatnumber(subAcousticYield*100,2)&"%</b></font>"
		else
		htDetl=htDetl&"-"
		end if
		htDetl=htDetl&"&nbsp;</td>"
'----------------------------------------------------------------------------------------------------------------------------					
		subFoiTargetYield=subFoiTargetQty/subTotalTarget
		htDetl=htDetl&"<td><b>"&Formatnumber(subFoiTargetYield*100,2)&"%</b></td><td>"
		if subFoiJobQty>0 then
			subFoiYield=subFoiQty/subFoiJobQty				
			if subFoiYield < subFoiTargetYield then
				fontColor="red"
			else
				fontColor="green"
			end if
			htDetl=htDetl&"<font color='"&fontColor&"' ><b>"&Formatnumber(subFoiYield*100,2)&"%</b></font>"
		else
		htDetl=htDetl&"-"	
		end if
		htDetl=htDetl&"&nbsp;</td>"
'----------------------------------------------------------------------------------------------------------------------------			
	
		subFpyTargetYield=subFpyTargetQty/subTotalTarget
		htDetl=htDetl&"<td><b>"&Formatnumber(subFpyTargetYield*100,2)&"%</b></td><td>"
		if subFpyJobQty >0 then
			subFpyYield=subFpyQty/subFpyJobQty				
			if subFpyYield < subFpyTargetYield then
				fontColor="red"
			else
				fontColor="green"
			end if
			htDetl=htDetl&"<font color='"&fontColor&"' ><b>"&Formatnumber(subFpyYield*100,2)&"%</b></font>"
        else
		htDetl=htDetl&"-"	
		end if
		htDetl=htDetl&"&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td></tr>"	
		
'----------------------------------------------------------------------------------------------------------------------------			
		'if rs.eof then
			'exit for
		'end if
		
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

else
htDetl=htDetl&"<tr ><td colspan='15' height='10'></td></tr>"
end if

rsC.movenext

Loop

rsC.close
htDetl=htDetl&"</table><table width='100%' border='0' cellpadding='0' cellspacing='0'><tr><td>&nbsp;</td></tr></table>"

DetailsA=htDetl
end function

	
function DetailsB(ProductName,ProdType)
htDetl="<table width='100%' border='1' cellpadding='0' cellspacing='0' ><tr align='center' bgcolor='#B9E9FF'><td colspan='17'><b>"&ProdType&"</b></td></tr>"
htDetl=htDetl&"<tr align='center' bgcolor='#B9E9FF'><td rowspan='2'><b>Product</b></td><td rowspan='2'><b>Line</b></td><td rowspan='2'><b>Target<br>[pcs]</b></td>"
htDetl=htDetl&"<td rowspan='2'><b>Actual<br>[pcs]</b></td><td rowspan='2'><b>Real to<br>W/H</b></td>"
htDetl=htDetl&"<td colspan='2'><b>Process Yield</b></td><td colspan='2'><b>Acoustic Yield</b></td><td colspan='2'><b>Cosmetic Yield</b></td>"
htDetl=htDetl&"<td colspan='2'><b>FPY</b></td><td colspan='2'><b>Blocked[kpcs]</b></td></tr>"
htDetl=htDetl&"<tr align='center' bgcolor='#B9E9FF'><td ><b>Target</b></td><td ><b>Real</b></td><td ><b>Target</b></td><td ><b>Real</b></td><td ><b>Target</b></td>"
htDetl=htDetl&"<td ><b>Real</b></td><td ><b>Target</b></td><td ><b>Real</b></td><td ><b>+/-24h</b></td><td ><b>total</b></td></tr>"
set rsC=server.createobject("adodb.recordset")




sqlC="SELECT count(*) Num,PRODUCT  FROM VW_BPS_DR_ALL where PRODUCT  in('"&replace(ProductName,",","','")&"') and TARGET<>0 group by PRODUCT order by product"
rsC.open sqlC,conn,1,3
   do while not rsC.eof
   
   product = rsC("product")
   ProductCount = cint(rsC("Num"))
   
subTotalTarget=0
subTotalAct=0
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
subTHoldQty=0
subHoldQty=0

sql="SELECT PRODUCT,LINE_NUM,CREATE_TIME,LINE_TYPE,TARGET,nvl(ACTUAL,'0') ACTUAL,nvl(REAL_TO_WH,'0') REAL_TO_WH ,nvl(PQ_OUT,0) PQ_OUT, nvl(PQ_START,0) PQ_START,PQ_TARGET, PQ_YIELD,TARGET*(PQ_TARGET/100) AS PROCESS_TARGET,AMS_GOOD,AMS_START,AMS_TARGET,AMS_YIELD,TARGET*(AMS_TARGET/100) AS ACOUSTIC_TARGET,COSMETIC_GOOD,COSMETIC_START,COSMETIC_TARGET,COSMETIC_YIELD,TARGET*(COSMETIC_TARGET/100) AS FOI_TARGET,TTY_OPEN,TTY_CLOSE,TTY_TARGET,TTY_YIELD,TARGET*(TTY_TARGET/100) AS FPY_QTY FROM VW_BPS_DR_ALL  where  PRODUCT  ='"&product&"' and TARGET<>0  ORDER BY PRODUCT,LINE_NUM  "
rs.open sql,conn,1,3


do while not rs.eof


   '  if i=0 then
		tempProduct = rs("PRODUCT")
	
	
	
	
'----------------------------------------------------------------------------------------------------------------------------		
	OEMtotalTarget=OEMtotalTarget+csng(rs("TARGET"))
	subTotalTarget=subTotalTarget+csng(rs("TARGET"))
	htDetl=htDetl&"<tr align='center'><td>"&product&"</td><td>"&"LINE"&"&nbsp;#"&right(rs("LINE_Num"),1)&"</td><td>"&Formatnumber(rs("TARGET"),0)&"</td><td>"
	'get actual qty
	
		OEMtotalActual=OEMtotalActual+clng(rs("actual"))
		subTotalAct=subTotalAct+clng(rs("actual"))
		if clng(rs("actual")) < clng(rs("TARGET")) then
			fontColor="red"
		else
			fontColor="green"
		end if
		htDetl=htDetl&"<font color='"&fontColor&"'>"&Formatnumber(rs("actual"),0)&"</font>"
	'end if
	htDetl=htDetl&"&nbsp;</td><td>"
	'rsReal.close
'------------------------------------------------------------------------------------------	
	'get real to W/H qty
	
	
		OEMtotalWH=OEMtotalWH+clng(rs("REAL_TO_WH"))			
		subWH=subWH+clng(rs("REAL_TO_WH"))
		htDetl=htDetl&Formatnumber(rs("REAL_TO_WH"),0)
	'end if
	htDetl=htDetl&"&nbsp;</td>"
	'rsReal.close
'------------------------------------------------------------------------------------------	
	'get process_yield 
	htDetl=htDetl&"<td>"&cstr(rs("PQ_TARGET"))&"%</td><td>"
	OEMtotalProcTargetQty=OEMtotalProcTargetQty+clng(rs("PROCESS_TARGET"))
	subProcTargetQty=subProcTargetQty+clng(rs("PROCESS_TARGET")	)  'PQ_YIELD,TARGET*(PQ_TARGET/100)
	
		OEMtotalProcQty=OEMtotalProcQty+csng(rs("PQ_OUT"))
		subProcQty=subProcQty+csng(rs("PQ_OUT"))
		OEMtotalProcJobQty=OEMtotalProcJobQty+csng(rs("PQ_START"))
		subProcJobQty=subProcJobQty+clng(rs("PQ_START"))
		if isnull(rs("PQ_YIELD") )then
		yield=0
		else
		yield=csng(rs("PQ_YIELD"))/100
		end if
		
		if yield < clng(rs("PQ_TARGET"))/100 then
			fontColor="red"
		else
			fontColor="green"
		end if
		htDetl=htDetl&"<font color='"&fontColor&"'>"&Formatnumber(yield*100,2)&"%</font>"
	
	htDetl=htDetl&"&nbsp;</td>"
	
'------------------------------------------------------------------------------------------
	'get Acoustic_Yield 
	htDetl=htDetl&"<td>"&csng(rs("AMS_TARGET"))&"%</td><td>"
	OEMtotalAcousticTargetQty=OEMtotalAcousticTargetQty+csng(rs("ACOUSTIC_TARGET"))
	subAcousticTargetQty=subAcousticTargetQty+csng(rs("ACOUSTIC_TARGET"))
	    
		if  isnull(rs("AMS_GOOD")) then
	    AMS_GOOD=0
		else
		AMS_GOOD=csng(rs("AMS_GOOD"))
		end if
	
	    if isnull(rs("AMS_START")) then
		AMS_START=0
		else
		AMS_START=csng(rs("AMS_START"))
		end if
	
		OEMtotalAcousticQty=OEMtotalAcousticQty+AMS_GOOD
		subAcousticQty=subAcousticQty+AMS_GOOD
		OEMtotalAcousticJobQty=OEMtotalAcousticJobQty+AMS_START
		subAcousticJobQty=subAcousticJobQty+AMS_START
		
		if isnull(rs("AMS_YIELD")) then
		yield=0
		else
		yield=csng(rs("AMS_YIELD"))/100
		end if
		if yield < csng(rs("AMS_TARGET"))/100 then
			fontColor="red"
		else

			fontColor="green"
		end if
		htDetl=htDetl&"<font color='"&fontColor&"'>"&Formatnumber(yield*100,2)&"%</font>"

	htDetl=htDetl&"&nbsp;</td>"
	
'------------------------------------------------------------------------------------------
	'get Foi_Yield 
	htDetl=htDetl&"<td>"&csng(rs("COSMETIC_TARGET"))&"%</td><td>"
	OEMtotalFoiTargetQty=OEMtotalFoiTargetQty+csng(rs("FOI_TARGET"))
	subFoiTargetQty=subFoiTargetQty+csng(rs("FOI_TARGET"))
	
	    if isnull(rs("COSMETIC_GOOD")) then
		COSMETIC_GOOD=0
		else
		COSMETIC_GOOD=csng(rs("COSMETIC_GOOD"))
		end if
		
		if  isnull(rs("COSMETIC_START")) then
		COSMETIC_START=0
		else
		COSMETIC_START=csng(rs("COSMETIC_START"))
		end if
		
		OEMtotalFoiQty=OEMtotalFoiQty+COSMETIC_GOOD
		subFoiQty=subFoiQty+COSMETIC_GOOD
		OEMtotalFoiJobQty=OEMtotalFoiJobQty+COSMETIC_START
		subFoiJobQty=subFoiJobQty+COSMETIC_START
		
		if  isnull(rs("COSMETIC_YIELD")) then
		yield=0
		else
		yield=csng(rs("COSMETIC_YIELD"))/100
		end if
		if yield < csng(rs("COSMETIC_TARGET"))/100 then
			fontColor="red"
		else
			fontColor="green"
		end if
		htDetl=htDetl&"<font color='"&fontColor&"'>"&Formatnumber(yield*100,2)&"%</font>"
	
	htDetl=htDetl&"&nbsp;</td>"
	
'------------------------------------------------------------------------------------------
	'get fpy 
	htDetl=htDetl&"<td>"&csng(rs("TTY_TARGET"))&"%</td><td>"
	OEMtotalFpyTargetQty=OEMtotalFpyTargetQty+csng(rs("FPY_QTY"))
	subFpyTargetQty=subFpyTargetQty+csng(rs("FPY_QTY"))
		
		if  isnull(rs("TTY_CLOSE")) then
		TTY_CLOSE=0
		else
		TTY_CLOSE=csng(rs("TTY_CLOSE"))
		end if
		if  isnull(rs("TTY_OPEN")) then
		TTY_OPEN=0
		else
		TTY_OPEN=csng(rs("TTY_OPEN"))
		end if
		OEMtotalFpyQty=OEMtotalFpyQty+TTY_CLOSE
		subFpyQty=subFpyQty+TTY_CLOSE
		uniqFpyQty=TTY_CLOSE
		OEMtotalFpyJobQty=OEMtotalFpyJobQty+TTY_OPEN
		subFpyJobQty=subFpyJobQty+TTY_OPEN
		uniqFpyJobQty=TTY_OPEN
		if  isnull(rs("TTY_YIELD")) then
		yield=0
		else
		yield=csng(rs("TTY_YIELD"))/100
		end if
		
		
		if yield < csng(rs("TTY_TARGET"))/100 then
			fontColor="red"
		else
			fontColor="green"
		end if
		htDetl=htDetl&"<font color='"&fontColor&"'>"&Formatnumber(yield*100,2)&"%</font>"
	
	'	uniqFpyQty=""
		'uniqFpyJobQty=""
	
	htDetl=htDetl&"&nbsp;</td><td>&nbsp;"
	

    htDetl=htDetl&"</td><td>&nbsp;</td></tr>"
	
	OEMlines=OEMlines+1
	
rs.movenext
Loop
rs.close

if ProductCount >1 then  
		htDetl=htDetl&"<tr align='center'><td><b>"&tempProduct&"</b></td><td><b>Total</b></td><td><b>"&subTotalTarget&"</br></td><td>"
'----------------------------------------------------------------------------------------------------------------------------		
		
		if subTotalAct>0 then
			if subTotalAct < subTotalTarget then
				fontColor="red"
			else
				fontColor="green"
			end if
			htDetl=htDetl&"<font color='"&fontColor&"' ><b>"&subTotalAct&"</b></font>"	
		else	
		htDetl=htDetl&"-"	
		end if
		htDetl=htDetl&"&nbsp;</td><td>"
		
'----------------------------------------------------------------------------------------------------------------------------		
		
		if subWH>0 then
			htDetl=htDetl&"<B>"&subWH&"</b>"
		else
		htDetl=htDetl&"-"	
		end if
		htDetl=htDetl&"&nbsp;</td>"
		
'----------------------------------------------------------------------------------------------------------------------------		
		subProcTargetYield=subProcTargetQty/subTotalTarget        '  TARGET*(PQ_TARGET/100) 叫加权平均算法  所有线总和/TARGET总和
		htDetl=htDetl&"<td><b>"&Formatnumber(subProcTargetYield*100,2)&"%</b></td><td>"   
		if subProcJobQty>0 then		
			subProcYield=subProcQty/subProcJobQty				 ''所有的良品总和/所有线投入总和
			if subProcYield < subProcTargetYield then
				fontColor="red"
			else
				fontColor="green"
			end if
			htDetl=htDetl&"<font color='"&fontColor&"' ><b>"&Formatnumber(subProcYield*100,2)&"%</b></font>"
		else
		htDetl=htDetl&"-"	
		end if		
		htDetl=htDetl&"&nbsp;</td>"
		
'----------------------------------------------------------------------------------------------------------------------------	
		subAcousticTargetYield=subAcousticTargetQty/subTotalTarget
		htDetl=htDetl&"<td><b>"&Formatnumber(subAcousticTargetYield*100,2)&"%</b></td><td>"
		if subAcousticJobQty>0 then
			subAcousticYield=subAcousticQty/subAcousticJobQty				
			if subAcousticYield < subAcousticTargetYield then
				fontColor="red"
			else
				fontColor="green"
			end if
			htDetl=htDetl&"<font color='"&fontColor&"' ><b>"&Formatnumber(subAcousticYield*100,2)&"%</b></font>"
		else
		htDetl=htDetl&"-"
		end if
		htDetl=htDetl&"&nbsp;</td>"
'----------------------------------------------------------------------------------------------------------------------------					
		subFoiTargetYield=subFoiTargetQty/subTotalTarget
		htDetl=htDetl&"<td><b>"&Formatnumber(subFoiTargetYield*100,2)&"%</b></td><td>"
		if subFoiJobQty>0 then
			subFoiYield=subFoiQty/subFoiJobQty				
			if subFoiYield < subFoiTargetYield then
				fontColor="red"
			else
				fontColor="green"
			end if
			htDetl=htDetl&"<font color='"&fontColor&"' ><b>"&Formatnumber(subFoiYield*100,2)&"%</b></font>"
		else
		htDetl=htDetl&"-"	
		end if
		htDetl=htDetl&"&nbsp;</td>"
'----------------------------------------------------------------------------------------------------------------------------			
	
		subFpyTargetYield=subFpyTargetQty/subTotalTarget
		htDetl=htDetl&"<td><b>"&Formatnumber(subFpyTargetYield*100,2)&"%</b></td><td>"
		if subFpyJobQty >0 then
			subFpyYield=subFpyQty/subFpyJobQty				
			if subFpyYield < subFpyTargetYield then
				fontColor="red"
			else
				fontColor="green"
			end if
			htDetl=htDetl&"<font color='"&fontColor&"' ><b>"&Formatnumber(subFpyYield*100,2)&"%</b></font>"
        else
		htDetl=htDetl&"-"	
		end if
		htDetl=htDetl&"&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td></tr>"	
		
'----------------------------------------------------------------------------------------------------------------------------			
		'if rs.eof then
			'exit for
		'end if
		
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

else
htDetl=htDetl&"<tr ><td colspan='15' height='10'></td></tr>"
end if

rsC.movenext

Loop

rsC.close
htDetl=htDetl&"</table><table width='100%' border='0' cellpadding='0' cellspacing='0'><tr><td>&nbsp;</td></tr></table>"

DetailsB=htDetl
end function
'------------------------------总和----------------------------------------------------

subHTTotal="<td>"&lines&"</td><td>"&Formatnumber(totalTarget,0)&"</td><td>"
if totalActual>0 then
	if totalActual < totalTarget then
		fontColor="red"
	else
		fontColor="green"
	end if
	subHTTotal=subHTTotal&"<font color='"&fontColor&"' >"&Formatnumber(totalActual,0)&"</font>"
end if
subHTTotal=subHTTotal&"&nbsp;</td><td>"
if totalWH>0 then
	subHTTotal=subHTTotal&Formatnumber(totalWH,0)
end if
subHTTotal=subHTTotal&"&nbsp;</td>"
targetYield=totalProcTargetQty/totalTarget
subHTTotal=subHTTotal&"<td>"&Formatnumber(targetYield*100,2)&"%</td><td>"
if totalProcJobQty >0 then
	realYield=totalProcQty/totalProcJobQty				
	if realYield < targetYield then
		fontColor="red"
	else
		fontColor="green"
	end if
	subHTTotal=subHTTotal&"<font color='"&fontColor&"' >"&Formatnumber(realYield*100,2)&"%</font>"	
end if
subHTTotal=subHTTotal&"&nbsp;</td>"

targetYield=totalAcousticTargetQty/totalTarget
subHTTotal=subHTTotal&"<td>"&Formatnumber(targetYield*100,2)&"%</td><td>"
if totalAcousticJobQty>0 then
	realYield=totalAcousticQty/totalAcousticJobQty		
	if realYield < targetYield then
		fontColor="red"
	else
		fontColor="green"
	end if
	subHTTotal=subHTTotal&"<font color='"&fontColor&"' >"&Formatnumber(realYield*100,2)&"%</font>"	
end if
subHTTotal=subHTTotal&"&nbsp;</td>"

targetYield=totalFoiTargetQty/totalTarget
subHTTotal=subHTTotal&"<td>"&Formatnumber(targetYield*100,2)&"%</td><td>"
if totalFoiJobQty >0 then
	realYield=totalFoiQty/totalFoiJobQty	
	if realYield < targetYield then
		fontColor="red"
	else
		fontColor="green"
	end if
	subHTTotal=subHTTotal&"<font color='"&fontColor&"' >"&Formatnumber(realYield*100,2)&"%</font>"	
end if
subHTTotal=subHTTotal&"&nbsp;</td>"







targetYield=totalFpyTargetQty/totalTarget
subHTTotal=subHTTotal&"<td>"&Formatnumber(targetYield*100,2)&"%</td><td>"
if totalFpyJobQty>0 then
	realYield=totalFpyQty/totalFpyJobQty				
	if realYield < targetYield then
		fontColor="red"
	else
		fontColor="green"
	end if
	subHTTotal=subHTTotal&"<font color='"&fontColor&"' >"&Formatnumber(realYield*100,2)&"%</font>"	
end if
subHTTotal=subHTTotal&"&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td></tr>"

'----------------------------------OEM Total-----------------------------

OEMsubHTTotal="<td>"&OEMlines&"</td><td>"&Formatnumber(OEMtotalTarget,0)&"</td><td>"
if OEMtotalActual>0 then
	if OEMtotalActual < OEMtotalTarget then
		fontColor="red"
	else
		fontColor="green"
	end if
	OEMsubHTTotal=OEMsubHTTotal&"<font color='"&fontColor&"' >"&Formatnumber(OEMtotalActual,0)&"</font>"
end if
OEMsubHTTotal=OEMsubHTTotal&"&nbsp;</td><td>"
if totalWH>0 then
	OEMsubHTTotal=OEMsubHTTotal&OEMtotalWH
end if
OEMsubHTTotal=OEMsubHTTotal&"&nbsp;</td>"


OEMtargetYield=OEMtotalProcTargetQty/OEMtotalTarget
OEMsubHTTotal=OEMsubHTTotal&"<td>"&Formatnumber(OEMtargetYield*100,2)&"%</td><td>"
if OEMtotalProcJobQty >0 then
	OEMrealYield=OEMtotalProcQty/OEMtotalProcJobQty				
	if OEMrealYield < OEMtargetYield then
		fontColor="red"
	else
		fontColor="green"
	end if
	OEMsubHTTotal=OEMsubHTTotal&"<font color='"&fontColor&"' >"&Formatnumber(OEMrealYield*100,2)&"%</font>"	
end if
OEMsubHTTotal=OEMsubHTTotal&"&nbsp;</td>"





OMEtargetYield=OEMtotalAcousticTargetQty/OEMtotalTarget
OEMsubHTTotal=OEMsubHTTotal&"<td>"&Formatnumber(OMEtargetYield*100,2)&"%</td><td>"
if totalAcousticJobQty>0 then
	realYield=OMEtotalAcousticQty/totalAcousticJobQty		
	if realYield < targetYield then
		fontColor="red"
	else
		fontColor="green"
	end if
	OEMsubHTTotal=OEMsubHTTotal&"<font color='"&fontColor&"' >"&Formatnumber(realYield*100,2)&"%</font>"	
end if
OEMsubHTTotal=OEMsubHTTotal&"&nbsp;</td>"

targetYield=totalFoiTargetQty/totalTarget
OEMsubHTTotal=OEMsubHTTotal&"<td>"&Formatnumber(targetYield*100,2)&"%</td><td>"
if totalFoiJobQty >0 then
	realYield=totalFoiQty/totalFoiJobQty	
	if realYield < targetYield then
		fontColor="red"
	else
		fontColor="green"
	end if
	OEMsubHTTotal=OEMsubHTTotal&"<font color='"&fontColor&"' >"&Formatnumber(realYield*100,2)&"%</font>"	
end if
OEMsubHTTotal=OEMsubHTTotal&"&nbsp;</td>"







targetYield=totalFpyTargetQty/totalTarget
OEMsubHTTotal=OEMsubHTTotal&"<td>"&Formatnumber(targetYield*100,2)&"%</td><td>"
if totalFpyJobQty>0 then
	realYield=totalFpyQty/totalFpyJobQty				
	if realYield < targetYield then
		fontColor="red"
	else
		fontColor="green"
	end if
	OEMsubHTTotal=OEMsubHTTotal&"<font color='"&fontColor&"' >"&Formatnumber(realYield*100,2)&"%</font>"	
end if
OEMsubHTTotal=OEMsubHTTotal&"&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td></tr>"
'--------------------------------------总和Total-----------------'

subHTTotalNew="<td><B>"&lines+OEMlines&"</b></td><td><b>"&Formatnumber(totalTarget+OEMtotalTarget,0)&"</b></td><td>"
if totalActual>0 then
	if totalActual < totalTarget then
		fontColor="red"
	else
		fontColor="green"
	end if
	subHTTotalNew=subHTTotalNew&"<font color='"&fontColor&"' ><B>"&Formatnumber(totalActual+OEMtotalActual,0)&"</b></font>"
end if
subHTTotalNew=subHTTotalNew&"&nbsp;</td><td>"
if totalWH>0 then
	subHTTotalNew=subHTTotalNew&"<B>"&Formatnumber(totalWH,0)&"</B>"
end if
subHTTotalNew=subHTTotalNew&"&nbsp;</td>"
targetYield=totalProcTargetQty/totalTarget
subHTTotalNew=subHTTotalNew&"<td><B>"&Formatnumber(targetYield*100,2)&"%</b></td><td>"
if totalProcJobQty >0 then
	realYield=totalProcQty/totalProcJobQty				
	if realYield < targetYield then
		fontColor="red"
	else
		fontColor="green"
	end if
	subHTTotalNew=subHTTotalNew&"<font color='"&fontColor&"' >"&Formatnumber(realYield*100,2)&"%</font>"	
end if
subHTTotalNew=subHTTotalNew&"&nbsp;</td>"

targetYield=totalAcousticTargetQty/totalTarget
subHTTotalNew=subHTTotalNew&"<td><B>"&Formatnumber(targetYield*100,2)&"%</b></td><td>"
if totalAcousticJobQty>0 then
	realYield=totalAcousticQty/totalAcousticJobQty		
	if realYield < targetYield then
		fontColor="red"
	else
		fontColor="green"
	end if
	subHTTotalNew=subHTTotalNew&"<font color='"&fontColor&"' ><B>"&Formatnumber(realYield*100,2)&"%</b></font>"	
end if
subHTTotalNew=subHTTotalNew&"&nbsp;</td>"

targetYield=totalFoiTargetQty/totalTarget
subHTTotalNew=subHTTotalNew&"<td><b>"&Formatnumber(targetYield*100,2)&"%</b></td><td>"
if totalFoiJobQty >0 then
	realYield=totalFoiQty/totalFoiJobQty	
	if realYield < targetYield then
		fontColor="red"
	else
		fontColor="green"
	end if
	subHTTotalNew=subHTTotalNew&"<font color='"&fontColor&"' ><b>"&Formatnumber(realYield*100,2)&"%</b></font>"	
end if
subHTTotalNew=subHTTotalNew&"&nbsp;</td>"





targetYield=totalFpyTargetQty/totalTarget
subHTTotalNew=subHTTotalNew&"<td><b>"&Formatnumber(targetYield*100,2)&"%</b></td><td>"
if totalFpyJobQty>0 then
	realYield=totalFpyQty/totalFpyJobQty				
	if realYield < targetYield then
		fontColor="red"
	else
		fontColor="green"
	end if
	subHTTotalNew=subHTTotalNew&"<font color='"&fontColor&"' ><b>"&Formatnumber(realYield*100,2)&"%</b></font>"	
end if

subHTTotalNew=subHTTotalNew&"&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td></tr>"



'--------------------------------------------------------'






'htDetlSemi =detailhead("Semi")
' htDetlHighFiex=detailhead("HighFiex")
 'DetailsSemi=  DetailsA("Hektor,Maple")
 'DetailsHighFiex=  DetailsA("Winslet,Marigold")
	




htTotal=htTotal&"<tr align='center'><td>Beijing</td>"&subHTTotal&"<tr align='center'><td><b>OEM</b></td>"&OEMsubHTTotal&"<tr align='center'><td><b>Total</b></td>"&subHTTotalNew&"</table><br>"

mailContents=htmls&htTotal&DetailsHighFiex&DetailsSemi&DetailsAutolines&DetailsOEM&htmlDesc&"</body></html>"
response.Write(mailContents)

'SendJMail application("MailSender"),mailTo," Test Report  "&mailTitleFromDate,mailContents
%>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->