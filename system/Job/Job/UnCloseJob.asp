<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<!--#include virtual="/Functions/RanPwd.asp" -->
<%
 
	if (request.querystring("Action")="1") then
		jobnumber=request.querystring("jobnumber")
		SQL="select job_number,part_number_tag,Line_name,input_time,start_quantity,"
		SQL=SQL+" final_good_quantity,final_input_quantity,confirm_good_quantity, final_scrap_quantity "
		SQL=SQL+" from job_master"
		SQL=SQL+" where store_status='0' and input_time>=to_date('2010-12-20','yyyy-mm-dd')"
		if(jobnumber<>"") then
			SQL=SQL+" AND JOB_NUMBER='"+jobnumber+"'"
		end if
		SQL=SQL+" AND JOB_NUMBER in (select distinct JOB_NUMBER from UNCLOSEJOB WHERE USER_CODE='"+session("code")+"')"
		SQL=SQL+" and factory_id='FA00000002' and line_name not in ('SUB1','SUB2','SECOP','ANNEAL','OSP','SUB3','SUB4','SPRAY')"
		session("sql")=sql
		rs.open SQL,conn,1,3
	end if 
	
	if (request.querystring("Action")="2") then	
			Set Upload = Server.CreateObject("Persits.Upload")
			Upload.LogonUser "", "ASPEncrypt", "Dickens;'2006"
			Upload.Save
			path=Upload.Form("path")
			query=Upload.Form("query")
			query=replace(query,"*","&")
			beforepath=path&"?"&query
			Set File = Upload.Files("file1")
			filePath=server.mappath("TEMP")
			filename=gen_key(10)+session("code")+".xlsx"
			
			File.SaveAs filePath &"\"+filename
			set Upload=nothing
			
			 strAddr=filePath &"\"+filename
	 
			 
		 
			 set excelconn=server.createobject("adodb.connection") 
			 excelconn.open "Provider = Microsoft.ACE.OLEDB.12.0; Data Source = "+strAddr+";Extended Properties='Excel 8.0;HDR=NO;IMEX=1'" 
			 set rsExcel=server.createobject("adodb.recordset")
			 SQL="select * from [Sheet1$]" 
			 rsExcel.open SQL,excelconn,1,3
			 set rsOracle=server.createobject("adodb.recordset")
			 rsOracle.open "delete from UNCLOSEJOB where USER_CODE='"+session("code")+"'",conn,1,3
			 for i=0 to rsExcel.recordcount-1
				 sql1="select * from UNCLOSEJOB where 1=2"
				 set rsOracle2=server.createobject("adodb.recordset")
				 rsOracle2.open sql1,conn,1,3
				 rsOracle2.addnew
				   rsOracle2("JOB_NUMBER")=rsExcel(0)
				   rsOracle2("LINE")=rsExcel(1)
				   rsOracle2("REMAIN_QTY")=rsExcel(2)
				   rsOracle2("MODEL_NAME")=rsExcel(3)
				   rsOracle2("CLASS")=rsExcel(4)
				   rsOracle2("QTY")=rsExcel(5)
				   rsOracle2("USER_CODE")=session("code")
			   rsOracle2.update
			   rsOracle2.close
			   rsExcel.movenext
			 next 
			 response.write "<script>window.alert('导入成功!')</script>"
		  
	end if

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System</title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<link href="/CSS/dynCalendar.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<script>
	function SubmitForm()
	{
		var jobnumber;
		jobnumber=form1.jobnumber.value;
		form1.action="UnCloseJob.asp?Action=1&jobnumber="+jobnumber;
		form1.submit();	 
	}
	
	function ImportExcel()
	{
		form1.action="UnCloseJob.asp?Action=2";
		form1.submit();
	}
</script>
</head>

<body>
<form name="form1" method="post" enctype="multipart/form-data">
<table width="99%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="8" class="t-c-greenCopy"><span id="inner_Search"></span></td>
  </tr>
  <tr>
    <td height="20" width="10%">Job Number</td>
    <td height="20" width="10%"><input name="jobnumber" type="text" id="jobnumber" value="<%=jobnumber%>"></td>
    
    <td width="10%"><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="SubmitForm()">
	<Td width="40%">
	<input name="file1" type="file" id="file1" size="100"></Td>
	<td colspan="4">
	<input name="importexcel" type="button" id="importexcel"  value="Import Excel" onClick="ImportExcel()">
	
	</td>
    </tr>
</table>
</form>
<table width="99%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="18" class="t-c-greenCopy"><span id="inner_Browse"></span></td>
</tr>
<TR>
	<TD align="right" colspan="18">结果导出excel:<span style="cursor:hand" title="Export Data to EXCEL File" onClick="javascript:window.open('UncloseJobExcel.asp')"><img src="/Images/EXCEL.gif" width="20" height="20"></span></TD>
</TR>
<TR>
	<TD class="t-t-Borrow">Job_Number</TD>
	<TD class="t-t-Borrow">Model Name</TD>
	<TD class="t-t-Borrow">Line Name</TD>
	<TD class="t-t-Borrow">Input Time</TD>
	<TD class="t-t-Borrow">Input Quantity</TD>
	<TD class="t-t-Borrow">Good Qty(Old)</TD>
	<TD class="t-t-Borrow">Good Qty(New System)</TD>
	<TD class="t-t-Borrow">Good Qty(New Confirm)</TD>
	<TD class="t-t-Borrow">Scrap</TD>
	<TD class="t-t-Borrow">WIP Status</TD>
</TR>

<%
	 if(rs.State > 0 ) then	
	 	for i=0 to rs.recordcount-1
%>
<tr>
<td><%=rs(0)%></td>
<td><%=rs(1)%></td>
<td><%=rs(2)%></td>
<td><%=rs(3)%></td>
<td><%=rs(4)%></td>
<td><%=rs(5)%></td>
<td><%=rs(6)%></td>
<td><%=rs(7)%></td>
<td><%=rs(8)%></td>
<td><%=GetWIPStatus(rs(0))%>&nbsp;</td>
<td>
</td>
</tr>
<%
		rs.movenext
		next 
	 end if 
%>


<%
	function GetWIPStatus(Job_Number)
		'1.主线为完成
		SQL="SELECT * FROM JOB WHERE JOB_NUMBER='"+Job_Number+"' AND CLOSE_TIME IS NULL"
		set RsMainLineWIP=server.CreateObject("adodb.recordset")
		RsMainLineWIP.OPEN sql,conn,1,3
		if (RsMainLineWIP.recordcount>0) then
			StatusStr="主线未完成|"
		end if 

		'2.OFFLINE未开工单
		SQL="SELECT * FROM OFFLINE_STORE WHERE JOB_NUMBER='"+Job_Number+"' AND TRANSACTIONTYPE IN ('1','4') AND ISPRINT='0'"
		set RsOfflineWIP=server.CreateObject("adodb.recordset")
		RsOfflineWIP.OPEN sql,conn,1,3
		if (RsOfflineWIP.recordcount>0) then
			StatusStr=StatusStr+"OFFLINE未打印|"
		end if 

		'3. OFFLINE已开,主线未投
		SQL="SELECT * FROM OFFLINE_STORE WHERE JOB_NUMBER='"+Job_Number+"' AND TRANSACTIONTYPE IN ('1','4') AND ISPRINT='1'"
		set RsOfflinePrint=server.CreateObject("adodb.recordset")
		RsOfflinePrint.open sql,conn,1,3
		if (RsOfflinePrint.recordcount>0) then
			SQL="SELECT * FROM JOB WHERE JOB_NUMBER='"+Job_Number+"' AND SHEET_NUMBER>=500"
			set RSMAIN2=server.CreateObject("adodb.recordset")
			RSMAIN2.OPEN SQL,conn,1,3
			if RSMAIN2.recordcount=0 then
				StatusStr=StatusStr+"OFFLINE已打印,主线未开始|"
			end if 
		end if
		
		
		'4.主线产生的rework没有完全入offline
		SQL="SELECT decode(sum(defect_quantity),null,0,sum(defect_quantity))FROM job_defectcodes jd, defectcode d where jd.defect_code_id=d.nid and jd.job_number='"+Job_Number+"' and d.transaction_type='1'"
		set RsTotalReworkQty=server.CreateObject("adodb.recordset")
		RsTotalReworkQty.open SQL,conn,1,3
		if (RsTotalReworkQty.recordcount>0) then
			TotalReworkQty=cdbl(RsTotalReworkQty(0))
		end if 
		SQL="SELECT decode(sum(qty),null,0,sum(qty)) from offline_store where job_number='"+Job_Number+"' and transactiontype='1'"
		set RsOfflineRecieveQty=server.CreateObject("adodb.recordset")
		RsOfflineRecieveQty.open SQL,conn,1,3
		if (RsOfflineRecieveQty.recordcount>0) then
			TotalOfflineRecieveQty=cdbl(RsOfflineRecieveQty(0))
		end if 
		
		SQL="SELECT decode(SUM(UNITQTY),null,0,SUM(UNITQTY)) FROM ENG_BORROW WHERE JOBNUMBER='"&Job_Number&"'  AND TRANSACTIONTYPE='1'"
		set rsEngBorrow=server.CreateObject("adodb.recordset")
		rsEngBorrow.open SQL,conn,1,3
		if (rsEngBorrow.recordcount>0) then
			EngBorrow=cdbl(rsEngBorrow(0))
		end if 
		
		if(TotalReworkQty>TotalOfflineRecieveQty+EngBorrow) then
			StatusStr=StatusStr+"Rework未入offline|"
		end if 

		'5.主线产生的change model没有完全入offline
		SQL="SELECT decode(sum(defect_quantity),null,0,sum(defect_quantity)) FROM job_defectcodes jd, defectcode d where jd.defect_code_id=d.nid and jd.job_number='"+Job_Number+"' and d.transaction_type='4'"
		set RsTotalChangeModelQty=server.CreateObject("adodb.recordset")
		RsTotalChangeModelQty.open SQL,conn,1,3
		if (RsTotalChangeModelQty.recordcount>0) then
			TotalChangeModelQty=cdbl(RsTotalChangeModelQty(0))
		end if 
		
		SQL="SELECT decode(sum(qty),null,0,sum(qty)) from offline_store where job_number='"+Job_Number+"' and transactiontype='4'"
		set RsOfflineRecieveQty2=server.CreateObject("adodb.recordset")
		RsOfflineRecieveQty2.open SQL,conn,1,3
		if (RsOfflineRecieveQty2.recordcount>0) then
			TotalOfflineRecieveQty2=cdbl(RsOfflineRecieveQty2(0))
		end if 
		
		SQL="SELECT decode(SUM(UNITQTY),null,0,SUM(UNITQTY)) FROM ENG_BORROW WHERE JOBNUMBER='"&Job_Number&"'  AND TRANSACTIONTYPE='4'"
		set rsEngBorrow2=server.CreateObject("adodb.recordset")
		rsEngBorrow2.open SQL,conn,1,3
		if (rsEngBorrow2.recordcount>0) then
			EngBorrow2=cdbl(rsEngBorrow2(0))
		end if 
		
		if(TotalChangeModelQty>TotalOfflineRecieveQty2+EngBorrow2) then
			StatusStr=StatusStr+"Change Model未入offline|"
		end if 
		
		'6 主线产生的slapping没有完全进入retest
		SQL=" select  DECODE(SUM(A.ACTION_VALUE),NULL,0,SUM(A.ACTION_VALUE)) AS QTY from job_actions a,action b "
		SQL=SQL+" where a.action_id=b.nid and job_number='"&Job_Number&"'and b.mother_action_id='AN00000260'"
		set RsSlappingQty=server.CreateObject("adodb.recordset")
		RsSlappingQty.open SQL,conn,1,3
		if (RsSlappingQty.recordcount>0) then
			TotalSlappingQty=cdbl(RsSlappingQty(0))
		end if 
		
		SQL=" select DECODE(sum(startqty),NULL,0,sum(startqty)) from job_retest where batchno like '"+Job_Number+"-%-5'"
		set RsSlappingRecieveQty=server.CreateObject("adodb.recordset")
		RsSlappingRecieveQty.open SQL,conn,1,3
		if (RsSlappingRecieveQty.recordcount>0) then
			TotalReceiveSlappingQty=cdbl(RsSlappingRecieveQty(0))
		end if 
		
		if(TotalSlappingQty>TotalReceiveSlappingQty) then
			StatusStr=StatusStr+"Slapping未入Retest|"
		end if 
		
		'7 主线已完成,未进入retest
		SQL="SELECT  DECODE(SUM(J.JOB_GOOD_QUANTITY),null,0,SUM(J.JOB_GOOD_QUANTITY)) AS QTY FROM job J WHERE J.JOB_NUMBER='"&Job_Number&"'"
		'response.write sql &"<br>"
		set rsGoodTotalQty=server.createobject("adodb.recordset")
		rsGoodTotalQty.open SQL,conn,1,3
		if rsGoodTotalQty.recordcount>0 then
			GoodTotalQty=cdbl(rsGoodTotalQty(0))
		end if 
		 
		'SQL=" select DECODE(sum(actualqty),NULL,0,sum(actualqty)) from job_retest a , label_print_history b  where a.batchno=b.batchno and a.batchno like '"+Job_Number+"-%--1'"
		SQL=" select DECODE(sum(actualqty),NULL,0,sum(actualqty)) from job_retest a , label_print_history b  where a.batchno=b.batchno "
		SQL=SQL+" and (a.batchno like '"+Job_Number+"-%--1' OR a.batchno like '"+Job_Number+"-%-5')"
		'response.write sql &"<br>"
		set RsGoodRecieveQty=server.CreateObject("adodb.recordset")
		RsGoodRecieveQty.open SQL,conn,1,3
		if (RsGoodRecieveQty.recordcount>0) then
			TotalReceiveGoodQty=cdbl(RsGoodRecieveQty(0))
		end if 
		
		if(GoodTotalQty>TotalReceiveGoodQty) then
			StatusStr=StatusStr+"Good未完全入Retest|"
		else
			SQL="select decode(sum(store_quantity),null,0,sum(store_quantity)) from job_master_store_pre where job_number='"+Job_Number+"'"
			set RsGoodStoreQty=server.CreateObject("adodb.recordset")
			RsGoodStoreQty.open SQL,conn,1,3	
			GoodStoreQty=0
			if RsGoodStoreQty.recordcount>0 then
				GoodStoreQty=cint(RsGoodStoreQty(0))
			end if 	
			if(GoodTotalQty>GoodStoreQty) then
				StatusStr=StatusStr+"已Retest未入库|"
			end if 
		end if 
		
		GetWIPStatus=StatusStr
	end function 
%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Commit.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" -->