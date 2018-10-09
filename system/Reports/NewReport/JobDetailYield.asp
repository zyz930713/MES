<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<%
	job_number=request.querystring("Job_Number")
	sql="SELECT job_number,sheet_number,part_number_tag,start_time,close_time,job_start_quantity,job_good_quantity,job_defectcode_quantity,job_type,line_name FROM job  "
	sql=sql+" WHERE JOB_NUMBER='"+job_number+"' order by job_number,sheet_number "
	rs.open SQL,conn,1,3
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<link href="/CSS/dynCalendar.css" rel="stylesheet" type="text/css">
<script language=javascript src="/Components/sniffer.js" type=text/javascript></script>
<script language=javascript src="/Components/dynCalendar.js" type=text/javascript></script>
</head>
<body>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="13"  align="center"> <font size="5">Job Detail Infomation</font> </td>
  </tr>
  <tr>
  	<td class="t-t-Borrow">&nbsp;</td>
	<td class="t-t-Borrow">&nbsp;</td>
	<td class="t-t-Borrow">&nbsp;</td>
	<td class="t-t-Borrow">&nbsp;</td>
	<td class="t-t-Borrow">&nbsp;</td>
	<td class="t-t-Borrow">&nbsp;</td>
	<td colspan="2" class="t-t-Borrow" align="center">Total Good Qty</td>
	<td colspan="3" class="t-t-Borrow" align="center">Total Scrap</td>
	<td class="t-t-Borrow">&nbsp;</td>
	<td class="t-t-Borrow">&nbsp;</td>
  </tr>
  <tr>
    <td height="20" class="t-t-Borrow"><div align="center">Job Number</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">Sheet Number</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">Model</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">Start Time</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">Close Time</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">Start Qty</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">Good Qty</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">Slapping Good Qty</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">Direct Scrap</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">Change Model</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">Rework</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">Job Type</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">Line Name</div></td>
  </tr>
	<%  if(rs.State > 0 ) then	
			for j=0 to rs.recordcount-1
	%>
	<tr>
		<td height="20" <%if cint(rs("sheet_number").value)>=500 then%>class="t-t-Borrow"<%end if %>><div align="center"><%=rs("job_number").value%>&nbsp;</div></td>
		<td height="20" <%if cint(rs("sheet_number").value)>=500 then%>class="t-t-Borrow"<%end if %>><div align="center"><%=rs("sheet_number").value%>&nbsp;</div></td>
		<td height="20" <%if cint(rs("sheet_number").value)>=500 then%>class="t-t-Borrow"<%end if %>><div align="center"><%=rs("part_number_tag").value%>&nbsp;</div></td>
		<td height="20" <%if cint(rs("sheet_number").value)>=500 then%>class="t-t-Borrow"<%end if %>><div align="center"><%=rs("start_time").value%>&nbsp;</div></td>
		<td height="20" <%if cint(rs("sheet_number").value)>=500 then%>class="t-t-Borrow"<%end if %>><div align="center"><%=rs("close_time").value%>&nbsp;</div></td>
		<td height="20" <%if cint(rs("sheet_number").value)>=500 then%>class="t-t-Borrow"<%end if %>><div align="center"><%=rs("job_start_quantity").value%>&nbsp;</div></td>
		
		<%
			'StartQty
			if(cint(rs("sheet_number").value)<500) then
				StartQty=StartQty+cint(rs("job_start_quantity").value)
			end if
			
			
			'Slapping Qty
			SQL3=" select  DECODE(SUM(A.ACTION_VALUE),NULL,0,SUM(A.ACTION_VALUE)) AS QTY from job_actions a,action b "
			SQL3=SQL3+" where a.action_id=b.nid and job_number='"&rs("job_number").value&"' and A.SHEET_NUMBER  = '"&rs("sheet_number").value&"'  and b.mother_action_id='AN00000260'"	 
			set rs3=server.createobject("adodb.recordset")
			rs3.open SQL3,conn,1,3	 
			if rs3.recordcount<>0 then
				SlappingQty=cint(rs3("QTY").value)
			else
				SlappingQty=0
			end if
		%> 
		<td height="20" <%if cint(rs("sheet_number").value)>=500 then%>class="t-t-Borrow"<%end if %>><div align="center"><%=cstr(cint(rs("job_good_quantity").value)-SlappingQty)%></div></td>
		<%
			'Total Good
				TotalGoodQty=TotalGoodQty+cint(rs("job_good_quantity").value)-SlappingQty
			
		%>
		<td height="20" <%if cint(rs("sheet_number").value)>=500 then%>class="t-t-Borrow"<%end if %>><div align="center"><%=SlappingQty%></div></td>
		
		<%'Rework , Change model, Scrap, Readjust Qty
			SQL1="SELECT SUM(JD.DEFECT_QUANTITY) AS QTY, DC.TRANSACTION_TYPE "
			SQL1=SQL1+" FROM JOB_DEFECTCODES JD , DEFECTCODE DC"
			SQL1=SQL1+ " WHERE JD.DEFECT_CODE_ID=DC.NID(+)"
			SQL1=SQL1+" AND JD.JOB_NUMBER='"&rs("job_number").value&"' AND JD.SHEET_NUMBER in ("&rs("sheet_number").value&") AND DC.TRANSACTION_TYPE IN ('1', '2','4')"
			SQL1=SQL1+" GROUP BY DC.TRANSACTION_TYPE"
			SQL1=SQL1+" ORDER BY DC.TRANSACTION_TYPE"
			set rs1=server.createobject("adodb.recordset")
			rs1.open SQL1,conn,1,3
			Scrap=0
			ChangeModel=0
			Rework=0
			while not rs1.eof 
				if(rs1("TRANSACTION_TYPE").value="1") then
					Rework=rs1("QTY").value
				end if
				if(rs1("TRANSACTION_TYPE").value="2") then
					Scrap=rs1("QTY").value
					TotalScrapQty=TotalScrapQty+cint(Scrap)
				end if
				if(rs1("TRANSACTION_TYPE").value="4") then
					ChangeModel=rs1("QTY").value
					TotalChangeModelQty=TotalChangeModelQty+cint(ChangeModel)
				end if
				
				rs1.movenext
			wend
	    %>
	
	
		<td height="20" <%if cint(rs("sheet_number").value)>=500 then%>class="t-t-Borrow"<%end if %>><div align="center"><%=Scrap%></div></td>
		<td height="20" <%if cint(rs("sheet_number").value)>=500 then%>class="t-t-Borrow"<%end if %>><div align="center"><%=ChangeModel%></div></td>
		<td height="20" <%if cint(rs("sheet_number").value)>=500 then%>class="t-t-Borrow"<%end if %>><div align="center"><%=Rework%></div></td>
		<td height="20" <%if cint(rs("sheet_number").value)>=500 then%>class="t-t-Borrow"<%end if %>><div align="center"><%=rs("job_type").value%></div></td>
		<td height="20" <%if cint(rs("sheet_number").value)>=500 then%>class="t-t-Borrow"<%end if %>><div align="center"><%=rs("line_name").value%></div></td>
  
	<%
			rs.movenext
			next 
	%>
	</tr>
	<%
		end if
	%>
  </tr>
  <%
  		'Get Retest reject Qty
'		SQL1="select  decode(SUM(a.SAMPLEQTY),null,0,SUM(a.SAMPLEQTY)),decode(SUM(A.REJECTQTY),null,0,SUM(A.REJECTQTY)) from Job_ReTest_detail   a, "
'		SQL1=SQL1+"(select batchno,max(testsequence) as testsequence from Job_ReTest_detail "
'		SQL1=SQL1+"group by batchno)b,"
'		SQL1=SQL1+" Job_ReTest C "
'		SQL1=SQL1+" where a.batchno=b.batchno and a.testsequence=b.testsequence"
'		SQL1=SQL1+" AND A.BATCHNO=C.BATCHNO AND C.isreleasetooqc='1' AND c.BATCHNO LIKE '"+job_number+"%'"
'		
'		set rsRetest=server.createobject("adodb.recordset")
'		rsRetest.open SQL1,conn,1,3
'		
'		if rsRetest.recordcount>0 then
'			RetestTotalIn=rsRetest(0)
'			RetestTotalReject=rsRetest(1)
'			
'			TotalScrapQty=TotalScrapQty+cint(RetestTotalReject)
'			
'			TotalGoodQty=TotalGoodQty-cint(RetestTotalReject)
'		end if
  %>
  <!--
   <tr>
    <td height="20"><div align="center">Retest</div></td>
	<td height="20"><div align="center">&nbsp;</div></td>
	<td height="20"><div align="center">&nbsp;</div></td>
	<td height="20"><div align="center">&nbsp;</div></td>
	<td height="20"><div align="center">&nbsp;</div></td>
	<td height="20"><div align="center">&nbsp;<%=RetestTotalIn%></div></td>
	<td height="20"><div align="center">&nbsp;</div></td>
	<td height="20"><div align="center">&nbsp;</div></td>
	<td height="20"><div align="center">&nbsp;<%=RetestTotalReject%></div></td>
	<td height="20"><div align="center">&nbsp;</div></td>
	<td height="20"><div align="center">&nbsp;</div></td>
	<td height="20"><div align="center">&nbsp;</div></td>
	<td height="20"><div align="center">&nbsp;</div></td>
  </tr>
  -->
  <%
  		'Get Retest reject  Qty
		'SQL1="SELECT sum(oqcreject) FROM JOB_OQC_DETAIL a  where a.BATCHNO LIKE '"+job_number+"%' AND a.ISPASSOQC=1"
		SQL1="SELECT decode(sum(oqcreject)+sum(productionretestrejectqty),null,0,sum(oqcreject)+sum(productionretestrejectqty)) FROM JOB_OQC_DETAIL a "
		SQL1=SQL1+" where a.BATCHNO LIKE '"+job_number+"%'"
		
		set rsOQCReject=server.createobject("adodb.recordset")
		rsOQCReject.open SQL1,conn,1,3
		
		if rsOQCReject.recordcount>0 then
			OQCTotalReject=rsOQCReject(0)
			TotalScrapQty=TotalScrapQty+cint(OQCTotalReject)
			TotalGoodQty=TotalGoodQty-cint(OQCTotalReject)
		end if
		
		SQL1="SELECT sum(SAMPLEQTY) FROM JOB_OQC a  "
		SQL1=SQL1+" where OQCNID IN (SELECT OQCNID FROM JOB_OQC_DETAIL  where BATCHNO LIKE '"+job_number+"%' AND ISPASSOQC=1)"
		set rsOQCTotal=server.createobject("adodb.recordset")
		rsOQCTotal.open SQL1,conn,1,3
		
		if rsOQCTotal.recordcount>0 then
			OQCTotalIn=rsOQCTotal(0)
		end if
  %>
   <tr>
    <td height="20"><div align="center">QA</div></td>
	<td height="20"><div align="center">&nbsp;</div></td>
	<td height="20"><div align="center">&nbsp;</div></td>
	<td height="20"><div align="center">&nbsp;</div></td>
	<td height="20"><div align="center">&nbsp;</div></td>
	<td height="20"><div align="center">&nbsp;<%=OQCTotalIn%></div></td>
	<td height="20"><div align="center">&nbsp;</div></td>
	<td height="20"><div align="center">&nbsp;</div></td>
	<td height="20"><div align="center">&nbsp;<%=OQCTotalReject%></div></td>
	<td height="20"><div align="center">&nbsp;</div></td>
	<td height="20"><div align="center">&nbsp;</div></td>
	<td height="20"><div align="center">&nbsp;</div></td>
	<td height="20"><div align="center">&nbsp;</div></td>
  </tr>
  
  
  
   <tr>
    <td height="20" class="t-t-Borrow"><div align="center">Total</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">&nbsp;</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">&nbsp;</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">&nbsp;</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">&nbsp;</div></td>
	<td height="20" class="t-t-Borrow"><div align="center"><%=StartQty%>&nbsp;</div></td>
	<td height="20" class="t-t-Borrow"><div align="center"><%=TotalGoodQty%>&nbsp;</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">&nbsp;</div></td>
	<td height="20" class="t-t-Borrow"><div align="center"><%=TotalScrapQty%>&nbsp;</div></td>
	<td height="20" class="t-t-Borrow"><div align="center"><%=TotalChangeModelQty%>&nbsp;</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">Final Yield</div></td>
	<td height="20" class="t-t-Borrow"><div align="center"><%=round(TotalGoodQty/StartQty*100,2)%>%</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">&nbsp;</div></td>
  </tr>
  <tr>
  	<Td colspan="13">
		<Table width="50%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
			<Tr>
				<td colspan="6" align="center"><font size="5">Engineer Sample History</font></td>
			</tr>
			<tr>
				<td height="20" class="t-t-Borrow"><div align="center">User Code</div></td>
				<td height="20" class="t-t-Borrow"><div align="center">Time</div></td>
				<td height="20" class="t-t-Borrow"><div align="center">Job Number</div></td>
				<td height="20" class="t-t-Borrow"><div align="center">Sheet Number</div></td>
				<td height="20" class="t-t-Borrow"><div align="center">Qty</div></td>
				<td height="20" class="t-t-Borrow"><div align="center">Transation Type</div></td>
			</tr>
			<%
				SQL="SELECT * FROM ENG_BORROW WHERE JOBNUMBER='"+job_number+"' ORDER BY TRANSACTIONTYPE,SHEETNUMBER"
				set rs1=server.createobject("adodb.recordset")
				rs1.open SQL,conn,1,3
				while not rs1.eof
			%>
			<tr>
				<td height="20" ><div align="center"><%=rs1("usercode").value%></div></td>
				<td height="20"><div align="center"><%=rs1("borrowtime").value%></div></td>
				<td height="20"><div align="center"><%=rs1("jobnumber").value%></div></td>
				<td height="20"><div align="center"><%=rs1("sheetnumber").value%></div></td>
				<td height="20"><div align="center"><%=rs1("unitqty").value%></div></td>
				<td height="20"><div align="center">  <%if rs1("TransactionType")="-1" then response.write "Normal" end if%> 
																		  <%if rs1("TransactionType")="0"  then response.write "None" end if %>  
																		  <%if rs1("TransactionType")="1" then response.write "Rework" end if%>  
																		  <%if rs1("TransactionType")="2" then response.write "Scrap" end if%>  
																		  <%if rs1("TransactionType")="3" then response.write "Readjust" end if%> 
																		  <%if rs1("TransactionType")="4" then response.write "Others" end if%> 
																		  <%if rs1("TransactionType")="5" then response.write "Slapping" end if%> 
													</div>
				</td>
			</tr>
			<% 
				rs1.movenext
				wend
			%>
		</Table>
	</Td>
  </tr>
   
 
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->