 <%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
 IsPrintTravler=request("selIsPrintTravler")
 BatchNo=request("batchNo")
 
 if(IsPrintTravler)="" then
 	IsPrintTravler="0"
 end if 
 SQL="SELECT * FROM offline_store where transactiontype='4' "
 
 if (IsPrintTravler<>"-2") then
 	SQL=SQL+" AND ISPRINT='"+IsPrintTravler+"'"
 end if 
 if (BatchNo<>"") then
 	SQL=SQL+" AND BATCHNO='"+BatchNo+"'"
 end if 
 SQL=SQL+" ORDER BY BATCHNO"
 set rsChangeModel=server.createobject("adodb.recordset")
 rsChangeModel.open SQL,conn,1,3
 

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=application("SystemName")%></title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<link href="/CSS/dynCalendar.css" rel="stylesheet" type="text/css">
<script language=javascript src="/Components/sniffer.js" type=text/javascript></script>
<script language=javascript src="/Components/dynCalendar.js" type=text/javascript></script>
<script language="javascript" type="text/javascript">
	 function FindData()
	 {
	 	form1.action="ChangeModel.asp";
		form1.submit();
		
	 }
</script>
</head>

<body>
<form method="post" name="form1">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr class="t-b-midautumn">
	<td colspan="6">Change Model</td>
</tr>
<tr>
	<td width="15%">Batch No</td>
	<td width="15%"><input type="text" id="batchNo" name="batchNo" value="<%=batchNo%>" /></td>
	<td width="15%">IsPrintTravler</td>
	<td width="15%"><select id="selIsPrintTravler"  name="selIsPrintTravler" style="width:100px">
			<option value="-2" <%if IsPrintTravler="-2" then%> selected <% end if%>>--All--</option>
			<option value="0" <%if IsPrintTravler="0" then%> selected <% end if%>>Not Printed</option>
			<option value="1" <%if IsPrintTravler="1" then%> selected <% end if%> >Printed</option>
		</select></td>
	
	<td colspan="2"><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand"  onclick="FindData();"></td>
</tr>
<tr >
	<td  colspan="6">
		<table width="100%">
			<tr>
				<td height="20" class="t-c-greenCopy" width="3%">Action</td>
				<Td class="t-c-greenCopy" width="17%">BatchNo</Td>
				<Td class="t-c-greenCopy" width="10%">Old Model</Td>
				<Td class="t-c-greenCopy" width="10%">New Model</Td>
				<Td class="t-c-greenCopy" width="10%">Qty</Td>
				<Td class="t-c-greenCopy" width="10%">SubJobList</Td>
				<Td class="t-c-greenCopy" width="10%">ActualQty</Td>
				<Td class="t-c-greenCopy" width="10%">InOfflineTime</Td>
				<Td class="t-c-greenCopy" width="10%">IsPrintTravler</Td>
				<Td class="t-c-greenCopy" width="10%">Change History</Td>
			</tr>
			<%for i=1 to rsChangeModel.recordcount%>
			<tr>
				 <td width="38" height="20" class="t-c-GrayLight" ><div align="center"><span style="cursor:hand" onClick="javascript:window.open ('ChangeModel_Detail.asp?batchno=<%=rsChangeModel("BatchNo")%>&jobno=<%=rsChangeModel("job_number")%>','main')">
<img src="/Images/IconEdit.gif" alt="Click to Change Model"></span></div></td>

				<td class="t-c-GrayLight" height="20"><%=rsChangeModel("BatchNo")%></td>
				<%
					SQL="SELECT * FROM Change_Model_History WHERE BatchNo='"+rsChangeModel("BatchNo")+"' AND IsCurrentVersion='1'"
					set rsModelName=server.createobject("adodb.recordset")
 					rsModelName.open SQL,conn,1,3
					OldModel=""
					NewModel=""
					if(rsModelName.recordcount>0)then
						OldModel=rsModelName("PrevoiusModel")
						NewModel=rsModelName("CurrentModel")
					else
						SQL="SELECT * FROM JOB_MASTER WHERE JOB_NUMBER='"+rsChangeModel("JOB_NUMBER")+"'"
						set rsModelName2=server.createobject("adodb.recordset")
 						rsModelName2.open SQL,conn,1,3
						if rsModelName2.recordcount>0 then
							OldModel=rsModelName2("PART_NUMBER_TAG")
							NewModel=rsModelName2("PART_NUMBER_TAG")
						end if 
					end if
				%>
				<td class="t-c-GrayLight"><%=OldModel%></td>
				<td class="t-c-GrayLight"><%=NewModel%></td>
				<td class="t-c-GrayLight"><%=rsChangeModel("Qty")%></td>
				<td class="t-c-GrayLight"><%=rsChangeModel("SubJobList")%></td>
				<td class="t-c-GrayLight"><%=rsChangeModel("ActualQty")%></td>
				<td class="t-c-GrayLight"><%=rsChangeModel("INTIMESTAMP")%></td>
				<td class="t-c-GrayLight"><%=rsChangeModel("ISPRINT")%></td>
				 <td width="38" height="20" class="t-c-GrayLight" ><div align="center"><span style="cursor:hand" onClick="javascript:window.open ('ChangeHistory.asp?batchno=<%=rsChangeModel("BatchNo")%>&jobno=<%=rsChangeModel("job_number")%>','main')">
<img src="/Images/IconEdit.gif" alt="Click to Change Model"></span></div></td>
			</tr>
			<%
				rsChangeModel.movenext
			NEXT 
			%>
			
		</table>
	</td>
</tr>
	 
</table>
</form>

</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->