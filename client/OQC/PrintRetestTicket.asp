<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>Print Retest Job Ticket 打印重测工单</title>
<style type="text/css">
<!--
.Barcode3of9 {	font-family: "3 of 9 Barcode";
	font-size: 30px;
}
.tdTitle{
	font-weight: bold; 
	font-size:14px
}
-->
</style>
</head>
<%
errMsg=""
batchNo=request("batch_no")
if batchNo <> "" then
	SQL="select job_number,seq,qty,model,station_name from rework_printing where batchno='"&batchNo&"' "
	rs.open SQL,conn,1,3
	if not rs.eof then
		jobNumer=rs("job_number")
		
		if len(rs("seq"))="1" then   '返检工单会直接生成到reword_printing 子工单只有一位。
		seq="00"&rs("seq")   ' 子工单需要三位,补00
		elseif len(rs("seq"))="2" then
		seq="0"&rs("seq")
		else
		seq=rs("seq")
		end if
		
		subJobNumer=jobNumer&"-"&seq
		qty=rs("qty")
		model=rs("model")
		stationId=rs("station_name")
		
	else
		errMsg="This batch no is not rejected by OQC.\n此批号没有被OQC拒绝。"
	end if
	rs.close
	
	if errMsg="" then
		SQL="select a.stations_index,b.line_name from PART a,job_master b where a.nid=b.part_number_id and b.job_number = '"&jobNumer&"' "
		rs.open SQL,conn,1,3
		if not rs.eof then
			stationIndex = rs("stations_index")
			lineName=rs("line_name")
		end if
		rs.close
		
		stationIndex=right(stationIndex,len(stationIndex)-instr(stationIndex,stationId)+1)
		SQL="select nid,station_name||'('||station_chinese_name||')' as station from station where nid in('"&replace(stationIndex,",","','")&"') "
		SQL=SQL+" order by instr('"&stationIndex&"',nid) "
		
		strStation=""
		rs.open SQL,conn,1,3
		i=1
		while not rs.eof
			strStation=strStation&"<tr><td >"&i&"</td><td nowrap>"&rs("station")&"</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td>"
			rs.movenext
			i=i+1
		wend
		rs.close
		
		'add print history
		SQL="select new_jobnumber,quantity,print_time from print_newjob_history where new_jobnumber='"&subJobNumer&"'"
		
		rs.open SQL,conn,1,3
		if rs.eof then
			rs.addnew
			rs("new_jobnumber")=subJobNumer
			rs("quantity")=qty
		end if
		rs("print_time")=now()
		rs.update
		rs.close
	end if
end if
if 	errMsg="" then
%>
<body >
<table width="700" border="0" cellpadding="0" cellspacing="0" align="center">
  <tr><td>
	<table width="100%" border="1" cellpadding="0" cellspacing="0">
	  <tr><td>
		<table width="100%" border="0" cellpadding="1" cellspacing="1" >
			<tr>
				<td class="tdTitle" colspan="2">DISCRETE JOB TRAVELER(Retest)</td>
				<td><span class="tdTitle">Qty:</span><%=qty%></td>
				<td><span class="tdTitle">Line Name:</span><%=lineName%></td>
			</tr>
			<tr>
				<td class="Barcode3of9" colspan="2"><%=subJobNumer%></td>
				<td><span class="tdTitle">P/N:</span><%=model%></td>
				<td colspan="2" ><%=now()%></td>
			</tr>
			<tr>
				<td><span class="tdTitle">Job Number:</span><%=subJobNumer%></td>
				<td class="Barcode3of9" colspan="2"><%=model%></td>			
			</tr>
		</table>
	  </td></tr>
	</table>
  </td></tr>
  <tr><td>
	<table border="1" cellpadding="0" cellspacing="0">
		<th>No</th>
		<th width="200">Operation</th>
		<th width="40">Op</th>
		<th>Good</th>	
		<th>Bad</th>
		<th width="50">Time</th>
	  <%=strStation%>
	</table>
  </td></tr>
  <tr><td>&nbsp;</td></tr>
  <tr><td>&nbsp;</td></tr>
  <tr><td>
	<div align="center"><span id="printspan" style="visibility:visible">	  
		  <input name="Print" type="button" id="Print" value="Print 打印" onClick="window.print();">
		  &nbsp;
		  <input name="Close" type="button" id="Close" onClick="javascript:window.close()" value="Close 关闭">
		  &nbsp;
	</span></div>
  </td></tr>
</table>
</body>
<script type="text/javascript">
function window.onafterprint()
{
	document.all.printspan.style.visibility="hidden";
	return true;
}
function window.onafterprint() 
{
	document.all.printspan.style.visibility="visible";
	return true;
}
</script>
<%else %>
<script type="text/javascript">
	alert("<%=errMsg%>");
	window.close();
</script>	
<%end if%>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->