<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
oqcNID=request("oqc_nid")

strSql="select a.batchno,a.oqcreject,a.defectdesc,a.oqcendtime,b.job_number||'-'||b.seq as new_job,b.model,"
strSql=strSql+" (select station_name||'('||station_chinese_name||')' from station where nid=b.station_name) as station,"
strSql=strSql+" nvl((select 1 from print_newjob_history where new_jobnumber=b.job_number||'-'||b.seq),0) as is_print"
strSql=strSql+" from job_oqc_detail a,rework_printing b where a.batchno=b.batchno and a.oqcnid='"+oqcNID+"'"
strSql=strSql+" order by showsequence"

rs.open strSql,conn,1,3
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System</title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Language/language.js" type="text/javascript"></script>
</head>

<body onLoad="language(<%=session("language")%>);">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="9">&nbsp;</td>
  </tr>
<%if(rs.State > 0 ) then %>  
  <tr align="center">
  	<td height="20" class="t-t-Borrow"><span id="inner_NO"></span></td>
	<TD class="t-t-Borrow"><span id="td_BatchNo"></span></TD>
	<TD class="t-t-Borrow"><span id="td_Model"></span></TD>
	<TD class="t-t-Borrow"><span id="td_RejectQty"></span></TD>
	<TD class="t-t-Borrow"><span id="td_OQCEndTime"></span></TD>
	<TD class="t-t-Borrow"><span id="td_DefectInfo"></span></TD>
	<TD class="t-t-Borrow"><span id="td_NewJob"></span></TD>
	<TD class="t-t-Borrow"><span id="td_NewStation"></span></TD>
	<TD class="t-t-Borrow"><span id="td_IsPrinted"></span></TD>
  </tr>
	<%for j=0 to rs.recordcount-1%>
	<tr align="center">
		<td><%=(j+1)%></td>
		<td><%=rs("batchno")%></td>
		<td><%=rs("model")%></td>
		<td><%=rs("oqcreject")%></td>
		<td><%=rs("oqcendtime")%></td>
		<td><%=rs("defectdesc")%></td>
		<td><%=rs("new_job")%></td>
		<td><%=rs("station")%></td>
		<td><%=rs("is_print")%></td>
	</tr>	
	<%rs.movenext
	next
end if
rs.close	 
	%>	
</table>

</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->