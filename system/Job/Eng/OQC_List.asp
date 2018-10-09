<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
batchNo=request("batchNo")
if batchNo<>"" then
	where=where&" and JR.BATCHNO='"&batchNo&"' "
end if
SQL="select JR.*,JD.*,detail.* from JOB_RETEST JR left join JOB_RETEST_DETAIL JD on JR.BATCHNO=JD.BATCHNO "
SQL=SQL+" left join JOB_RETEST_DETAIL detail on detail.DETAILID=JD.DETAILID where ISLOWYIELDHOLD=1 and ISRELEASETOOQC=0 "&where
'yield=formatpercent(1-(cint(25)/cint(727)),2,-1)
'response.Write(yield)
rs.open SQL,conn,1,3
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=application("SystemName")%></title>
<link href="../../CSS/General.css" rel="stylesheet" type="text/css" />
<script language="javascript">
function FindData()
{
	document.form1.action="OQC_List.asp"
	document.form1.submit();
}
</script>
</head>

<body>
<form method="post" name="form1">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr class="t-b-midautumn"  height="20">
	<td colspan="6">Search Retest Infomation</td>
</tr>
<tr  height="20">
	<td width="15%">Batch Number</td>
	<td width="15%"><input type="text" id="batchN" name="batchNo" value="<%=batchNo%>" /></td>
	<td colspan="4"><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand"  onclick="FindData();"></td>
</tr>
</table>
</form>
<table  width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr class="t-c-greenCopy"  height="20">
	<td colspan="10">Browse List</td>
</tr>
<tr class="t-c-GrayLight" height="20">
	<td align="center">No</td>
	<td align="center">BatchNo</td>
	<td align="center">Retest Start Time</td>
	<td align="center">Start Qty</td>
	<td align="center">Sample Qty</td>
	<td align="center">Model</td>
	<td align="center">Good Qty</td>
	<td align="center">Reject Qty</td>
	<td align="center">Line loss Qty</td>
	<td align="center">Retest Yield</td>
</tr>
<%
i=1
if not rs.eof then
while not rs.eof 
%>
<tr  height="20">
	<td align="center"><%=i%></td>
	<td align="center"><a href="Engineer_OQC.asp?batchNo=<%=rs("BatchNo")%>" target="_blank"> <%=rs("BatchNo")%></a></td>
	<td align="center"><%=rs("RETESTSTARTTIME")%></td>
	<td align="center"><%=rs("STARTQTY")%></td>
	<td align="center"><%=rs("SAMPLEQTY")%></td>
	<td align="center"><%=rs("MODELNAME")%></td>
	<td align="center"><%=rs("GOODQTY")%></td>
	<td align="center"><%=rs("REJECTQTY")%></td>
	<td align="center"><%=rs("LINELOSSQTY")%></td>
	<%
		'retestYield=1-(formatpercent(cdbl(rs("REJECTQTY"))/cdbl(rs("SAMPLEQTY")),2,-1))
		rejectQty=rs("REJECTQTY")
		sampleQty=rs("SAMPLEQTY")
		retestYield=formatpercent(1-(cint(rejectQty)/cint(sampleQty)),2,-1)
	%>
	<td align="center"><%=retestYield%></td>
</tr>
<%
i=i+1
rs.movenext
wend
else
%>
  <tr>
    <td height="20" colspan="10"><div align="center">No Records </div></td>
  </tr>
<%end if
rs.close%>
</table>

</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->