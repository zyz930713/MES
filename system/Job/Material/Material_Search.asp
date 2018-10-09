<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Job/JobCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Functions/GetAction.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
MaterialLotNumber=request("txtMaterialLotNumber")

set rs1=server.CreateObject("adodb.recordset")
SQL=" SELECT material_part_number, material_lot_number, job_number, sheet_number, dispatch__material_quantity, kitting_datetime FROM MR_DISPATCH WHERE 1=1 "
SQL=SQL+" AND lower(material_lot_number)='"+lcase(MaterialLotNumber)+"'"
session("SQL")=SQL
rs.open SQL,conn,1,3
%>
<!--#include virtual="/Components/PageSelect.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System</title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<link href="/CSS/dynCalendar.css" rel="stylesheet" type="text/css">
<script language=javascript src="/Components/sniffer.js" type=text/javascript></script>
<script language=javascript src="/Components/dynCalendar.js" type=text/javascript></script>
<!--#include virtual="/Language/Job/Material/Lan_Material.asp" -->
<script language="javascript">
function DoQuery()
{
	 if(form1.txtMaterialLotNumber.value=="")
	 {
	 	alert("Please input Material Lot Number!");
		return;
	 }
	 else
	 {
	 	form1.action="Material_Search.asp";
	 	form1.submit();
	 }
}
</script>
</head>

<body onLoad="language();language_page()">
<form name="form1" method="post">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="3" class="t-c-greenCopy"><span id="inner_Search"></span></td>
  </tr>
  <tr>
    <td height="20">Material Lot Number</td>
    <td height="20"><input name="txtMaterialLotNumber" type="text" id="txtMaterialLotNumber" value="<%=MaterialLotNumber%>"></td>
    <td><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="DoQuery()"></td>
    </tr>
</table>
</form>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="9" class="t-c-greenCopy"><span id="inner_Browse"></span></td>
</tr>
<tr>
  <td height="20" colspan="9"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_NO"></span></div></td>
  <td class="t-t-Borrow"><div align="center">Material Part Number</div></td>
  <td class="t-t-Borrow"><div align="center">Material Lot Number</div></td>
  <td class="t-t-Borrow"><div align="center">Job Number</div></td>
  <td class="t-t-Borrow"><div align="center">Sheet Number</div></td>
  <td class="t-t-Borrow"><div align="center">Qty</div></td>
  <td class="t-t-Borrow"><div align="center">Kitting Time</div></td>
  </tr>
<%
i=1
if not rs.eof then
rs.absolutepage=session("strpagenum")
while not rs.eof 
%>
	<tr>
	  <td height="20"><div align="center"><%=i%></div></td>
		<td><div align="center"><%= rs("material_part_number")%></div></td>
		<td><div align="center"><%= rs("material_lot_number")%></div></td>
		<td><div align="center"><%= rs("job_number")%></div></td>
		<td><div align="center"><%= rs("sheet_number")%></div></td>
		<td><div align="center"><%= rs("dispatch__material_quantity")%></div></td>
		<td><div align="center"><%= rs("kitting_datetime")%></div></td>
  </tr>
<%
i=i+1
rs.movenext
wend
else
%>
  <tr>
    <td height="20" colspan="7"><div align="center">No Records </div></td>
  </tr>
<%end if
rs.close
set rs1=nothing%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->