<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Reports/FinalYield/FinalYieldCheck.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
factory=trim(request.Form("factory"))
fromdate=request.Form("jfromdate")
todate=request.Form("jtodate")
fromhour=request.Form("jfromhour")
tohour=request.Form("jtohour")
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>Untitled Document</title>
<script language="javascript">
function shiftframe(type)
{
	if(type=="Table")
	{
	document.all.DIVTable.style.visibility="visible";
	document.all.DIVChart.style.visibility="hidden";
	}
	else
	{
	document.all.DIVTable.style.visibility="hidden";
	document.all.DIVChart.style.visibility="visible";
	}
}
</script>
<link href="/CSS/General.css" rel="stylesheet" type="text/css" />
</head>

<body>
<table width="100%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr>
    <td valign="middle"><span style="cursor:hand" onclick="shiftframe("Table")" class="FINBig"><img src="/Images/IconReportTable.GIF" width="36" height="36" />View Tables</span></span></td>
    <td valign="middle"><span style="cursor:hand" onclick="shiftframe("Chart")" class="FINBig"><img src="/Images/IconReportChart.GIF" width="36" height="36" />View Chart</span></td>
  </tr>
  <tr>
    <td colspan="2">
	<span id="DIVTable" style="visibility:visible"><iframe src="/Reports/FinalYield/FinalFamilyYield/GenerateFinalFamilyYield.asp?factory=<%=factory%>&fromdate=<%=fromdate%>&todate=<%=todate%>&fromhour=<%=fromhour%>&tohour=<%=tohour%>" width="100%" height="500" scrolling="auto" id="FrameFinalYieldTable"></iframe></span>
	<span id="DIVChart" style="visibility:hidden"><iframe src="/Reports/FinalYield/FinalFamilyYield/GenerateFinalFamilyYield.asp?factory=<%=factory%>&fromdate=<%=fromdate%>&todate=<%=todate%>&fromhour=<%=fromhour%>&tohour=<%=tohour%>" width="100%" height="500" scrolling="auto" id="FrameFinalYieldChart"></iframe></span></td>
  </tr>
</table>
</body>
</html>
