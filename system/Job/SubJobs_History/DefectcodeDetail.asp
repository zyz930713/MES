<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/Job/JobCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
jobnumber=request.QueryString("jobnumber")
sheetnumber=request.QueryString("sheetnumber")
if sheetnumber="" then
SQL="Select JD.*,S.STATION_NAME,D.DEFECT_NAME from JOB_DEFECTCODES JD inner join STATION S on JD.STATION_ID=S.NID left join DEFECTCODE D on JD.DEFECT_CODE_ID=D.NID where JD.JOB_NUMBER='"&jobnumber&"'"
else
SQL="Select JD.*,S.STATION_NAME,D.DEFECT_NAME from JOB_DEFECTCODES JD inner join STATION S on JD.STATION_ID=S.NID left join DEFECTCODE D on JD.DEFECT_CODE_ID=D.NID where JD.JOB_NUMBER='"&jobnumber&"' and JD.SHEET_NUMBER='"&sheetnumber&"'"
end if
rs.open SQL,conn,1,3
if not rs.eof then
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<!--#include virtual="/Language/Job/SubJobs/Lan_DefectCodeDetail.asp" -->
</head>

<body onLoad="language()">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="10" nowrap class="t-t-Borrow"><span id="inner_DefectCodeInfo"></span></td>
</tr>
<tr>
  <td class="t-b-blue"><div align="center"><span id="inner_No"></span></div></td>
  <td class="t-b-blue"><div align="center"><span id="inner_JobNumber"></span></div></td>
  <td height="20" class="t-b-blue"><div align="center"><span id="inner_Station"></span></div></td>
  <td height="20" class="t-b-blue"><div align="center"><span id="inner_DefectName"></span></div></td>
  <td height="20" class="t-b-blue"><div align="center"><span id="inner_DefectQuantity"></span></div></td>
  <td class="t-b-blue"><div align="center"><span id="inner_StepQuantity"></span></div></td>
</tr>
<%
i=1
total_quantity=0
plus_quantity=0
this_job_number=rs("JOB_NUMBER")&"-"&repeatstring(rs("SHEET_NUMBER"),"0",3)
while not rs.eof 
	if this_job_number<>rs("JOB_NUMBER")&"-"&repeatstring(rs("SHEET_NUMBER"),"0",3) then
	step_quantity=plus_quantity
	plus_quantity=0
	plus_quantity=plus_quantity+cint(rs("DEFECT_QUANTITY"))
%>
<tr>
  <td height="20" colspan="5">&nbsp;</td>
  <td><div align="center"><%=step_quantity%>&nbsp;</div></td>
</tr>
<%
	else
	plus_quantity=plus_quantity+cint(rs("DEFECT_QUANTITY"))
	step_quantity=0
	end if
%>
<tr>
  <td><div align="center"><%= i %></div></td>
  <td><div align="center"><%=rs("JOB_NUMBER")%>-<%=repeatstring(rs("SHEET_NUMBER"),"0",3)%></div></td>
  <td height="20"><div align="center"><%= rs("STATION_NAME") %></div></td>
  <td height="20"><div align="center"><%= rs("DEFECT_NAME") %> (<%= rs("DEFECT_CODE_ID") %>)</div></td>
  <td height="20"><div align="center"><%= rs("DEFECT_QUANTITY") %></div></td>
  <td><div align="center">&nbsp;</div></td>
</tr>
<%
this_job_number=rs("JOB_NUMBER")&"-"&repeatstring(rs("SHEET_NUMBER"),"0",3)
total_quantity=total_quantity+cint(rs("DEFECT_QUANTITY"))
i=i+1
rs.movenext
wend
step_quantity=plus_quantity
%>
<tr>
  <td height="20" colspan="5">&nbsp;</td>
  <td><div align="center"><%=step_quantity%>&nbsp;</div></td>
</tr>
<tr>
  <td><div align="center"><span id="inner_Total"></span></div></td>
  <td height="20" colspan="3"><div align="center">&nbsp;</div> </td>
  <td height="20"><div align="center"><%=total_quantity%></div></td>
  <td><div align="center">&nbsp;</div></td>
</tr>
<tr>
  <td height="20" colspan="6"><div align="center">
	<input name="Close" type="reset" id="Close" value="Close" onClick="window.close()">
  </div></td>
</tr>
<%
end if
rs.close%>
</table>
<div align="center"></div>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->