<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Job/JobCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/POCF_Open.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
jobnumber=replace(request.QueryString("jobnumber"),"$","+")
part_number_tag=request.QueryString("PART_NUMBER_TAG")
line_name=request.QueryString("LINE_NAME")
SQL="select * from tbl_Rework_Defectcode where REWORK_JOB_NUMBER='"&jobnumber&"'"
rsPR.open SQL,connPR,1,3
if not rsPR.eof then
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charsPRet=gb2312">
<title><%=application("SystemName")%></title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<!--#include virtual="/Language/Job/SubJobs/Lan_ReworkJobDetail.asp" -->
</head>

<body onLoad="language();">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
    <tr>
      <td height="20" colspan="4" class="t-c-greenCopy"><span id="inner_Browse"></span> (<%=jobnumber%>) </td>
    </tr>
    <tr>
      <td height="20" colspan="4" class="t-t-Borrow"><span id="inner_Summary"></span></td>
    </tr>
    <tr>
      <td width="109"><span id="inner_JobNumber"></span></td>
      <td width="229"><%= jobnumber %></td>
      <td width="111"><span id="inner_PartNumber"></span></td>
      <td width="215" height="20"><%= part_number_tag %></td>
    </tr>
    <tr>
      <td><span id="inner_Line"></span></td>
      <td height="20" colspan="3"><%= line_name %></td>
    </tr>
</table>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
    <tr>
      <td height="20" colspan="3" class="t-t-Borrow"><span id="inner_DefectCodeInfo"></span></td>
    </tr>
    <tr class="t-t-Borrow">
      <td><div align="center"><span id="inner_NO"></span></div></td>
      <td height="20"><div align="center"><span id="inner_DefectCodeName"></span></div></td>
      <td><div align="center"><span id="inner_DefectCodeQuantity"></span></div></td>
    </tr>
<%i=1
while not rsPR.eof%>
    <tr>
      <td><div align="center"><%=i%></div></td>
      <td height="20"><%=rsPR("DEFECT_CHINESE_NAME")%></td>
      <td><div align="center"><%=rsPR("QUANTITY")%></div></td>
    </tr>
<%
rsPR.movenext
i=i+1
wend
end if
rsPR.close%>
        <tr>
        <td colspan="3"><div align="center"><input name="Close" type="button" id="Close" value="Close" onClick="javascript:window.close()"></div></td>
      </tr>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/POCF_Close.asp" -->