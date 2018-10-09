
<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>

<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Job/JobCheck.asp" -->
<!--#include virtual="/Job/IsDBA.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<%
	dim ReworkJobNumber
	dim SubJobnumber
	dim ReworkType
	dim StartQty
	dim GoodQty
	dim RejectQty
	
	ReworkJobNumber=request("jobnumber") 
	SubJobnumber=request("SUB_JOB_NUMBER") 
	ReworkType=request("REWORK_TYPE") 
	
	StartQty=request("StartQty") 
	GoodQty=request("GoodQty") 
	RejectQty=request("RejectQty") 

	if(StartQty="" AND GoodQty="" AND RejectQty="" ) then
		SQL="select * FROM JOB_REWORK where REWORK_JOB_NUMBER='"&ReworkJobNumber&"' AND SUB_JOB_NUMBER='"&SubJobnumber&"' AND REWORK_TYPE='"&ReworkType&"' ORDER BY REWORK_END_TIME DESC,REWORK_JOB_NUMBER,TO_NUMBER(SUB_JOB_NUMBER)"
		rs.open SQL,conn,1,3
		IF NOT RS.EOF then
			StartQty=rs("SUB_JOB_REWORK_QTY").value
			GoodQty=rs("REWORK_GOOD_QTY").value
			RejectQty=rs("REWORK_REJECT_QTY").value
		end if
	else
		IF cdbl(StartQty)<>(cdbl(GoodQty)+cdbl(RejectQty)) THEN
			response.write("<script>window.alert('数量输入有误!')</script>")
		else
			SQL="select * FROM JOB_REWORK where REWORK_JOB_NUMBER='"&ReworkJobNumber&"' AND SUB_JOB_NUMBER='"&SubJobnumber&"' AND REWORK_TYPE='"&ReworkType&"' ORDER BY REWORK_END_TIME DESC,REWORK_JOB_NUMBER,TO_NUMBER(SUB_JOB_NUMBER)"
			rs.open SQL,conn,1,3
			IF NOT RS.EOF then
				rs("SUB_JOB_REWORK_QTY").value=StartQty
				rs("REWORK_GOOD_QTY").value=GoodQty
				rs("REWORK_REJECT_QTY").value=RejectQty
				rs.update
				rs.close
				response.write("<script>window.alert('更新成功!')</script>")
			end if
		end if
	end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>ReworkJobUpdate</title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<link href="/CSS/dynCalendar.css" rel="stylesheet" type="text/css">
<script language=javascript src="/Components/sniffer.js" type=text/javascript></script>
<script language=javascript src="/Components/dynCalendar.js" type=text/javascript></script>
<script>
		function CheckNumber(ControlID)
		{
			var ControlText;
			ControlText=document.getElementById(ControlID).value;
			if(ControlText!="" && isNaN(ControlText))
			{
				window.alert("请输入数字!");
				document.getElementById(ControlID).focus();
			}
		}
</script>
<!--#include virtual="/Language/Job/Reworkjob/Lan_Job.asp" -->
</head>
<body onLoad="language();">
	<form name="form1" method="get" action="/Job/rework/ReworkQuantityUpdate.asp">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
	<td colspan="2" align="center"><span id="inner_UpdateTitle"  style="font-size:20px"></span></td>
</tr>
<tr>
	<td><span id="inner_JobNumber"></span></td>
	<td><%=ReworkJobNumber%><input type="hidden" id="jobnumber" name="jobnumber" value="<%=ReworkJobNumber%>"/><input type="hidden" name="SUB_JOB_NUMBER" id="SUB_JOB_NUMBER" value="<%=SubJobnumber%>"/></td>
</tr>
<tr>
	<td><span id="inner_SearchReworkType"></span></td>
	<td><% if ReworkType="0" then response.write "Readjust"  end if 
	   if ReworkType="1" then response.write "Teardown" end if  
	   if ReworkType="2" then response.write "Slapping"  end if %><input type="hidden" id="REWORK_TYPE" name="REWORK_TYPE"value="<%=ReworkType%>"/></td>
</tr>
<tr>
	<td><span id="inner_ReworkQty"></span></td>
	<td><input type="text" id="StartQty" NAME="StartQty"  value="<%=StartQty%>" onBlur="CheckNumber('StartQty')"/></td>
</tr>
<tr>
	<td><span id="inner_GoodQty"></span></td>
	<td><input type="text" id="GoodQty" NAME="GoodQty"  value="<%=GoodQty%>" onBlur="CheckNumber('GoodQty')"/></td>
</tr>
<tr>
	<td><span id="inner_RejectQty"></span></td>
	<td><input type="text" id="REJECTQTY" NAME="REJECTQTY"  value="<%=RejectQty%>" onBlur="CheckNumber('REJECTQTY')"/></td>
</tr>
<tr>
	<td colspan="2"><input type="button" id="btnUpdate" onclick="form1.submit();" /></td>
</tr>
</body>
</html>
