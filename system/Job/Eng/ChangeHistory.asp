
<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
	
		batchno=request.querystring("batchno")
		 
	
		SQL="SELECT * FROM Change_Model_History WHERE BatchNo='"+batchno+"' order by ChangeDateTime desc"
		set rsModelName=server.createobject("adodb.recordset")
		rsModelName.open SQL,conn,1,3
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<link href="../../CSS/General.css" rel="stylesheet" type="text/css">
<Script>
 
</script>
</head>

<body>

<form  method="post" name="form1" target="_self">
<table width="100%">
			<tr>
				<Td class="t-c-greenCopy">BatchNo</Td>
				<Td class="t-c-greenCopy">Old Model</Td>
				<Td class="t-c-greenCopy">New Model</Td>
				<Td class="t-c-greenCopy">Change Person</Td>
				<Td class="t-c-greenCopy">Change Time</Td>
				<Td class="t-c-greenCopy">IsCurrentVersion</Td>
			</tr>
				<%for i=1 to rsModelName.recordcount%>
			<tr>
				<td class="t-c-GrayLight"><%=rsModelName("Batchno")%></td>
				<td class="t-c-GrayLight"><%=rsModelName("PrevoiusModel")%></td>
				<td class="t-c-GrayLight"><%=rsModelName("CurrentModel")%></td>
				<td class="t-c-GrayLight"><%=rsModelName("ChangePerson")%></td>
				<td class="t-c-GrayLight"><%=rsModelName("ChangeDateTime")%></td>
				<td class="t-c-GrayLight"><%=rsModelName("IsCurrentVersion")%></td>
			</tr>
			<%
				rsModelName.movenext
			NEXT 
			%>
			<tr>
				<td><input type="button" name="btnBack" value="Back" onclick="javascript:history.back()"</td>
			</tr>
</table>
</form>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
