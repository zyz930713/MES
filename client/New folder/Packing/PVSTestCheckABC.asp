<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/WOCF/PVS_Open.asp" -->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=application("SystemName")%></title>
</head>
<link href="/CSS/GeneralKES1.css" rel="stylesheet" type="text/css">

<body bgcolor="#339966">
<%
code=request("txt_2d_code")
isQuery=false
if code <> "" then
	isQuery=true
	sql="select * from TempCheckJSQ where Dcode='"&code&"'"
	rsPVS.open sql,connPVS,1,3
	
end if
%>
<table width="100%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
	<tr>
        <td colspan="4" class="t-t-DarkBlue" align="center">PVS Test Data Check ABC</td>
    </tr>
	<tr align="center">
		<td >NO 序列</td>
		<td >2D Code 二维码</td>
		<td >CodeType</td>
		<td >CheckABC 时间</td>
	</tr>
<%	
if isQuery then
i=1
while not rsPVS.eof
	result="Pass"
	trColor="#00CC00"
	defect="&nbsp;"
	criterion="&nbsp;"
	if rsPVS("CodeType")="A" then
	        trColor="#00CC00"
	elseif rsPVS("CodeType")="B" then
	
			trColor="#FFFF00"
	else
			trColor="#FF0000"
	end if
	'end if
%>
	<tr align="center" bgcolor="<%=trColor%>" >
		<td><%=i%></td>
		<td><%=rsPVS("DCode")%></td>
		<td><%=rsPVS("CodeType")%></td>
		<td><%=rsPVS("CodeDate")%></td>
	</tr>	
<%
i=i+1
rsPVS.movenext		
wend
end if
%>
</table>
</body>
</html>
<!--#include virtual="/WOCF/PVS_Close.asp" -->