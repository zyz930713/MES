<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/include/Functions.asp"-->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/TSD01_Open.asp" -->
<html>
<head>
<title>Grounding 电阻测试结果</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/css.css" rel="stylesheet" type="text/css">
<script type="text/javascript" src="../Scripts/jquery-1.8.3.js"></script>
<script type="text/javascript" src="../Scripts/jquery-ui-1.9.2/js/jquery-ui-1.9.2.custom.min.js"></script>
<script language="javascript" type="text/javascript">

</script>

<style type="text/css">
<!--
.STYLE5 {font-size: 18pt}
.STYLE6 {
	font-size: 14pt;
	font-weight: bold;
}
.STYLE7 {
	font-size: larger;
	font-family: Arial, Helvetica, sans-serif;
}
.STYLE8 {font-family: Arial, Helvetica, sans-serif}
.STYLE9 {font-size: 18pt; font-family: Arial, Helvetica, sans-serif; }
.STYLE10 {font-size: 14pt; font-weight: bold; font-family: Arial, Helvetica, sans-serif; }
body,td,th {
	font-family: Arial, Helvetica, sans-serif;
}
.RED{font-size: 22pt; font-family: Arial, Helvetica, sans-serif; }
.STYLE11 {font-size: x-large}

-->
</style>
</head>
<body onload= "javascript:document.all.Dcode.focus(); ">

<form name="form2" method="post" action="KEB_Grounding_Explorer.asp">
	<table width="100%" align="center">
		<tr bgcolor="#E8F1FF"> 
			<td width="74%" align="center"> 
				<span class="RED">2D Code ：</span>
				<input name="Dcode" type="text" id="Dcode" onFocus="this.value=''" value="" size="50">
				<input type="submit" name="Submit2" value="检  查">
			</td>
		</tr>
	</table>
</form>
<%
UserCode = trim(session("UserCode"))
Dcode = trim(request("Dcode"))
if Dcode <> "" then
	if len(Dcode) <> "17" then
		response.Write("<center>2D Barcode 位数不对！</center>")
		response.End()
	end if
	
	set sqlRs = server.createobject("adodb.recordset")
	strSql = "SELECT [serialnumber],[value],[testDateTime],[empNo],[pcName] FROM [dbo].[RowDataRes] WHERE serialNumber = '"&Dcode&"' order by testDateTime"
	'response.Write(strSql)
	sqlRs.open strSql,ConnTSD01,1,1
	if sqlRs.eof then
		response.Write("<center>2D Barcode 不存在！</center>")
		response.End()
	else
		PNO=mid(Dcode,12,4) 
		response.Write("<BR>")
		response.Write("<DIV align=center class=STYLE5>")
		DcodeA=left(Dcode,7)
		DcodeB=mid(Dcode,8,1)
		DcodeC=mid(Dcode,9,9)
		response.Write(DcodeA)
		response.Write("<font size=18>")
		response.Write(DcodeB)
		response.Write("</font>")
		response.Write(DcodeC)
		response.Write("&nbsp;&nbsp;&nbsp;")
		response.Write("</DIV>")
		response.Write("<BR>")
		
		set oraRS = server.createobject("adodb.recordset")
		oraSQL = "select l.product,b.line_name,a.job_number from job_2d_code a inner join job b on a.JOB_NUMBER = b.JOB_NUMBER and a.SHEET_NUMBER = b.SHEET_NUMBER inner join line l on b.line_name = l.line_name where rownum < 2 and code = '"&Dcode&"'"
		oraRS.open oraSQL,conn,1,1
		if not oraRS.eof then
			product = oraRS("product")
			lineName = oraRS("line_name")
		else
			product = "<font color='red'>工单未找到</font>"
			lineName = ""
		end if
		oraRS.close
		set oraRS = nothing
%>
		<table class="tableBorder" width="98%" border="1" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
		<tr bgcolor="#E8F1FF"><th>产品</th><th>线号</th><th>测试时间</th><th>结果</th><th>测试员</th><th>测试机</th></tr>
<%		
			while not sqlRs.eof
				testDateTime = sqlRs("testDateTime")
				empNo = sqlRs("empNo")
				pcName = sqlRs("pcName")

				TestValue = sqlRs("value")
				
				if not IsNumeric(TestValue) then
					TestValue = -999
					TestValue = "<font color='red'>"&TestValue&"</font>"
					testStyle = "background:#FFFF00;"
				else
					TestValue = csng(TestValue)
					if TestValue > 0 and TestValue < 0.4 then
						TestValue = "<font color='green'>"&Formatnumber(TestValue,3,-1)&"</font>"
						testStyle = "background:#ffffff;"
					else
						TestValue = "<font color='red'>"&Formatnumber(TestValue,3,-1)&"</font>"
						testStyle = "background:#FFFF00;"
					end if
				end if

				
				%>
				<tr style="<%=testStyle%>" >
					<td align="center"><%=product%></td>
					<td align="center"><%=lineName%></td>
					<td align="center"><%=testDateTime%></td>
					<td align="center"><%=TestValue%></td>
					<td align="center"><%=empNo%></td>
					<td align="center"><%=pcName%></td>
				</tr>
				<%
				sqlRs.movenext
			wend
		%>
		</table>
		<%
	end if
end if
%>
<BR>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
<!--#include virtual="/WOCF/TSD01_Close.asp" -->