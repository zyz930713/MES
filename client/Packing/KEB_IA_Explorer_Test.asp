<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/include/Functions.asp"-->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/TSD01_Open.asp" -->
<html>
<head>
<title>(Explorer_OQC出货确认)</title>
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

<form name="form2" method="post" action="KEB_IA_Explorer_Test.asp">
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

	strSql = "select [serialNumber],[testDateTime],[upDateTime],[pr_fail],[errorName] FROM [tsd].[IA_List] WHERE serialNumber = '"&Dcode&"' order by testDateTime"
	'response.Write(strSql)
	rs.open strSql,ConnTSD01,1,1
	if rs.eof And rs.bof then
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
		
		testDateTime = rs("testDateTime")
		prFail = rs("pr_fail")
		
		set oraRS = server.createobject("adodb.recordset")
		oraSQL = "select l.product,b.line_name,a.job_number from job_2d_code a inner join job b on a.JOB_NUMBER = b.JOB_NUMBER and a.SHEET_NUMBER = b.SHEET_NUMBER inner join line l on b.line_name = l.line_name where rownum < 2 and code = '"&Dcode&"'"
		oraRS.open oraSQL,conn,1,1
		if not rs.eof then
%>
		<table class="tableBorder" width="98%" border="1" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
		<tr bgcolor="#E8F1FF"><th>产品</th><th>线号</th><th>测试时间</th><th>结果</th></tr>
<%		
			while not rs.eof
			
				product = oraRS("product")
				lineName = oraRS("line_name")
				
				testStyle = "background:#ffffff;"
				if prFail = "1"	then	'如果是1,那么就是不良品,需要输出不良类型
					errorName = rs("errorName")
					testStyle = "background:red;"
				else
					errorName = "<font color='green'>通过</font>"
				end if
				%>
				<tr style="<%=testStyle%>" >
					<td align="center"><%=product%></td>
					<td align="center"><%=lineName%></td>
					<td align="center"><%=testDateTime%></td>
					<td width="70%" align="center" class="STYLE9"><%=errorName%></td>
				</tr>
				<%
				rs.movenext
			wend
		%>
		</table>
		<%
		end if
	end if
	
end if
%>
<BR>


</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->