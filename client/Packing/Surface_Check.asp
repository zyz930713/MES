<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%session("Page_Role") = ",Surface_Check"%>
<!--#include virtual="/Functions/Login_Check.asp"-->
<!--#include virtual="/Functions/Role_Check.asp"-->
<!--#include virtual="/include/Functions.asp"-->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<html>
<head>
<title> (外观检查追溯)</title>
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

<form name="form2" method="post" action="Surface_Check.asp">
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

	strSql = "select code,FOI_USER FROM job_2d_code WHERE CODE = '"&Dcode&"'"
	'response.Write(strSql)
	rs.open strSql,conn,1,1
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

		FoiUser = rs("FOI_USER")
		
		if FoiUser <> "" then
			FoiUser = "已经与: " & FoiUser & " 绑定!"
		else
			UpdateSql = "update job_2d_code set foi_user = '"&UserCode&"' WHERE CODE = '"&Dcode&"'"
			conn.execute(UpdateSql)
			FoiUser = "<font color='red'>已绑定到: " & UserCode & "</font>"
		end if
	end if
	
end if
%>
<BR>
<table class="tableBorder" width="100%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
	<tr bgcolor="#E8F1FF">
		<td width="70%" align="center" class="STYLE9"><%=FoiUser%></td>
	</tr>
</table>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->