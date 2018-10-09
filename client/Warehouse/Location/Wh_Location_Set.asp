<%
	response.Expires=0
	response.CacheControl="no-cache"
	'if instr(session("role"),",JOB_ADMINISTRATOR")<>0 or instr(session("role"),",JOB_READER")<>0 then
%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<title>WH_Management</title>
<link href="/CSS/GeneralKES1.css" rel="stylesheet" type="text/css">
</head>
<body style="margin:20px;background-color:#339966;">
<%
set Rs = Server.CreateObject("adodb.recordset")
LoctorSql = "select loctor_name from wh_cfg_loctor group by loctor_name order by loctor_name"
Rs.open LoctorSql,conn,1,1
LoctorStr = ""
while not Rs.eof
	LoctorStr = LoctorStr & "<option value="&("loctor_name")&">"&Rs("loctor_name")&"</option>"
Rs.movenext
wend
%>
<center>
  <table width="500" border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF" >
    <tr>
    <td height="20" class="t-t-DarkBlue" align="center">库位编号管理</td>
  </tr>
  <tr>
    <td>开发中...</td>
  </tr>
</table>
</center>
</body> 
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->