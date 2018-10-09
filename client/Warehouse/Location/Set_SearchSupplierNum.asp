<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%session("Page_Role") = ",WH_Location_Set"%>
<!--#include virtual="/Functions/Login_Check.asp"-->
<!--#include virtual="/Functions/Role_Check.asp"-->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/include/Functions.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>供应商信息查询</title>
<link href="Styles/Basic.css" rel="stylesheet" type="text/css">
<script type="text/javascript" src="../../Scripts/jquery-1.8.3.js"></script>
<style type="text/css">
tr.change:hover
{
	background-color:#D1D1D1
}
</style>
</head>
<body style="margin:20px;background-color:#339966;">
<center>
<%
SubmitSearch = request("SubmitSearch")
VendorNum = trim(request("VendorNum"))
VendorNameEn = trim(request("VendorNameEn"))
VendorNameCn = trim(request("VendorNameCn"))
%>
<form name="form1" method="post" action="Set_SearchSupplierNum.asp" >
<table width="1300" border="1" align="center" cellpadding="0" cellspacing="0" >
	<tr>
		<td height="20" colspan="2" class="t-t-DarkBlue">
			<div align="center">查找供应商编号</div>
		</td>
	</tr>	
	<tr style="background:#cccccc;">
		<td>
		供应商(编号):<input type="text" name="VendorNum" value="<%=VendorNum%>" />&nbsp;
		供应商(英文):<input type="text" name="VendorNameEn" value="<%=VendorNameEn%>" />&nbsp;
		供应商(中文):<input type="text" name="VendorNameCn" value="<%=VendorNameCn%>" />&nbsp;
			<input type="submit" name="SubmitSearch" value="查找" />&nbsp;
		</td>
	</tr>
</table>
</form>
<%
if SubmitSearch = "查找" then
	SQL = "select rownum,vendor_num,vendor_name,vendor_name_alt from suppliers where rownum < 200 and vendor_num like '%"&VendorNum&"%' and Upper(vendor_name) LIKE Upper('%"&VendorNameEn&"%') and vendor_name_alt like '%"&VendorNameCn&"%'"
	'response.write SQL
	rs.open SQL,conn,1,1
	if not rs.eof then
		VendorNum = rs("vendor_num")
		VendorNameEn = rs("vendor_name")
		VendorNameCn = rs("vendor_name_alt")
		%>
		<table width='1300' border='1' cellpadding='0' cellspacing='0' style="background:#ffffff;">
		<tr class="t-t-DarkGray"><th>编号</th><th>供应商(英文)</th><th>供应商(中文)</th></tr>
		<%
		while not rs.eof
		VendorNum = rs("vendor_num")
		VendorNameEn = rs("vendor_name")
		VendorNameCn = rs("vendor_name_alt")
	%>
		<tr class="change">
			<td width="70"><%=VendorNum%></td>
			<td align="left"><%=VendorNameEn%></td>
			<td align="left"><%=VendorNameCn%>&nbsp;</td>
		</tr>
		<%
		rs.movenext
		wend
	end if
	rs.close
response.write "</table>"
end if
%>
<br>
  <table width="1300" border="1" cellspacing="0" cellpadding="0">
    <tr style="background:#cccccc;">
      <td align="center">
      <input type="button" value="修改密码" onClick="window.location.reload('../../Functions/UserPass_Modify.asp');" />
      　　　　
      <input type="button" value="Close关闭" onClick="window.location.reload('../../Functions/User_Logout.asp');" /></td>
    </tr>
  </table>
</center>
</body> 
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->