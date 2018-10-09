<%
response.charset = "UTF-8"
if session("ShiftEditUser") = "" then call sussLoctionHref("网络超时或者您还没有登录请登录","user_login.asp")
if session("ShiftEditType") <> "OEM" then response.redirect "default.asp"
%>

<!--#include file= "Conn_open.asp"-->
<!--#include file= "include/function.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>OEM - Editor</title>
<link rel="stylesheet" type="text/css" href="Styles/basic.css">
<script language="javascript" type="text/javascript" src="include/DatePicker/WdatePicker.js"></script>
<style type="text/css">
body{
	font:normal 16px/16px "Arial";
}
table{
	border-collapse:collapse;
	border-spacing:0;
	border:2px solid black;
}
table tr{
	height:26px;
	line-height:26px;
	background:#F2F2F2;
	}
table th{
	border-color:black;
	font:normal 16px/16px "Arial";
}
table td{
	font:normal 14px/16px "Arial";
	border-color:black;
}
body,td,th {
	font-family: Arial;
}
</style>
</head>
<body bgcolor="#339966">
</br>
<center>
<table border="1" cellspacing="0" cellpadding="0" width="600">
<tr><td colspan="7" class="TitleLable"><p style="height:30px;line-height:30px;"><strong>项目开线/停线设置</strong></p></td></tr>
<tr><th>项目名称</th><th>线号</th><th>生产目标</th><th>过程良率</th><th>声学良率</th><th>外观良率</th><th>项目状态</th></tr>
<%
RoleGroup = session("ShiftEditRole")
if RoleGroup = "0" or RoleGroup = "1" then RoleGroup = "%"
productname = request("productname")
linename = request("linename")
ActiveSet = request("ActiveSet")

if ActiveSet = "开线" or ActiveSet = "停线" then
	if ActiveSet = "开线" then
		conn.execute("UPDATE report_oem_target SET IsActive = '0' WHERE productname = '"&productname&"' and linename = '"&linename&"'")
	elseif ActiveSet = "停线" then
		conn.execute("UPDATE report_oem_target SET IsActive = '1' WHERE productname = '"&productname&"' and linename = '"&linename&"'")
	end if
end if

moditarget = request("moditarget")

if moditarget = "修改" then
	PcsTarget = request("PcsTarget")
	ProcessTarget = request("ProcessTarget")
	AcousticTarget = request("AcousticTarget")
	CosmeticTarget = request("CosmeticTarget")
	'response.write "我要修改"
	'response.end
	SqlModi = "UPDATE report_oem_target SET PcsTarget = '"&PcsTarget&"',ProcessTarget = '"&ProcessTarget&"',AcousticTarget = '"&AcousticTarget&"',CosmeticTarget = '"&CosmeticTarget&"' WHERE productname = '"&productname&"' and linename = '"&linename&"'"
	conn.execute(SqlModi)
	call sussLoctionHref("修改成功","SetOem.asp")
end if

ModStr = "<input type='submit' name='moditarget' value='修改' />"

LoopNo = 1
set rs = Server.CreateObject("adodb.recordset")
sql = "select productname,linename,pcstarget,processtarget,acoustictarget,cosmetictarget,isactive from report_oem_target where Egroup like '"&RoleGroup&"' order by productname,linename"
rs.open sql,conn,1,1
do while not rs.eof
ProductName = rs("productname")
linename = rs("linename")
PcsTarget = rs("PcsTarget")
ProcessTarget = rs("ProcessTarget")
AcousticTarget = rs("AcousticTarget")
CosmeticTarget = rs("CosmeticTarget")
IsActive = rs("IsActive")

if IsActive then
	ProActive = "开线"
	ActiveStr = "style='background:#00FF00;'"
else
	ProActive = "停线"
	ActiveStr = "style='background:#FF0000;'"
end if

if (LoopNo mod 2 = 0) then
	GgColor = "#FFFFFF"
else
	GgColor = "#E6E6E6"
end if
%>
<form method="post" name="form<%=ProductName%>" action="SetOem.asp">
<tr style="background:<%=GgColor%>;">
	<td>
		<%=ProductName%><input type="hidden" name="ProductName" value="<%=ProductName%>" /><input type="hidden" name="linename" value="<%=linename%>" />
	</td>
	<td>
		<%=linename%>
	</td>
	<td>
		<input type="text" name="PcsTarget" value="<%=PcsTarget%>" size="4" />
	</td>
	<td>
		<input type="text" name="ProcessTarget" value="<%=ProcessTarget%>" size="4" />
	</td>
	<td>
		<input type="text" name="AcousticTarget" value="<%=AcousticTarget%>" size="4" />
	</td>
	<td>
		<input type="text" name="CosmeticTarget" value="<%=CosmeticTarget%>" size="4" />
	</td>
	<td>
		<%=ModStr%> <input type="submit" name="ActiveSet" value="<%=ProActive%>" <%=ActiveStr%> />
	</td>
</tr>
</form>
<%
LoopNo = LoopNo + 1
rs.movenext
loop
%>
</table>


<br/>
<input type="button" value="返回" onClick="window.location.reload('InputOem.asp');" />
　　　　
<input type="button" value="修改密码" onClick="window.location.reload('modify_pass.asp');" />
　　　　
<input type="button" value="退出" onClick="window.location.reload('user_logout.asp');" />
</center>


<script type="text/javascript">
function $(obj){
   return document.getElementById(obj);
}
<%=ScriptStr%>
</script>
</body>
</html>
