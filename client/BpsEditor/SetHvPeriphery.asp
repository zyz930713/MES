<%
response.charset = "UTF-8"
if session("ShiftEditUser") = "" then call sussLoctionHref("网络超时或者您还没有登录请登录","user_login.asp")
if session("ShiftEditType") <> "TARGET" then response.redirect "default.asp"
ReadonlyStr = "readonly"
if session("ShiftEditRole") = "0" then ReadonlyStr = ""
%>

<!--#include file= "Conn_open.asp"-->
<!--#include file= "include/function.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<script language="javascript" type="text/javascript" src="../include/ThreeMenuMa.js"></script>
<script language="javascript" type="text/javascript" src="include/DatePicker/WdatePicker.js"></script>
<title>Periphery Target Set</title>
</head>
<style type="text/css">
table tr{
	height:26px;
	line-height:26px;
	background:#F2F2F2;
}
table tr td{
	text-align:center;
}
</style>
<body bgcolor="#339966">
</br>
<%
productname = request("productname")
MaName = request("MaName")
SearchTarget = request("SearchTarget")
TargetSet = request("TargetSet")
RoleGroup = session("ShiftEditRole")
moditarget = request("moditarget")

if TargetSet = "新增" or TargetSet = "修改" then
	MaID = request("MaID")
	MaSum = request("MaSum")
	tdate = request("tdate")
	Adate = request("Adate")
	PcsTarget = request("PcsTarget")
	FORTarget = request("FORTarget")
		
	if TargetSet = "新增" then
		AddSql = "insert into dbo.Cfg_Periphery([ProductName],[MaName],[MaSum],[ADate],[AName],[UDate],[UName],[TDate],[PcsTarget],[FORTarget])"
		AddSql = AddSql & " values('"&productname&"','"&MaName&"','"&MaSum&"',getdate(),'"&session("ShiftEditUser")&"',getdate(),'"&session("ShiftEditUser")&"','"&tdate&"','"&PcsTarget&"','"&FORTarget&"')"
		'response.write AddSql
		'response.end
		ConnSql.execute(AddSql)
		call sussLoctionHref("添加成功","SetHvPeriphery.asp?SearchTarget=ok&ProductName="&ProductName&"&MaName="&MaName)
	end if
	
	if TargetSet = "修改" then
		ModiSql = "UPDATE dbo.Cfg_Periphery SET udate = getdate(),MaSum = '"&MaSum&"',uname = '"&session("ShiftEditUser")&"',tdate = '"&tdate&"',PcsTarget='"&PcsTarget&"',FORTarget = '"&FORTarget&"' WHERE Ma_ID = '"&MaID&"'"
		'response.write ModiSql
		'response.end
		ConnSql.execute(ModiSql)
		call sussLoctionHref("修改成功","SetHvPeriphery.asp?SearchTarget=ok&ProductName="&ProductName&"&MaName="&MaName)
	end if
end if
%>
<center>
<form name="SetHv" method="post" action="SetHvPeriphery.asp" >
		<table width="800" border="1" cellspacing="0" cellpadding="0">
			<tr style="height:30px;line-height:30px;">
				<td align="center">
				产品：<select name="ProductName" id="TmProduct"></select>
				设备：<select name="MaName" id="TmLine" ></select>
				<select name="CellName" id="TmCell" style="display:none" ></select>
					<script type="text/javascript">
						MenuInit('TmProduct','TmLine','TmCell','<%=ProductName%>','<%=MaName%>','1');
					</script>
				<input type="Submit" name="SearchTarget" value="刷新" />
				</td>
			</tr>
			</table>
</form>


<%
if SearchTarget = "刷新" or SearchTarget = "ok" then
%>
<table border="1" cellspacing="0" cellpadding="0" width="800">
<tr style="height:30px;line-height:30px;">
	<th>项目名称</th><th>设备</th><th>设备数量</th><th>Target日期</th><th>产出</th><th>FOR</th><th>操作</th>
</tr>
	<form method="post" name="AddForm" action="SetHvPeriphery.asp">
		<tr style="background:#FFFFFF;">
			<td>
				<%=ProductName%><input type="hidden" name="ProductName" value="<%=ProductName%>" />
			</td>
			<td>
				<%=MaName%><input type="hidden" name="MaName" value="<%=MaName%>" />
			</td>
			<td>
				<input type="text" name="MaSum" value="0" size="4" />
			</td>
			<td>
				<input type="text" name="tdate" class="Wdate" id="d122" onFocus="WdatePicker({isShowWeek:true,onpicked:function() {$dp.$('d122_1').value=$dp.cal.getP('W','W');$dp.$('d122_2').value=$dp.cal.getP('W','WW');}})" value="<%=tdate%>" size="10" />
			</td>
			<td>
				<input type="text" name="PcsTarget" value="0" size="4" />
			</td>
			<td>
				<input type="text" name="FORTarget" value="0" size="4" />
			</td>
			<td>
				<input type="submit" name="TargetSet" value="新增" style="background:#FFBFBF;" />
			</td>
		</tr>
	</form>
<%
	LoopNo = 1
	set rs = Server.CreateObject("adodb.recordset")
	sql = "select [Ma_ID],[ProductName],[MaName],[MaSum],[ADate],[AName],[UDate],[UName],[TDate],[PcsTarget],[FORTarget] from dbo.Cfg_Periphery"
	sql = sql & " where productname = '"&productname&"' and MaName = '"&MaName&"' order by ADate desc"
	'response.write sql
	'response.end
	rs.open sql,ConnSql,1,1
	do while not rs.eof
		MaID = rs("Ma_ID")
		ProductName = rs("productname")
		MaName = rs("MaName")
		MaSum = rs("MaSum")
		uname = rs("uname")
		tdate = rs("tdate")
		Adate = rs("Adate")
		PcsTarget = rs("PcsTarget")
		FORTarget = Formatnumber(rs("FORTarget"),2,-1)
		
		if (LoopNo mod 2 = 0) then
			GgColor = "#FFFFFF"
		else
			GgColor = "#E6E6E6"
		end if
		%>
		<form method="post" name="ListForm" action="SetHvPeriphery.asp">
		<tr style="background:<%=GgColor%>;">
			<td>
				<%=ProductName%><input type="hidden" name="MaID" value="<%=MaID%>" /><input type="hidden" name="ProductName" value="<%=ProductName%>" /><input type="hidden" name="MaName" value="<%=MaName%>" />
			</td>
			<td>
				<%=MaName%>
			</td>
			<td>
				<input type="text" name="MaSum" value="<%=MaSum%>" size="4" />
			</td>
			<td>
				<input type="text" name="tdate" class="Wdate" id="mdate" onFocus="WdatePicker({isShowWeek:true,onpicked:function() {$dp.$('mdate_1').value=$dp.cal.getP('W','W');$dp.$('mdate_2').value=$dp.cal.getP('W','WW');}})" value="<%=tdate%>" size="10" />
			</td>
			<td>
				<input type="text" name="PcsTarget" value="<%=PcsTarget%>" size="4" />
			</td>
			<td>
				<input type="text" name="FORTarget" value="<%=FORTarget%>" size="4" />
			</td>
			<td>
				<input type="submit" name="TargetSet" value="修改" style="background:#00FF00;" />
			</td>
		</tr>
		</form>
		<%
		LoopNo = LoopNo + 1
		rs.movenext
	loop
end if
%>
</table>

</br>
<table width="800" border="1" cellspacing="0" cellpadding="0">
	<tr>
		<td colspan="5" align="center">
		<%if session("ShiftEditRole") = "0" then%>
			<input type="button" value="设置线上Target" onClick="window.location.reload('SetHv.asp');" />　　　　
		<%end if%>
		<input type="button" value="修改密码" onClick="window.location.reload('modify_pass.asp');" />　　　　
		<input type="button" value="Close关闭" onClick="window.location.reload('user_logout.asp');" />
		</td>
	</tr>
</table>
</center>

</body>
</html>
