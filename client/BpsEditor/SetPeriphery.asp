<%
response.charset = "UTF-8"
if session("ShiftEditUser") = "" then call sussLoctionHref("网络超时或者您还没有登录请登录","user_login.asp")
if session("ShiftEditType") = "OEM" then response.redirect "InputOem.asp"
ReadonlyStr = "readonly"
if session("ShiftEditRole") = "0" then ReadonlyStr = ""
%>

<!--#include file= "Conn_open.asp"-->
<!--#include file= "include/function.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<script language="javascript" type="text/javascript" src="../include/ThreeMenu.js"></script>
<script language="javascript" type="text/javascript" src="include/DatePicker/WdatePicker.js"></script>
<title>Target Set</title>
</head>
<style type="text/css">
table tr{
	height:26px;
	line-height:26px;
}
table tr td{
text-align:center;
}
</style>
<body bgcolor="#339966">
</br>
<%
productname = request("productname")
linename = request("linename")
SearchTarget = request("SearchTarget")
TargetSet = request("TargetSet")
RoleGroup = session("ShiftEditRole")
moditarget = request("moditarget")

if TargetSet = "新增" or TargetSet = "修改" then
	tdate = request("tdate")
	Adate = request("Adate")
	OutPCS = request("OutPCS")
	T1 = request("T1")
	C1 = request("C1")
	C2 = request("C2")
	C3 = request("C3")
	C4 = request("C4")
	C5 = request("C5")

	response.write T1 & "</br>"
		' response.write C1 & "</br>"
		' response.write C2 & "</br>"
		' response.write C3 & "</br>"
		' response.write C4 & "</br>"
		' response.write C5 & "</br>"
		' response.write tdate & "</br>"
		' response.write adate & "</br>"
		' response.write session("ShiftEditUser") & "</br>"
		
	if TargetSet = "新增" then

		AddSql = "insert into report_hv_target(productname,linename,adate,aname,udate,uname,tdate,OutPCS,t1,c1,c2,c3,c4,c5)"
		AddSql = AddSql & " values('"&productname&"','"&linename&"',sysdate,'"&session("ShiftEditUser")&"',sysdate,'"&session("ShiftEditUser")&"','"&tdate&"','"&OutPCS&"','"&T1&"','"&C1&"','"&C2&"','"&C3&"','"&C4&"','"&C5&"')"
		'response.write AddSql
		'response.end
		conn.execute(AddSql)
		call sussLoctionHref("添加成功","SetHv.asp?SearchTarget=ok&ProductName="&ProductName&"&LineName="&LineName)
	end if
	
	if TargetSet = "修改" then
		response.write "修改"
		ModiSql = "UPDATE report_hv_target SET udate = sysdate,uname = '"&session("ShiftEditUser")&"',tdate = '"&tdate&"',OutPCS='"&OutPCS&"',T1 = '"&T1&"',C1 = '"&C1&"',C2 = '"&C2&"',C3 = '"&C3&"',C4 = '"&C4&"',C5 = '"&C5&"' WHERE productname = '"&productname&"' and linename = '"&linename&"' and adate = '"&adate&"'"
		'response.write ModiSql
		'response.end
		conn.execute(ModiSql)
		call sussLoctionHref("修改成功","SetHv.asp?SearchTarget=ok&ProductName="&ProductName&"&LineName="&LineName)
	end if
end if
%>
<center>
<form name="SetHv" action="SetPeriphery.asp" >
		<table width="800" border="1" cellspacing="0" cellpadding="0">
			<tr style="height:30px;line-height:30px;">
				<td align="center"><lable>产品：</lable><select name="ProductName" id="TmProduct"></select>
				　线别：<select name="LineName" id="TmLine"></select>
				<select name="CellName" id="TmCell" style="display:none" ></select>
					<script type="text/javascript">
						MenuInit('TmProduct','TmLine','TmCell','<%=ProductName%>','<%=LineName%>','1');
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
	<th>项目名称</th><th>线号</th><th>Target日期</th><th>Output</th><th>Transfer</th><th>CELL 1</th><th>CELL 2</th><th>CELL 3</th><%if ProductName = "RA" or ProductName = "Donau Slim" then%><th>CELL 4</th><th>CELL 5</th><%end if%><th>操作</th>
</tr>
	<form method="post" name="AddForm" action="SetHv.asp">
		<tr style="background:#FFFFFF;">
			<td>
				<%=ProductName%><input type="hidden" name="ProductName" value="<%=ProductName%>" />
				
			</td>
			<td>
				<%=linename%><input type="hidden" name="linename" value="<%=linename%>" />
			</td>
			<td>
				<input type="text" id="tdate" name="tdate" value="<%=tdate%>" size="8"/> <img onClick="WdatePicker({el:'tdate'})" src="include/DatePicker/skin/datePicker.gif" width="16" height="22" align="absmiddle">
			</td>
			<td>
				<input type="text" name="OutPCS" value="0" size="4" />
			</td>
			<td>
				<input type="text" name="T1" value="0" size="4" />
			</td>
			<td>
				<input type="text" name="C1" value="0" size="4" />
			</td>
			<td>
				<input type="text" name="C2" value="0" size="4" />
			</td>
			<td>
				<input type="text" name="C3" value="0" size="4" />
			</td>
            <%if ProductName = "RA" or ProductName = "Donau Slim" then%>
                <td>
                    <input type="text" name="C4" value="0" size="4" />
                </td>
                <td>
                    <input type="text" name="C5" value="0" size="4" />
                </td>
            <%end if%>
			<td>
				<input type="submit" name="TargetSet" value="新增" />
			</td>
		</tr>
	</form>
<%
	LoopNo = 1
	set rs = Server.CreateObject("adodb.recordset")
	sql = "select productname,linename,adate,aname,udate,uname,tdate,OutPCS,T1,C1,C2,C3,C4,C5 from report_hv_target"
	sql = sql & " where productname = '"&productname&"' and linename = '"&linename&"' order by tdate desc"
	'response.write sql
	'response.end
	rs.open sql,conn,1,1
	do while not rs.eof
		ProductName = rs("productname")
		linename = rs("linename")
		uname = rs("uname")
		tdate = rs("tdate")
		Adate = rs("Adate")
		OutPCS = rs("OutPCS")
		T1 = Formatnumber(rs("T1"),2,-1)
		C1 = Formatnumber(rs("C1"),2,-1)
		C2 = Formatnumber(rs("C2"),2,-1)
		C3 = Formatnumber(rs("C3"),2,-1)
		C4 = Formatnumber(rs("C4"),2,-1)
		C5 = Formatnumber(rs("C5"),2,-1)
		
		if (LoopNo mod 2 = 0) then
			GgColor = "#FFFFFF"
		else
			GgColor = "#E6E6E6"
		end if
		%>
		<form method="post" name="ListForm" action="SetPeriphery.asp">
		<tr style="background:<%=GgColor%>;">
			<td>
				<%=ProductName%><input type="hidden" name="ProductName" value="<%=ProductName%>" />
				
			</td>
			<td>
				<%=linename%><input type="hidden" name="linename" value="<%=linename%>" />
				<input type="hidden" name="Adate" value="<%=Adate%>" />
			</td>
			<td>
				<input type="text" name="tdate" id="mdate<%=LoopNo%>" value="<%=tdate%>" size="8"/> <img onClick="WdatePicker({el:'mdate<%=LoopNo%>'})" src="include/DatePicker/skin/datePicker.gif" width="16" height="22" align="absmiddle">
			</td>
			<td>
				<input type="text" name="OutPCS" value="<%=OutPCS%>" size="4" />
			</td>
			<td>
				<input type="text" name="T1" value="<%=T1%>" size="4" />
			</td>
			<td>
				<input type="text" name="C1" value="<%=C1%>" size="4" />
			</td>
			<td>
				<input type="text" name="C2" value="<%=C2%>" size="4" />
			</td>
			<td>
				<input type="text" name="C3" value="<%=C3%>" size="4" />
			</td>
            <%if ProductName = "RA" or ProductName = "Donau Slim" then%>
                <td>
                    <input type="text" name="C4" value="<%=C4%>" size="4" />
                </td>
                <td>
                    <input type="text" name="C5" value="<%=C5%>" size="4" />
                </td>
            <%end if%>
			<td>
				<input type="submit" name="TargetSet" value="修改" />
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


<br/>
<input type="button" value="返回" onClick="window.location.reload('InputOem.asp');" />
　　　　
<input type="button" value="修改密码" onClick="window.location.reload('modify_pass.asp');" />
　　　　
<input type="button" value="退出" onClick="window.location.reload('user_logout.asp');" />
</center>

</body>
</html>
