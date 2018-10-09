<%
response.charset = "UTF-8"
if session("ShiftEditUser") = "" then call sussLoctionHref("Connection timeout or not logged on","user_login.asp")
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
<script language="javascript" type="text/javascript" src="../include/ThreeMenu.js"></script>
<script language="javascript" type="text/javascript" src="include/DatePicker/WdatePicker.js"></script>
<title>Target Set</title>
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
linename = request("linename")
SearchTarget = request("SearchTarget")
TargetSet = request("TargetSet")
RoleGroup = session("ShiftEditRole")
moditarget = request("moditarget")
AYear = request("AYear")

if AYear = "" then AYear = year(now())

if TargetSet = "Add" or TargetSet = "Modify" then
	tdate = request("tdate")
	Adate = request("Adate")
	OutPCS = request("OutPCS")
	T1 = request("T1")
	C1 = request("C1")
	C2 = request("C2")
	C3 = request("C3")
	C4 = request("C4")
	C5 = request("C5")
	OEE = request("OEE")
	
		'response.write T1 & "</br>"
		' response.write C1 & "</br>"
		' response.write C2 & "</br>"
		' response.write C3 & "</br>"
		' response.write C4 & "</br>"
		' response.write C5 & "</br>"
		' response.write tdate & "</br>"
		' response.write adate & "</br>"
		' response.write session("ShiftEditUser") & "</br>"
		
	if TargetSet = "Add" then
		AddSql = "insert into report_hv_target(productname,linename,adate,aname,udate,uname,tdate,OutPCS,t1,c1,c2,c3,c4,c5,OEE)"
		AddSql = AddSql & " values('"&productname&"','"&linename&"',sysdate,'"&session("ShiftEditUser")&"',sysdate,'"&session("ShiftEditUser")&"','"&tdate&"','"&OutPCS&"','"&T1&"','"&C1&"','"&C2&"','"&C3&"','"&C4&"','"&C5&"','"&OEE&"')"
		'response.write AddSql
		'response.end
		conn.execute(AddSql)
		call sussLoctionHref("添加成功","SetHv.asp?SearchTarget=ok&ProductName="&ProductName&"&LineName="&LineName&"&AYear="&AYear)
	end if
	
	if TargetSet = "Modify" then
		'response.write (AYear)
		ModiSql = "UPDATE report_hv_target"
		ModiSql = ModiSql & " SET udate = sysdate,uname = '"&session("ShiftEditUser")&"',OutPCS='"&OutPCS&"',T1 = '"&T1&"',C1 = '"&C1&"',C2 = '"&C2&"',C3 = '"&C3&"',C4 = '"&C4&"',C5 = '"&C5&"',OEE = '"&OEE&"'"
		ModiSql = ModiSql & " WHERE productname = '"&productname&"' and linename = '"&linename&"' and tdate > = '"&tdate&"'"
		'response.write ModiSql
		'response.end
		conn.execute(ModiSql)
		call sussLoctionHref("Modify成功","SetHv.asp?SearchTarget=ok&ProductName="&ProductName&"&LineName="&LineName&"&AYear="&AYear)
	end if
end if
%>
<center>
<form name="SetHv" method="post" action="SetHv.asp" >
		<table width="800" border="1" cellspacing="0" cellpadding="0">
			<tr style="height:30px;line-height:30px;">
				<td align="center"><lable>Product：</lable><select name="ProductName" id="TmProduct"></select>
				　Line：
				  <select name="LineName" id="TmLine"></select>
				<select name="CellName" id="TmCell" style="display:none" ></select>
					<script type="text/javascript">
						MenuInit('TmProduct','TmLine','TmCell','<%=ProductName%>','<%=LineName%>','1');
					</script>
				　Year：
				<input type="text" name="AYear" value="<%=AYear%>" size="6"/>
				　<input type="Submit" name="SearchTarget" value="refurbish" />
				</td>
			</tr>
			</table>
</form>


<%
if SearchTarget = "refurbish" or SearchTarget = "ok" then
%>
<table border="1" cellspacing="0" cellpadding="0" width="800">
<tr style="height:30px;line-height:30px;">
	<th>&nbsp;</th>
	<th>Line</th>
	<th>Target date</th><th>Output</th><th>Transfer</th><th>CELL 1</th><th>CELL 2</th><th>CELL 3</th>
	<%if ProductName = "RA" or ProductName = "Slim" then%>
    	<th>CELL 4</th><th>CELL 5</th>
	<%end if%>
    <th>OEE</th><th>Operation</th>
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
				<input type="text" name="tdate" class="Wdate" id="d122" onFocus="WdatePicker({isShowWeek:true,firstDayOfWeek:1})" value="<%=tdate%>" size="10" />
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
            <%if ProductName = "RA" or ProductName = "Slim" then%>
                <td>
                    <input type="text" name="C4" value="0" size="4" />
                </td>
                <td>
                    <input type="text" name="C5" value="0" size="4" />
                </td>
            <%end if%>
			<td><input type="text" name="OEE" value="0" size="3" /></td>
			<td>
				<input type="submit" name="TargetSet" value="Add" style="background:#FFBFBF;" />
			</td>
		</tr>
	</form>
<%
	LoopNo = 1
	set rs = Server.CreateObject("adodb.recordset")
	sql = "select productname,linename,adate,aname,udate,uname,tdate,OutPCS,nvl(T1,0) T1,NVL(C1,0) C1,NVL(C2,0) C2,NVL(C3,0) C3,NVL(C4,0) C4,NVL(C5,0) C5,OEE from report_hv_target"
	sql = sql & " where productname = '"&productname&"' and linename = '"&linename&"' and to_char(tdate,'YYYY') = '"&AYear&"' order by tdate"
	'response.write sql
	'response.end
	rs.open sql,conn,1,1
	l = 1
	do while not rs.eof
		ProductName = rs("productname")
		linename = rs("linename")
		uname = rs("uname")
		tdate = rs("tdate")
		Adate = rs("Adate")
		OutPCS = rs("OutPCS")
		OEE = rs("OEE")
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
		<form method="post" name="ListForm" action="SetHv.asp">
		<tr style="background:<%=GgColor%>;">
			<td>
				<%=ProductName%><input type="hidden" name="ProductName" value="<%=ProductName%>" />
				
			</td>
			<td>
				<%=linename%><input type="hidden" name="linename" value="<%=linename%>" />
				<input type="hidden" name="Adate" value="<%=Adate%>" /><input type="hidden" name="AYear" value="<%=AYear%>" size="6"/>
			</td>
			<td>
				<input type="text" name="tdate" class="Wdate" id="d122" onFocus="WdatePicker({isShowWeek:true,firstDayOfWeek:1})" value="<%=tdate%>" size="10" />
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
            <%if ProductName = "RA" or ProductName = "Slim" then%>
                <td>
                    <input type="text" name="C4" value="<%=C4%>" size="4" />
                </td>
                <td>
                    <input type="text" name="C5" value="<%=C5%>" size="4" />
                </td>
            <%end if%>
			<td><input type="text" name="OEE" value="<%=OEE%>" size="3" /></td>
			<td>
				<input type="submit" name="TargetSet" value="Modify" style="background:#00FF00;" />
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
<table width="800" border="1" cellspacing="0" cellpadding="0">
	<tr>
		<td colspan="5" align="center">
			<%if session("ShiftEditRole") = "0" then%>
		  <input type="button" value="Set Target for Periphery" onClick="window.location.reload('SetHvPeriphery.asp');" />　　　　
			<%end if%>
			<input type="button" value="change password" onClick="window.location.reload('modify_pass.asp');" />　　　　
			<input type="button" value="Quit" onClick="window.location.reload('user_logout.asp');" />
		</td>
	</tr>
</table>
</center>

</body>
</html>
