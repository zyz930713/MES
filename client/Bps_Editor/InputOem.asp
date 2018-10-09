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
<script type="text/javascript">
function $(obj){
   return document.getElementById(obj);
}
function jscf(TypeNr,cellno){
	var cg = TypeNr + 'In' + cellno;
	var cb = TypeNr + 'Out' + cellno;
	var cf = TypeNr + 'Real' + cellno;
	var t1g = parseFloat($(cg).value);
	var t1b = parseFloat($(cb).value);
	t1 = t1b / t1g;
	t1 = t1 * 100;
	t1 = t1.toFixed(2);
	$(cf).value=t1;
}
</script>
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
	height:30px;
	line-height:26px;
	background:#F2F2F2;
	}
table th{
	border-color:black;
}
table td{
	font:normal 14px/16px "Arial";
	border-color:black;
}
tr.change:hover
{
	background-color:#D1D1D1
}
</style>
</head>
<body bgcolor="#339966">
<%
EditorName = session("ShiftEditUser")
RoleGroup = session("ShiftEditRole")
if RoleGroup = "0" or RoleGroup = "1" then RoleGroup = "%"
'response.write EditorName

HDate = request("HDate")
if HDate = "" then
	HDate = FormatTime(date(),2)
end if
AddMod = request("AddMod")
deldata = request("deldata")

' response.write "修改限制:" & date()
' response.write "<br/>"
' response.write "设定时间:" & HDate
' response.write "<br/>"

if AddMod = "添加" or AddMod = "修改" or AddMod = "删除" then

	productname = request("productname")
	linename = request("linename")
	PcsTarget = request("PcsTarget")
	PcsReal = request("PcsReal")
	Rtowh = request("Rtowh")

	ProcessIn = request("ProcessIn")
	ProcessOut = request("ProcessOut")
	ProcessTarget = request("ProcessTarget")
	ProcessReal = request("ProcessReal")

	AcousticIn = request("AcousticIn")
	AcousticOut = request("AcousticOut")
	AcousticTarget = request("AcousticTarget")
	AcousticReal = request("AcousticReal")

	CosmeticIn = request("CosmeticIn")
	CosmeticOut = request("CosmeticOut")
	CosmeticTarget = request("CosmeticTarget")
	CosmeticReal = request("CosmeticReal")
	TtyTarget = ProcessTarget * AcousticTarget * CosmeticTarget / 10000
	TtyReal = ProcessReal * AcousticReal * CosmeticReal / 10000
	if AddMod = "添加" then
		SqlAdd = "INSERT INTO REPORT_OEM_DATA(product,line_num,create_time,up_date,line_type,TARGET,ACTUAL,REAL_TO_WH,PQ_START,PQ_OUT,pq_target,pq_yield,ams_start,ams_good,ams_target,ams_yield,cosmetic_start,cosmetic_good,cosmetic_target,cosmetic_yield,tty_close,tty_open,tty_target,tty_yield)"
		SqlAdd = SqlAdd & " VALUES('"&productname&"','"&linename&"','"&cdate(HDate)&"',sysdate,'OEM','"&PcsTarget&"','"&PcsReal&"','"&Rtowh&"','"&ProcessIn&"','"&ProcessOut&"','"&ProcessTarget&"','"&ProcessReal&"','"&AcousticIn&"','"&AcousticOut&"','"&AcousticTarget&"','"&AcousticReal&"','"&CosmeticIn&"','"&CosmeticOut&"','"&CosmeticTarget&"','"&CosmeticReal&"','"&CosmeticOut&"','"&ProcessIn&"','"&TtyTarget&"','"&TtyReal&"')" 
		'response.write SqlAdd
		'response.end
		conn.execute(SqlAdd)
		call sussLoctionHref("添加成功","inputOem.asp?HDate="&cdate(HDate))
	end if
		
	if AddMod = "修改" then
		SqlMod = "UPDATE REPORT_OEM_DATA"
		SqlMod = SqlMod & " set up_date = sysdate,target = '"&PcsTarget&"',actual = '"&PcsReal&"',REAL_TO_WH = '"&Rtowh&"',pq_start = '"&ProcessIn&"',pq_out = '"&ProcessOut&"',pq_target = '"&ProcessTarget&"',pq_yield = '"&ProcessReal&"',ams_start = '"&AcousticIn&"',ams_good = '"&AcousticOut&"',ams_target = '"&AcousticTarget&"',ams_yield = '"&AcousticReal&"',cosmetic_start = '"&CosmeticIn&"',cosmetic_good = '"&CosmeticOut&"',cosmetic_target = '"&CosmeticTarget&"',cosmetic_yield = '"&CosmeticReal&"',tty_close = '"&CosmeticOut&"',tty_open = '"&ProcessIn&"',tty_target = '"&TtyTarget&"',tty_yield = '"&TtyReal&"'"
		SqlMod = SqlMod & " where line_type = 'OEM' and product = '"&productname&"' and line_num = '"&linename&"' and create_time = '"&cdate(HDate)&"'"
		'response.write SqlMod
		'response.end
		conn.execute(SqlMod)
		call sussLoctionHref("修改成功","inputOem.asp?HDate="&cdate(HDate))
	end if

	if AddMod = "删除" then
		SqlDel = "DELETE FROM REPORT_OEM_DATA where line_type = 'OEM' and product = '"&productname&"' and line_num = '"&linename&"' and create_time = '"&cdate(HDate)&"'"
		'response.write SqlDel
		conn.execute(SqlDel)
		call sussLoctionHref("删除成功","inputOem.asp?HDate="&cdate(HDate))
	end if
end if

%>

</br>
<center>

<form method="post" name="form1" action="inputOem.asp">
日期：<input type="text" name="HDate" id="HDate" value="<%=HDate%>" size="7" onClick="WdatePicker()" OnChange="this.form.submit()" />
	<input type="submit" name="Getdata" value="设定日期" />
</form>
<br/>
<table border="1" cellspacing="0" cellpadding="0" width="1200">
	<tr>
		<td colspan="18" class="TitleLable"><p style="height:34px;line-height:34px;"><strong>Details Beijing (OEM)：<%=FormatTime(HDate,4)%></strong></p></td>
    </tr>
    <tr>
		<td rowspan="2"><p align="center"><strong>Product</strong></p></td>
		<td rowspan="2"><p align="center"><strong>Lines</strong></p></td>
		<td rowspan="2"><p align="center"><strong>Target<br/>[pcs]</strong></p></td>
		<td rowspan="2"><p align="center"><strong>Actual<br/>[pcs]</strong></p></td>
		<td rowspan="2"><p align="center"><strong>Real to<br/>W/H</strong></p></td>
		<td colspan="4"><p align="center"><strong>Process Yield</strong></p></td>
		<td colspan="4"><p align="center"><strong>Acoustic Yield</strong></p></td>
		<td colspan="4"><p align="center"><strong>Cosmetic Yield</strong></p></td>
		<td rowspan="2"><p align="center"><strong>操作</strong></p></td>
    </tr>
    <tr>
		<td><p align="center"><strong>Input</strong></p></td>
		<td><p align="center"><strong>Output</strong></p></td>
		<td><p align="center"><strong>Target</strong></p></td>
		<td><p align="center"><strong>Real</strong></p></td>
		<td><p align="center"><strong>Input</strong></p></td>
		<td><p align="center"><strong>Output</strong></p></td>
		<td><p align="center"><strong>Target</strong></p></td>
		<td><p align="center"><strong>Real</strong></p></td>
		<td><p align="center"><strong>Input</strong></p></td>
		<td><p align="center"><strong>Output</strong></p></td>
		<td><p align="center"><strong>Target</strong></p></td>
		<td><p align="center"><strong>Real</strong></p></td>
    </tr>
<%
LoopNo = 1
set rs = Server.CreateObject("adodb.recordset")
sql = "select OemProductId,productname,linename,pcstarget,processtarget,acoustictarget,cosmetictarget from report_oem_target where isactive = '1' and Egroup like '"&RoleGroup&"' ORDER BY productname,linename"
rs.open sql,conn,1,1
do while not rs.eof
OemProductId = rs("OemProductId")
productname = rs("productname")
linename = rs("linename")
PcsTarget = rs("PcsTarget")
ProcessTarget = rs("ProcessTarget")
AcousticTarget = rs("AcousticTarget")
CosmeticTarget = rs("CosmeticTarget")

PcsReal = 0
Rtowh = 0
ProcessIn = 0
ProcessOut = 0
ProcessReal = 0
AcousticIn = 0
AcousticOut = 0
AcousticReal = 0
CosmeticIn = 0
CosmeticOut = 0
CosmeticReal = 0
DelSubmit = ""
AMStr = "<input type='submit' name='AddMod' value='添加' style='background:#FFBFBF;'/>"
DelSub = 0
	set rs2 = Server.CreateObject("adodb.recordset")
	sql2 = "select product,line_num,create_time,line_type,TARGET,ACTUAL,REAL_TO_WH,PQ_START,PQ_OUT,pq_target,pq_yield,ams_start,ams_good,ams_target,ams_yield,cosmetic_start,cosmetic_good,cosmetic_target,cosmetic_yield,tty_close,tty_open,tty_target,tty_yield from REPORT_OEM_DATA where line_type = 'OEM' and product = '"&productname&"' and line_num = '"&linename&"' and create_time = '"&cdate(HDate)&"'"
	'response.write sql2
	rs2.open sql2,conn,1,1
	'response.write rs2("ACTUAL")
	if not rs2.eof then
		PcsTarget = rs2("TARGET")
		PcsReal = rs2("ACTUAL")
		Rtowh = rs2("REAL_TO_WH")
		
		ProcessIn = rs2("PQ_START")
		ProcessOut = rs2("PQ_OUT")
		ProcessTarget = rs2("PQ_TARGET")
		ProcessReal = rs2("PQ_YIELD")
		
		AcousticIn = rs2("AMS_START")
		AcousticOut = rs2("AMS_GOOD")
		AcousticTarget = rs2("AMS_TARGET")
		AcousticReal = rs2("AMS_YIELD")
		
		CosmeticIn = rs2("COSMETIC_START")
		CosmeticOut = rs2("COSMETIC_GOOD")
		CosmeticTarget = rs2("COSMETIC_TARGET")
		CosmeticReal = rs2("COSMETIC_YIELD")
		AMStr = "<input type='submit' name='AddMod' value='修改' style='background:#BFFFBF;'/>"
		DelSub = 1
	end if
	rs2.close
	
' response.write cdate(HDate) & "<br/>"
' response.write date() & "<br/>"
' response.write time() & "<br/>"
' response.write DelSub & "<br/>"

%>
<form method="post" name="form<%=OemProductId%>" action="inputOem.asp">
    <tr class="change">
		<td align="left"><input type="hidden" name="HDate" value="<%=HDate%>" size="6" /><input type="hidden" name="productname" value="<%=productname%>" size="1" /><input type="hidden" name="linename" value="<%=linename%>" size="1" /><%=productname%></td>
		<td align="center"><%=linename%></td>
		<td align="center"><input type="text" name="PcsTarget" value="<%=PcsTarget%>" size="4" style="background:#CCCCCC;" /></td>
		<td align="center"><input type="text" name="PcsReal" value="<%=PcsReal%>" size="4" /></td>
		<td align="center"><input type="text" name="Rtowh" value="<%=Rtowh%>" size="4" /></td>
	  
		<td align="center"><input type="text" name="ProcessIn" value="<%=ProcessIn%>" id="ProcessIn<%=OemProductId%>" size="4" OnChange="jscf('Process',<%=OemProductId%>)" /></td>
		<td align="center"><input type="text" name="ProcessOut" value="<%=ProcessOut%>" id="ProcessOut<%=OemProductId%>" size="4" OnChange="jscf('Process',<%=OemProductId%>)" /></td>
		<td align="center"><input type="text" name="ProcessTarget" value="<%=ProcessTarget%>" size="4" style="background:#CCCCCC;" /></td>
		<td align="center"><input type="text" name="ProcessReal" value="<%=ProcessReal%>" id="ProcessReal<%=OemProductId%>" size="4" /></td>
	  
		<td align="center"><input type="text" name="AcousticIn" value="<%=AcousticIn%>" id="AcousticIn<%=OemProductId%>" size="4" OnChange="jscf('Acoustic',<%=OemProductId%>)" /></td>
		<td align="center"><input type="text" name="AcousticOut" value="<%=AcousticOut%>" id="AcousticOut<%=OemProductId%>" size="4" OnChange="jscf('Acoustic',<%=OemProductId%>)" /></td>
		<td align="center"><input type="text" name="AcousticTarget" value="<%=AcousticTarget%>" size="4" style="background:#CCCCCC;" /></td>
		<td align="center"><input type="text" name="AcousticReal" value="<%=AcousticReal%>" id="AcousticReal<%=OemProductId%>" size="4" /></td>
	  
		<td align="center"><input type="text" name="CosmeticIn" value="<%=CosmeticIn%>" id="CosmeticIn<%=OemProductId%>" size="4" OnChange="jscf('Cosmetic',<%=OemProductId%>)" /></td>
		<td align="center"><input type="text" name="CosmeticOut" value="<%=CosmeticOut%>" id="CosmeticOut<%=OemProductId%>" size="4" OnChange="jscf('Cosmetic',<%=OemProductId%>)" /></td>
		<td align="center"><input type="text" name="CosmeticTarget" value="<%=CosmeticTarget%>" size="4" style="background:#CCCCCC;" /></td>
		<td align="center"><input type="text" name="CosmeticReal" value="<%=CosmeticReal%>" id="CosmeticReal<%=OemProductId%>" size="4" /></td>
		<td align="center"><%=AMStr%>&nbsp;<%if cdate(HDate) = date() and time() < cdate("08:40:00") and DelSub = 1 then%><input type="submit" name="AddMod" value="删除" style="background:Red;" onClick="return confirm('确定进行删除操作吗？')"/><%end if%></td>
    </tr>
</form>
<%
rs.movenext
loop
%>
</table>
<br/>
<input type="button" value="Set 设置" onClick="window.location.reload('SetOem.asp');" />　　　　
<input type="button" value="Close 退出" onClick="window.location.reload('user_logout.asp');">
</center>

</body>
</html>
