<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/WOCF/PVS_Open.asp" -->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=application("SystemName")%></title>
</head>
<link href="/CSS/GeneralKES1.css" rel="stylesheet" type="text/css">

<body bgcolor="#339966">
<%
code=request("txt_2d_code")
isQuery=false
if code <> "" then
	isQuery=true
	sql="select a.ad_id, linename, measuredatetime, adfail, error_name, cerror_name, pvs.func_gethohd(cerror_name) as hold, "
	sql=sql+" measurementpcname, preassemblycode,b.serialnumber "
	sql=sql+" from pvs.vw_adid_by_sn a,pvs.ad_serial b where a.ad_id=b.ad_id and b.serialnumber='"&code&"' "
	sql=sql+" order by a.measuredatetime desc "
	rsPVS.open sql,connPVS,1,3
	
end if
%>
<table width="100%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
	<tr>
        <td colspan="9" class="t-t-DarkBlue" align="center">PVS Test Data PVS测试数据</td>
    </tr>
	<tr align="center">
		<td >NO 序列</td>
		<td >2D Code 二维码</td>
		<td >Test ID</td>
		<td >Result 结果</td>
		<td >Defect 缺陷</td>
		<td >Criterion 标准</td>
		<td >Test Name 测试名称</td>
		<td >Machine 测试机</td>
		<td >Test Time 测试时间</td>
	</tr>
<%	
if isQuery then
i=1
while not rsPVS.eof
	result="Pass"
	trColor="#00CC00"
	defect="&nbsp;"
	criterion="&nbsp;"
	if rsPVS("adfail")="True" then
		result="Fail"
		defect=rsPVS("error_name")
		criterion=rsPVS("cerror_name")
		if rsPVS("hold")="0" then
			trColor="#FFFF00"
		else
			trColor="#FF0000"
		end if
	end if
%>
	<tr align="center" bgcolor="<%=trColor%>" >
		<td><%=i%></td>
		<td><%=code%></td>
		<td><%=rsPVS("ad_id")%></td>
		<td><%=result%></td>
		<td><%=defect%></td>
		<td><%=criterion%></td>
		<td><%=rsPVS("linename")%>&nbsp;</td>
		<td><%=rsPVS("measurementpcname")%>&nbsp;</td>
		<td><%=rsPVS("measuredatetime")%>&nbsp;</td>
	</tr>	
<%
i=i+1
rsPVS.movenext		
wend
end if
%>
</table>
</body>
</html>
<!--#include virtual="/WOCF/PVS_Close.asp" -->