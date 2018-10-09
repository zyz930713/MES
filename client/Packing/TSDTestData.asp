<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/WOCF/TSD_Open.asp" -->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=application("SystemName")%></title>
</head>
<link href="/CSS/GeneralKES1.css" rel="stylesheet" type="text/css">

<body bgcolor="#339966">
<%
PRODUCT=request("PRODUCT")
code=request("txt_2d_code")
isQuery=false
if code <> "" then
	isQuery=true
'	sql="select a.ad_id, linename, measuredatetime, adfail, error_name, cerror_name, pvs.func_gethohd(cerror_name) as hold, "
'	sql=sql+" measurementpcname, preassemblycode,b.serialnumber "
'	sql=sql+" from pvs.vw_adid_by_sn a,pvs.ad_serial b where a.ad_id=b.ad_id and b.serialnumber='"&code&"' "
'	sql=sql+" order by a.measuredatetime desc "
' sql="select summary_id,serialNumber,testTime,jobNumber,pcName,product,testResult,failName,TEST_ENVIRONMENT from  TEST_RESULT_SUMMARY a left join TEST_RESULT_MAIN_KEB b on(a.archive_id=b.ARCHIVE_ID) where serialnumber = '"&code&"' " 
	'sql="select summary_id,serialNumber,testTime,jobNumber,pcName,product,testResult,failName,TEST_ENVIRONMENT from  TEST_RESULT_SUMMARY a,TEST_RESULT_MAIN_KEB b where a.archive_id=b.ARCHIVE_ID and serialnumber = '"&code&"'"
   if PRODUCT="Explorer" then
   sql="select * from Vw_Packing where serialnumber = '"&code&"'" 
   else
   sql="select  * from package a ,serial_index b where a.serial_id=b.serial_id and serialnumber = '"&code&"'" 
   end if
   ' sql="select  a.serial_id,b.serialnumber,a.failname,a.jobnumber,a.pcname,a.testresult,a.testitem,a.testtime from package a ,serial_index b where a.serial_id=b.serial_id and serialnumber = '"&code&"'"
	rsTSD.open sql,connTSD,1,3
	
end if
%>
<table width="100%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
	<tr>
        <td colspan="9" align="center" class="t-t-DarkBlue">TSD Test Data TSD测试数据</td>
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
while not rsTSD.eof
	result="Pass"
	trColor="#00CC00"
	defect="&nbsp;"
	criterion="&nbsp;"
	if rsTSD("testResult")="FAIL" then
		result="Fail"
		defect=rsTSD("failName")
		'criterion=rsTSD("cerror_name")
		'if rsTSD("hold")="0" then
		''	trColor="#FFFF00"
		'else
			trColor="#FF0000"
		'end if
	end if
%>
	<tr align="center" bgcolor="<%=trColor%>" >
		<td><%=i%></td>
		<td><%=rsTSD("serialNumber")%></td>
		<td><%=rsTSD("Resultid")%></td>
		<td><%=result%></td>
		<td><%=defect%></td>
		<td><%=criterion%></td>
		<td><%=rsTSD("testitem")%>&nbsp;</td>
		<td><%=rsTSD("PCNAME")%>&nbsp;</td>
		<td><%=rsTSD("testTime")%>&nbsp;</td>
	</tr>	
<%
i=i+1
rsTSD.movenext		
wend
end if
%>
</table>
</body>
</html>
<!--#include virtual="/WOCF/TSD_Close.asp" -->