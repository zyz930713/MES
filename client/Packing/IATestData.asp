<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/WOCF/IA_Open.asp" -->

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
'	sql="select a.ad_id, linename, measuredatetime, adfail, error_name, cerror_name, pvs.func_gethohd(cerror_name) as hold, "
'	sql=sql+" measurementpcname, preassemblycode,b.serialnumber "
'	sql=sql+" from pvs.vw_adid_by_sn a,pvs.ad_serial b where a.ad_id=b.ad_id and b.serialnumber='"&code&"' "
'	sql=sql+" order by a.measuredatetime desc "
' sql="select summary_id,serialNumber,testTime,jobNumber,pcName,product,testResult,failName,TEST_ENVIRONMENT from  TEST_RESULT_SUMMARY a left join TEST_RESULT_MAIN_KEB b on(a.archive_id=b.ARCHIVE_ID) where serialnumber = '"&code&"' " 
	'sql="select summary_id,serialNumber,testTime,jobNumber,pcName,product,testResult,failName,TEST_ENVIRONMENT from  TEST_RESULT_SUMMARY a,TEST_RESULT_MAIN_KEB b where a.archive_id=b.ARCHIVE_ID and serialnumber = '"&code&"'"
  ' sql="select  * from package a ,serial_index b where a.serial_id=b.serial_id and serialnumber = '"&code&"'" 
   sql="select [Pr_ID],[serialNumber],[testDateTime],[upDateTime],[pr_fail],[errorName],'INLINE2ND' as testitem from [TSD_EXPLORER].[tsd].[IA_List] where  serialnumber = '"&code&"'" 
   ' sql="select  a.serial_id,b.serialnumber,a.failname,a.jobnumber,a.pcname,a.testresult,a.testitem,a.testtime from package a ,serial_index b where a.serial_id=b.serial_id and serialnumber = '"&code&"'"
'response.Write(sql)
	rsIA.open sql,connIA,1,3

	
end if
%>
<table width="100%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
	<tr>
        <td colspan="7" align="center" class="t-t-DarkBlue">TSD Test Data TSD��������</td>
    </tr>
	<tr align="center">
		<td >NO ����</td>
		<td >2D Code ��ά��</td>
		<td >Test ID</td>
		<td >Result ���</td>
		<td >Defect ȱ��</td>		
		<td >Test Name ��������</td>
		<td >Test Time ����ʱ��</td>
	</tr>
<%	
if isQuery then
i=1
while not rsIA.eof
	result="Pass"
	trColor="#00CC00"
	defect="&nbsp;"
	criterion="&nbsp;"
	if rsIA("pr_fail")="1" then
		result="Fail"
		defect=rsIA("errorName")
		'criterion=rsIA("cerror_name")
		'if rsIA("hold")="0" then
		''	trColor="#FFFF00"
		'else
			trColor="#FF0000"
		'end if
	end if
%>
	<tr align="center" bgcolor="<%=trColor%>" >
		<td><%=i%></td>
		<td><%=rsIA("serialNumber")%></td>
		<td><%=rsIA("Pr_ID")%></td>
		<td><%=result%></td>
		<td><%=defect%></td>
		<td><%=rsIA("testitem")%></td>
		<td><%=rsIA("testDateTime")%>&nbsp;</td>
		
	</tr>	
<%
i=i+1
rsIA.movenext		
wend
end if
%>
</table>
</body>
</html>
<!--#include virtual="/WOCF/IA_close.asp" -->