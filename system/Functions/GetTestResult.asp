<%
function getTestResult(jobnumber,test_defectcode_id)	
set rsS=server.CreateObject("adodb.recordset")
SQLS="Select 1,J.TEST_DEFECTCODE_VALUE,T.VALUE_TYPE,T.SCALE from JOB_TEST_RESULT J inner join TEST_DEFECTCODE T on J.TEST_DEFECTCODE_ID=T.NID where J.JOB_NUMBER='"&jobnumber&"' and J.TEST_DEFECTCODE_ID='"&test_defectcode_id&"'"
rsS.open SQLS,conn,1,3
if not rsS.eof then
	if rsS("VALUE_TYPE")="1" then
	getTestResult=rsS("TEST_DEFECTCODE_VALUE")
	elseif rsS("VALUE_TYPE")="2" then
	getTestResult=formatnumber(csng(rsS("TEST_DEFECTCODE_VALUE")),rsS("SCALE"),-1)
	else
	getTestResult=""
	end if
else
getTestResult=""
end if
rsS.close
set rsS=nothing
end function
%>