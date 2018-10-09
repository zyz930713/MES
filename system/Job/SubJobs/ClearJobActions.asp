<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Job/JobCheck.asp" -->
<!--#include virtual="/Job/IsDBA.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Functions/GetStation.asp" -->
<!--#include virtual="/Functions/GetRoleMember.asp" -->
<%
jobnumber=request.QueryString("jobnumber")
sheetnumber=request.QueryString("sheetnumber")
jobtype=request.QueryString("jobtype")
SQL="select J.CONTROL_TYPE,J.CONTROL_STATION,J.CONTROL_REASON,J.CONTROL_PERSON,J.CONTROL_TIME from JOB J where J.JOB_NUMBER='"&jobnumber&"' and SHEET_NUMBER='"&sheetnumber&"' and JOB_TYPE='"&jobtype&"'"
rs.open SQL,conn,1,3
if not rs.eof then
rs("CONTROL_TYPE")=null
rs("CONTROL_STATION")=null
rs("CONTROL_REASON")=null
rs("CONTROL_PERSON")=null
rs("CONTROL_TIME")=null
rs.update
end if
rs.close
%>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
<script language="javascript">
alert("Succeessfully to clear actions!")
window.close()
</script>