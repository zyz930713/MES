<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/Job/JobCheck.asp" -->
<!--#include virtual="/Job/IsDBA.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/ADOVB.asp" -->
<%
p_nid=request.QueryString("p_nid")
set cmd=server.CreateObject("Adodb.Command") 
cmd.ActiveConnection=conn 
cmd.CommandText="KE_OBI_KES1_UTIL.UPDATE_JOB_MASTER_STORE_CHG"
cmd.CommandType=4
cmd.Parameters.Append cmd.CreateParameter("p_nid", adVarChar, adParamInput, 50, p_nid)
cmd.Parameters.Append cmd.CreateParameter("p_txn_id", adNumeric, adParamInput, 50)
cmd.Parameters.Append cmd.CreateParameter("p_message_status", adVarChar, adParamInput, 50, "")
cmd.Parameters.Append cmd.CreateParameter("p_erp_job_close_status", adVarChar, adParamInput, 50, "")
cmd.Parameters.Append cmd.CreateParameter("p_error_code", adVarChar, adParamInput, 50, "")
cmd.Parameters.Append cmd.CreateParameter("p_error_explanation", adVarChar, adParamInput, 200)
cmd.Parameters.Append cmd.CreateParameter("p_allow_erp_resubmit", adVarChar, adParamInput, 10, "N")
cmd.execute
set cmd=nothing
word="改变提交状态为无效！"
%>
<script language="javascript">
alert("<%=word%>");
location.href='<%=beforepath%>';
</script>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
</head>

<body>

</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->