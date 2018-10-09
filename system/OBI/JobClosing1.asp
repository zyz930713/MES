<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/ADOVB.asp" -->
<%
idcount=request.Form("idcount")
OK=0
Fail=0
for i=1 to idcount
	if request.Form("id"&i)="1" then
		set cmd=server.CreateObject("Adodb.Command") 
		cmd.ActiveConnection=conn 
		cmd.CommandText="KE_OBI_KES1_UTIL.submit_erp_job_complete"
		cmd.CommandType=4
		cmd.Parameters.Append cmd.CreateParameter("p_table_type", adInteger, adParamInput, 50, 1)
		cmd.Parameters.Append cmd.CreateParameter("p_nid", adVarChar, adParamInput, 50, request.Form("nid"&i))
		cmd.Parameters.Append cmd.CreateParameter("x_return_status", adVarChar, adParamOutput, 200)
		cmd.Parameters.Append cmd.CreateParameter("x_return_error", adVarChar, adParamOutput, 200)
		cmd.execute
		if cmd("x_return_status")="SUCCESS" then
		OK=OK+1
		else
		Fail=Fail+1
		FailNumber=FailNumber&request.Form("nid"&i)&" ("&cmd("x_return_error")&"),"
		end if
		set cmd=nothing
	end if
next
word="成功提交"&OK&"；失败"&Fail&"；失败工单是："&FailNumber
action="location.href='"&beforepath&"'"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
</head>
<script language="javascript">
alert("<%=word%>");
<%=action%>;
</script>
<body>

</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->