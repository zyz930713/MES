<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Job/Scrap/ScrapCheck.asp" -->
<!--#include virtual="/WOCF/ADOVB.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.Form("path")
query=request.Form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
trans=false
for i=1 to request.Form("idcount")
	if request.Form("id"&i)="1" then
	trans=true
	exit for
	end if
next
if trans=true then
%>
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<%
end if
words=""
j=1
for i=1 to request.Form("idcount")
	if request.Form("id"&i)="1" then
		jobnumber=trim(request.Form("jobnumber"&i))
		set cmd=server.CreateObject("Adodb.Command") 
		cmd.ActiveConnection=conn 
		cmd.CommandText="JOB_MASTER_SINGLE_UPDATE$"
		cmd.CommandType=4 
		'response.Write(request.Form("idcount"))
		'response.End()
		cmd.Parameters.Append cmd.CreateParameter("v_job_number", adVarChar, adParamInput, 10, jobnumber)
		cmd.execute
		set cmd=nothing
		if err.number=0 then
		word=word&"\n"&jobnumber&"更新成功！"
		else
		word=word&"\n"&jobnumber&"更新失败！"
		end if
	end if
next
action="location.href='"&beforepath&"'"
if trans=true then
%>
<!--#include virtual="/WOCF/BOCF_Commit.asp" -->
<%end if%>
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