<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Series/SeriesCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/ADOVB.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<!--#include virtual="/Functions/NID_SEQ.asp" -->
<%
path=request.Form("path")
query=request.Form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query

SERIESNAME=trim(request.Form("seriesname"))
FACTORY=request.Form("factory")
SECTION_ID=request.Form("section")
LINE_ID=request.Form("line")
INCLUDEMODEL=request.Form("txtIncludeModel")
'response.Write(INCLUDEMODEL)
'response.End()
if request.Form("overall_exception")="1" then
	OVERALL_EXCEPTION="1"
else
	OVERALL_EXCEPTION="0"
end if
TARGET_FIRSTYIELD=request.Form("firstyield")
TARGET_INTERNALYIELD=csng(request.Form("internal_yield"))/100
TARGET_YIELD=request.Form("yield")
BELONGEDSEIRES=request.form("Series")

 
 
set cmd=server.CreateObject("Adodb.Command") 
cmd.ActiveConnection=conn 
cmd.CommandText="ADD_SUB_SERIES"
cmd.CommandType=4
cmd.Parameters.Append cmd.CreateParameter("SERIESNAME", adVarChar, adParamInput, 50, SERIESNAME)
cmd.Parameters.Append cmd.CreateParameter("FACTORY", adVarChar, adParamInput, 10, FACTORY)
cmd.Parameters.Append cmd.CreateParameter("SECTION_ID", adVarChar, adParamInput, 10, SECTION_ID)
cmd.Parameters.Append cmd.CreateParameter("LINE_ID", adVarChar, adParamInput, 10, LINE_ID)
cmd.Parameters.Append cmd.CreateParameter("INCLUDEMODEL", adVarChar, adParamInput, 8000, INCLUDEMODEL)
cmd.Parameters.Append cmd.CreateParameter("OVERALL_EXCEPTION", adVarChar, adParamInput, 10, OVERALL_EXCEPTION)
cmd.Parameters.Append cmd.CreateParameter("TARGET_FIRSTYIELD", adVarChar, adParamInput, 10, TARGET_FIRSTYIELD)
cmd.Parameters.Append cmd.CreateParameter("TARGET_INTERNALYIELD", adVarChar, adParamInput, 10, TARGET_INTERNALYIELD)
cmd.Parameters.Append cmd.CreateParameter("TARGET_YIELD", adVarChar, adParamInput, 10, TARGET_YIELD)
cmd.Parameters.Append cmd.CreateParameter("BELONGEDSEIRES", adVarChar, adParamInput, 10, BELONGEDSEIRES)
cmd.Parameters.Append cmd.CreateParameter("return_order_id", adVarChar, adParamOutput,20)
cmd.execute
return_order_id=cmd("return_order_id")
set cmd=nothing
if err.number=0 then
	word=word &" Successfully add sub series " & SERIESNAME 
	action="location.href='"&beforepath&"'"
else
	word=word&" Fail to add sub series " & SERIESNAME 
	action="history.back()"
end if
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
<!--#include virtual="/WOCF/BOCF_Commit.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" -->