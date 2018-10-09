<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/Admin/Model/ModelCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Functions/Event_Transaction.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.Form("path")
query=request.Form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
CCL=trim(request("CCL"))
SMALL_PACK=trim(request("SMALL_PACK"))
Box_Size=trim(request("Box_Size"))
CUSTOMER_PN=trim(request("CUSTOMER_PN"))
CUSTOMER_DEFINE=trim(request("CUSTOMER_DEFINE"))
CUSTOMER_LABEL=trim(request("CUSTOMER_LABEL"))
CUSTOMER_DESC=trim(request("CUSTOMER_DESC"))
CUSTOMER_PEGAPN=trim(request("CUSTOMER_PEGAPN"))
CUSTOMER_CONFIG=trim(request("CUSTOMER_CONFIG"))
YNLittleLable=trim(request("YNLittleLable"))
 if SMALL_PACK="" then
 SMALL_PACK=0
 end if
status=request("status")

id=request.Form("id")
modelname=request.Form("modelname")
SQL="select * from PRODUCT_MODEL where ITEM_ID='"&id&"'"

rs.open SQL,conn,1,3
if not rs.eof then
	lead_time=timeconvert(request.Form("lead_time"),request.Form("lead_time_unit"))
	rs("LEAD_TIME")=lead_time
	lead_time2=timeconvert(request.Form("lead_time2"),request.Form("lead_time_unit2"))
	rs("LEAD_TIME2")=lead_time2
	rs("ITEM_STATUS")=request("status")
	rs("Box_Size")=Box_Size
	rs("CUSTOMER_PN")=CUSTOMER_PN
	rs("CUSTOMER_DEFINE")=CUSTOMER_DEFINE
	rs("CUSTOMER_LABEL")=CUSTOMER_LABEL
	rs("SMALL_PACK")=SMALL_PACK
	rs("CCL")=CCL
	rs("CUSTOMER_DESC")=CUSTOMER_DESC
	rs("CUSTOMER_PEGAPN")=CUSTOMER_PEGAPN
	rs("CUSTOMER_CONFIG")=CUSTOMER_CONFIG
	rs("YESNOLITTLELABLE")=YNLittleLable
rs.update
rs.close
word="Successfully edit Model."
action="location.href='"&beforepath&"'"
else
word="Series Model of "&modelname&" has not existed, please input again."
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
<!--#include virtual="/WOCF/BOCF_Close.asp" -->