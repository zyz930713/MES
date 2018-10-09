<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/System/SystemCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.Form("path")
query=request.Form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
id=request.Form("id")
formname=trim(request.Form("formname"))
SQL="select * from FORM where NID='"&id&"'"
rs.open SQL,conn,3,3
if not rs.eof then
rs("FORM_NAME")=formname
rs("FORM_CHINESE_NAME")=trim(request.Form("formchinesename"))
rs("DESCRIPTION")=trim(request.Form("description"))
rs("APPLY_GROUP")=replace(request.Form("toitem")," ","")
rs("PACKAGE")=trim(request.Form("package"))
rs("PARAM1")=trim(request.Form("param1"))
rs("PARAM_CHINESE1")=trim(request.Form("cparam1"))
rs("PARAM_TYPE1")=trim(request.Form("paramtype1"))
rs("PARAM_SCRIPTS1")=trim(request.Form("paramscripts1"))
rs("PARAM_SHOWBUTTON1")=trim(request.Form("paramshowbutton1"))
rs("PARAM_CHINESE_SHOWBUTTON1")=trim(request.Form("paramchineseshowbutton1"))
rs("PARAM_TITLE1")=trim(request.Form("paramtitle1"))
rs("PARAM_CHINESE_TITLE1")=trim(request.Form("paramchinesetitle1"))
rs("PARAM_BUTTON_SCRIPTS1")=trim(request.Form("parambuttonscripts1"))
rs("PARAM2")=trim(request.Form("param2"))
rs("PARAM_CHINESE2")=trim(request.Form("cparam2"))
rs("PARAM_TYPE2")=trim(request.Form("paramtype2"))
rs("PARAM_SCRIPTS2")=trim(request.Form("paramscripts2"))
rs("PARAM_SHOWBUTTON2")=trim(request.Form("paramshowbutton2"))
rs("PARAM_CHINESE_SHOWBUTTON2")=trim(request.Form("paramchineseshowbutton2"))
rs("PARAM_TITLE2")=trim(request.Form("paramtitle2"))
rs("PARAM_CHINESE_TITLE2")=trim(request.Form("paramchinesetitle2"))
rs("PARAM_BUTTON_SCRIPTS2")=trim(request.Form("parambuttonscripts2"))
rs("PARAM3")=trim(request.Form("param3"))
rs("PARAM_CHINESE3")=trim(request.Form("cparam3"))
rs("PARAM_TYPE3")=trim(request.Form("paramtype3"))
rs("PARAM_SCRIPTS3")=trim(request.Form("paramscripts3"))
rs("PARAM_SHOWBUTTON3")=trim(request.Form("paramshowbutton3"))
rs("PARAM_CHINESE_SHOWBUTTON3")=trim(request.Form("paramchineseshowbutton3"))
rs("PARAM_TITLE3")=trim(request.Form("paramtitle3"))
rs("PARAM_CHINESE_TITLE3")=trim(request.Form("paramchinesetitle3"))
rs("PARAM_BUTTON_SCRIPTS3")=trim(request.Form("parambuttonscripts3"))
rs("PARAM4")=trim(request.Form("param4"))
rs("PARAM_CHINESE4")=trim(request.Form("cparam4"))
rs("PARAM_TYPE4")=trim(request.Form("paramtype4"))
rs("PARAM_SCRIPTS4")=trim(request.Form("paramscripts4"))
rs("PARAM_SHOWBUTTON4")=trim(request.Form("paramshowbutton4"))
rs("PARAM_CHINESE_SHOWBUTTON4")=trim(request.Form("paramchineseshowbutton4"))
rs("PARAM_TITLE4")=trim(request.Form("paramtitle4"))
rs("PARAM_CHINESE_TITLE4")=trim(request.Form("paramchinesetitle4"))
rs("PARAM_BUTTON_SCRIPTS4")=trim(request.Form("parambuttonscripts4"))

dim approve(7)
dim amount(7)

for i=0 to 6
	approve(i)=trim(request.Form("approve" & cstr(i+1))&"")
	
	if trim(request.Form("interval" & cstr(i+1))&"")="" then
		amount(i)=0
	else
		amount(i)=trim(request.Form("interval" & cstr(i+1))&"")
	end if
next

rs("APPROVE1")=approve(0)
rs("AMOUNT1")=0
rs("APPROVE2")=approve(1)
rs("AMOUNT2")=amount(0)
rs("APPROVE3")=approve(2)
rs("AMOUNT3")=amount(1)
rs("APPROVE4")=approve(3)
rs("AMOUNT4")=amount(2)
rs("APPROVE5")=approve(4)
rs("AMOUNT5")=amount(3)
rs("APPROVE6")=approve(5)
rs("AMOUNT6")=amount(4)
rs("APPROVE7")=approve(6)
rs("AMOUNT7")=amount(5)

if request.Form("alert_time")<>"" then
rs("ALERT_TIME")=request.Form("alert_time")
end if
rs("ALERT_PERSON")=request.Form("alert_person")
rs("ACT_PERSON")=request.Form("actperson")
rs.update
word="Successfully edit Form."
action="location.href='"&beforepath&"'"
else
word="Form of "&formname&" has existed, please input again."
action="history.back()"
end if
rs.close
%>
<html>
<head>
<title>Update</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<script language="JavaScript">
alert("<%=word%>");
<%=action%>;
</script>
<body>

</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->