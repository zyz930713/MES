<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
objectname=request.QueryString("objectname")
if request.QueryString("filterstring_notin")<>"" then
	filterstring_notin=ucase(request.QueryString("filterstring_notin"))
	a_filterstring_notin=split(filterstring_notin,",")
	SQLfilter_notin=""
	for i=0 to ubound(a_filterstring_notin)
	SQLfilter_notin=SQLfilter_notin&" and lower(PM.ITEM_NAME) not like '%"&ucase(a_filterstring_notin(i))&"%'"
	next
end if
if request.QueryString("filterstring_in")<>"" then
	filterstring_in=ucase(request.QueryString("filterstring_in"))
	a_filterstring_in=split(filterstring_in,",")
	SQLfilter_in=" and ("
	for i=0 to ubound(a_filterstring_in)
		if i=0 then
		SQLfilter_in=SQLfilter_in&" lower(PM.ITEM_NAME) like '%"&ucase(a_filterstring_in(i))&"%'"
		else
		SQLfilter_in=SQLfilter_in&" or lower(PM.ITEM_NAME) like '%"&ucase(a_filterstring_in(i))&"%'"
		end if
	next
	SQLfilter_in=SQLfilter_in&")"
end if
if request.QueryString("filter_forcedin")="1" then
SQL="select PM.ITEM_ID,PM.ITEM_NAME,PM.ITEM_TYPE,PM.ITEM_STATUS,F.FACTORY_NAME from PRODUCT_MODEL PM left join FACTORY F on PM.FACTORY_ID=F.NID where PM.ITEM_NAME is not null "&SQLfilter_notin&SQLfilter_in&session("orderModel")
else
SQL="select PM.ITEM_ID,PM.ITEM_NAME,PM.ITEM_TYPE,PM.ITEM_STATUS,F.FACTORY_NAME from PRODUCT_MODEL PM left join FACTORY F on PM.FACTORY_ID=F.NID where PM.ITEM_NAME is not null "&SQLfilter_notin&SQLfilter_in&session("fromModel")&session("orderModel")
end if
'response.Write(request.QueryString("filter_forcedin"))
response.Write(SQL)
'response.End()
rs.open SQL,conn,1,3
if not rs.eof then
i=1
options=""
while not rs.eof
	options=options&"var newElem"&i&"=document.createElement('option');newElem"&i&".text='"&rs("ITEM_NAME")&" ("&rs("FACTORY_NAME")&")';newElem"&i&".value='"&rs("ITEM_ID")&"';parent.form1."&objectname&".options.add(newElem"&i&");"
i=i+1
rs.movenext
wend
end if
rs.close
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charsPRet=gb2312" />
<title>Untitled Document</title>
<script language="JavaScript" src="/Functions/FromTo.js" type="text/javascript"></script>
<script language="javascript" type="text/javascript">
<%="do{parent.form1."&objectname&".options[0]=null;}while(parent.form1."&objectname&".length>0);"%>
<%=options%>
<%="parentdeselectedcount();"%>
</script>

</head>

<body>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->