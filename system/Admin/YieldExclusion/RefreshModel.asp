<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/POCF_Open.asp" -->
<%
objectname=request.QueryString("objectname")
if request.QueryString("filterstring_notin")<>"" then
	filterstring_notin=request.QueryString("filterstring_notin")
	a_filterstring_notin=split(filterstring_notin,",")
	SQLfilter_notin=""
	for i=0 to ubound(a_filterstring_notin)
	SQLfilter_notin=SQLfilter_notin&" and segment1 not like '"&a_filterstring_notin(i)&"-%'"
	next
end if
if request.QueryString("filterstring_in")<>"" then
	filterstring_in=request.QueryString("filterstring_in")
	a_filterstring_in=split(filterstring_in,",")
	SQLfilter_in=" and ("
	for i=0 to ubound(a_filterstring_in)
		if i=0 then
		SQLfilter_in=SQLfilter_in&" segment1 like '"&a_filterstring_in(i)&"-%'"
		else
		SQLfilter_in=SQLfilter_in&" or segment1 like '"&a_filterstring_in(i)&"-%'"
		end if
	next
	SQLfilter_in=SQLfilter_in&")"
end if
SQL="Select inventory_item_id,organization_id,segment1 from tbl_System_Items where inventory_item_status_code<>'SSPI' "&SQLfilter_notin&SQLfilter_in&session("fromModel")&session("orderModel")
rsPR.open SQL,connPR,1,3
'esponse.Write(SQL)
if not rsPR.eof then
i=1
options=""
while not rsPR.eof
	if rsPR("organization_id")=24 then
	organization_id="C10"
	else
	organization_id="C11"
	end if
	options=options&"var newElem"&i&"=document.createElement('option');newElem"&i&".text='"&rsPR("segment1")&" ("&organization_id&")';newElem"&i&".value='"&rsPR("segment1")&"';parent.form1."&objectname&".options.add(newElem"&i&");"
i=i+1
rsPR.movenext
wend
end if
rsPR.close
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
<!--#include virtual="/WOCF/POCF_Close.asp" -->