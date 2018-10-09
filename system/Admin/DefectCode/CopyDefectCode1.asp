<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/DefectCode/DefectCodeCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/NID_SEQ.asp" -->
<%
path=request.Form("path")
query=request.Form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
defectcode=request.Form("defectcode")
station=replace(request.Form("station")," ","")
if station<>"" then
astation=split(station,",")
	for i=0 to ubound(astation)
		SQL="select * from DEFECTCODE where DEFECT_CODE='"&defectcode&"' and STATION_ID='"&astation(i)&"'"
		rs.open SQL,conn,1,3
		if rs.eof then
		rs.addnew
		rs("NID")="DF"&NID_SEQ("DEFECTCODE")
		rs("DEFECT_CODE")=defectcode
		rs("DEFECT_NAME")=request.Form("defectname")
		rs("DEFECT_CHINESE_NAME")=trim(request.Form("chinesename"))
		rs("FACTORY_ID")=request.Form("factory")
		rs("MATERIAL_ID")=replace(request.Form("toitem")," ","")
		rs("APPLIED_PART_ID")=request.Form("part")
		rs("STATION_ID")=astation(i)
		rs.update
		word="Successfully save New Defect Code."
		action="location.href='"&beforepath&"'"
		else
		word="Defect Code of "&defectcode&" has existed, please input again."
		action="history.back()"
		exit for
		end if
		rs.close
	next
else
word="No New Defect Code saved."
action="location.href='"&beforepath&"'"
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