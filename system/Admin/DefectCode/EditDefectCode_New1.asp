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
id=request.Form("id")
defectcode=request.Form("defectcode")
SQL="select * from DEFECTCODE_New where NID='"&id&"'"
rs.open SQL,conn,1,3
if not rs.eof then
	rs("DEFECT_CODE")=defectcode
	rs("DEFECT_NAME")=trim(request.Form("defectname"))
	rs("DEFECT_CHINESE_NAME")=trim(request.Form("chinesename"))
	rs("FACTORY_ID")=request.Form("factory")
	rs("IS_DELETE")=0
	rs("LASTUPDATE_PERSON")=session("code")
	rs("LASTUPDATE_TIME")=date()
	rs.update
	word="Successfully edit New Defect Code."
	action="location.href='"&beforepath&"'"
else
	word="Defect Code of "&defectcode&" has not existed, please input again."
	action="history.back()"
end if
rs.close
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