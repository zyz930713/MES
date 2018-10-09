<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/GetTaskParamHTML.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
id=request.QueryString("id")
SQL="select * from TASK where NID='"&id&"'"
rs.open SQL,conn,1,3
if not rs.eof then
	for pi=1 to 4
	thisoutput=""
	thiscoutput=""
	session("aerror")="DD"&request.QueryString("paramvalue"&pi)
	getTaskParamHTML pi,rs("PARAM_TYPE"&pi),rs("PARAM"&pi),rs("PARAM_CHINESE"&pi),request.QueryString("paramvalue"&pi),thisoutput,thiscoutput
	output=output&"parent.paramhtml"&pi&".innerHTML="""&thisoutput&""";"
	coutput=coutput&"parent.paramhtml"&pi&".innerHTML="""&thiscoutput&""";"
	next
end if
rs.close
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>Untitled Document</title>
<script language="javascript">
<%if session("language")="0" then%>
<%=output%>
<%else%>
<%=coutput%>
<%end if%>
</script>
</head>

<body>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->