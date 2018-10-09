<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/GetFormParamHTML.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
id=request.QueryString("id")
jobnumber=request.QueryString("jobnumber")
SQL="select * from FORM where NID='"&id&"'"
rs.open SQL,conn,1,3
if not rs.eof then
	for pi=1 to 4
	thisoutput=""
	thiscoutput=""
	getFormParamHTML pi,rs("PARAM_TYPE"&pi),rs("PARAM"&pi),rs("PARAM_CHINESE"&pi),rs("PARAM_SCRIPTS"&pi),request.QueryString("paramvalue"&pi),rs("PARAM_SHOWBUTTON"&pi),rs("PARAM_CHINESE_SHOWBUTTON"&pi),rs("PARAM_BUTTON_SCRIPTS"&pi),rs("PARAM_TITLE"&pi),rs("PARAM_CHINESE_TITLE"&pi),thisoutput,thiscoutput,jobnumber
	output=output&"parent.paramhtml"&pi&".innerHTML="""&thisoutput&""";"
	coutput=coutput&"parent.paramhtml"&pi&".innerHTML="""&thiscoutput&""";"
	if request.QueryString("paramvalue"&pi)<>"" and rs("PARAM_BUTTON_SCRIPTS"&pi)<>"" then
	scripts=scripts&"parent."&rs("PARAM_BUTTON_SCRIPTS"&pi)&";"
	end if
	next

	'approveflow=rs("APPROVE1")
	actperson=rs("ACT_PERSON")
	
	flow=""
	cflow=""
	if instr(actperson,"applicant")>0 then
	person=person&"Applicant;"
	cperson=cperson&"申请人;"
	end if
	if instr(actperson,"groupleader")>0 then
	person=person&"Job's Group Leader;"
	cperson=cperson&"工单产线的领班;"
	end if
	if instr(actperson,"supervisor")>0 then
	person=person&"Job's Supervisor;"
	cperson=cperson&"工单产线的生产工程师;"
	end if
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
parent.approveflow.innerHTML="<%=flow%>";
parent.actperson.innerHTML="<%=person%>";
<%else%>
<%=coutput%>
parent.approveflow.innerHTML="<%=cflow%>";
parent.actperson.innerHTML="<%=cperson%>";
<%end if%>
<%=scripts%>
</script>
</head>

<body>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->