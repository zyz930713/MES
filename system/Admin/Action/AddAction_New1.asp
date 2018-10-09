<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Action/ActionCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<!--#include virtual="/Functions/NID_SEQ.asp" -->
<%
path=request.Form("path")
query=request.Form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
actionname=trim(request.Form("actionname"))
station=request.Form("station")
rcount=request.Form("rcount")
actionname=replace(actionname,"'","''")
actioncode=request.Form("actioncode")

SQL="select * from ACTION_New where ACTION_NAME='"&actionname&"' and Is_Delete=0"
rs.open SQL,conn,1,3
if rs.eof then
	rs.addnew
	NID="AN"&NID_SEQ("ACTION_New")
	rs("NID")=NID
	rs("ACTION_NAME")=actionname
	rs("ACTION_CHINESE_NAME")=trim(request.Form("actionchinesename"))
	rs("FACTORY_ID")=request.Form("factory")
	rs("ACTION_PURPOSE")=request.Form("actionpurpose")
	rs("STATION_POSITION")=cint(request.Form("position"))
	rs("APPEND_ALLOW")=request.Form("append")
	rs("NULL_ALLOW")=request.Form("null")
	rs("WITH_LOT")=request.Form("has_lot")
	rs("ACTION_TYPE")=request.Form("actiontype")
	rs("ELEMENT_TYPE")="TEXT"
	if not isnull(trim(request.Form("componentnumber"))) and trim(request.Form("componentnumber"))<>"" then
		rs("ELEMENT_NUMBER")=request.Form("componentnumber")
	end if
	rs("IS_DELETE")=0
	rs("LASTUPDATE_PERSON")=session("code")
	rs("LASTUPDATE_TIME")=date()
	rs("ACTION_CODE")=actioncode
	rs.update
	rs.close
	 
	word="Successfully save a New Action."
	action="location.href='"&beforepath&"'"
else
	word="Action of "&actionname&" has existed, please input again."
	action="history.back()"
end if
%>
<!--#include virtual="/WOCF/BOCF_Commit.asp" -->
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