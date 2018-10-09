<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Action/ActionCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<%
path=request.Form("path")
query=request.Form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
id=request.Form("id")
actionname=trim(request.Form("actionname"))
rcount=request.Form("rcount")
gcount=request.Form("gcount")
actioncode=request.Form("actioncode")

SQL="select * from ACTION_New where NID='"&id&"'"
rs.open SQL,conn,1,3
if not rs.eof then
	rs("ACTION_NAME")=actionname
	rs("ACTION_CHINESE_NAME")=trim(request.Form("actionchinesename"))
	if request.Form("factory")<>"" then
		rs("FACTORY_ID")=request.Form("factory")
	else
		rs("FACTORY_ID")=null
	end if
	rs("ACTION_PURPOSE")=request.Form("actionpurpose")
	rs("STATION_POSITION")=cint(request.Form("position"))
	rs("APPEND_ALLOW")=request.Form("append")
	rs("NULL_ALLOW")=request.Form("null")
	rs("WITH_LOT")=request.Form("has_lot")
	rs("ACTION_TYPE")=request.Form("actiontype")
	rs("ELEMENT_TYPE")="TEXT"
	
	if request.Form("componentnumber")<>"" then
		rs("ELEMENT_NUMBER")=request.Form("componentnumber")
	end if
	rs("IS_DELETE")=0
	rs("LASTUPDATE_PERSON")=session("code")
	rs("LASTUPDATE_TIME")=date()
	rs("ACTION_CODE")=actioncode
	rs.update
	rs.close
	word="Successfully edit Action."
	action="location.href='"&beforepath&"'"

		
	'Update the related action in action table 2011-1-17
	SQL="UPDATE ACTION SET ACTION_NAME='"+actionname+"',ACTION_TYPE='"+request.Form("actiontype")+"',"
	SQL=SQL+"ACTION_PURPOSE='"+request.Form("actionpurpose")+"',"
	SQL=SQL+"STATION_POSITION='"+request.Form("position")+"',"
	SQL=SQL+"APPEND_ALLOW='"+request.Form("append")+"',"
	SQL=SQL+"ELEMENT_NUMBER='"+request.Form("componentnumber")+"',"
	SQL=SQL+"ACTION_CHINESE_NAME='"+request.Form("actionchinesename")+"',"
	if request.Form("factory")<>"" then
		SQL=SQL+"FACTORY_ID='"+request.Form("factory")+"',"
	else
		SQL=SQL+"FACTORY_ID=NULL,"
	end if
	SQL=SQL+"NULL_ALLOW='"+request.Form("null")+"',"
	SQL=SQL+"WITH_LOT='"+request.Form("has_lot")+"'"
	SQL=SQL+" WHERE MOTHER_ACTION_ID='"+id+"'"
	set rs0=server.createobject("adodb.recordset")
	rs0.open sql,conn,1,3
	'END ADD

else
	word="Action of "&actionname&" has not existed, please input again."
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