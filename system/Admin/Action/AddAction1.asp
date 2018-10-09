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
SQL="select * from ACTION where ACTION_NAME='"&actionname&"' and STATION_ID='"&station&"'"
rs.open SQL,conn,1,3
if rs.eof then
rs.addnew
NID="AC"&NID_SEQ("ACTION")
rs("NID")=NID
rs("ACTION_NAME")=actionname
rs("ACTION_CHINESE_NAME")=trim(request.Form("actionchinesename"))
rs("FACTORY_ID")=request.Form("factory")
rs("STATION_ID")=station
rs("ACTION_PURPOSE")=request.Form("actionpurpose")
rs("STATION_POSITION")=cint(request.Form("position"))
rs("APPEND_ALLOW")=request.Form("append")
rs("NULL_ALLOW")=request.Form("null")
rs("WITH_LOT")=request.Form("has_lot")
rs("VALID_MACHINE")=request.Form("machine")
rs("RELATED_ACTION_ID")=request.Form("relatedaction")
rs("MAX_QUANTITY")=request.Form("max")
rs("MIN_QUANTITY")=request.Form("min")
rs("ACTION_TYPE")=request.Form("actiontype")
rs("ELEMENT_TYPE")=request.Form("componenttype")
if not isnull(trim(request.Form("componentnumber"))) and trim(request.Form("componentnumber"))<>"" then
rs("ELEMENT_NUMBER")=request.Form("componentnumber")
end if
rs.update
rs.close
	if request.Form("actionpurpose")=1 or request.Form("actionpurpose")=2 or request.Form("actionpurpose")=3 then
	SQL="select * from PART_ACTION_VALUE where ACTION_ID='"&NID&"'"
	rs.open SQL,conn,1,3
		if rs.eof then
			for i=1 to rcount
				if trim(request.Form("material"&i))<>"" or trim(request.Form("machine"&i))<>"" then
					rs.addnew
					rs("PART_ID")=request.Form("part_id"&i)
					rs("ACTION_ID")=NID
					rs("MATERIAL")=trim(request.Form("material"&i))
					rs("MACHINE")=trim(request.Form("machine"&i))
					rs.update
				end if
			next
		end if
	rs.close
	SQL="select * from GROUP_ACTION_VALUE where ACTION_ID='"&NID&"'"
	rs.open SQL,conn,1,3
		if rs.eof then
			for i=1 to gcount
				if trim(request.Form("gmaterial"&i))<>"" or trim(request.Form("gmachine"&i))<>"" then
					rs.addnew
					rs("GROUP_ID")=request.Form("group_id"&i)
					rs("ACTION_ID")=NID
					rs("MATERIAL")=trim(request.Form("gmaterial"&i))
					rs("MACHINE")=trim(request.Form("gmachine"&i))
					rs.update
					end if
				next
		end if
	rs.close
	end if
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