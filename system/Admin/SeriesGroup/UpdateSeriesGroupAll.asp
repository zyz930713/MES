<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/SeriesGroup/SeriesGroupCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.Form("path")
query=request.Form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
j=0
for i=1 to request.Form("idcount")
	if request.Form("id"&i)="1" then
		SQL="select * from SERIES_GROUP where NID='"&request.Form("nid"&i)&"'"
		rs.open SQL,conn,1,3
		if not rs.eof then
			if rs("INCLUDED_SERIES")<>"" then
			included_system_items=""
				set rsS=server.CreateObject("adodb.recordset")
				SQLS="select INCLUDED_SYSTEM_ITEMS from SERIES where NID in ('"&replace(rs("INCLUDED_SERIES"),",","','")&"') order by SERIES_NAME"
				rsS.open SQLS,conn,1,3
				if not rsS.eof then
					while not rsS.eof
						if instr(included_system_items,rsS("INCLUDED_SYSTEM_ITEMS"))<=0 then
						included_system_items=included_system_items&rsS("INCLUDED_SYSTEM_ITEMS")&","
						end if
					rsS.movenext
					wend
				end if
				rsS.close
				set rsS=nothing
			end if
			if included_system_items<>"" then
			included_system_items=left(included_system_items,len(included_system_items)-1)
			end if
		rs("INCLUDED_SYSTEM_ITEMS")=included_system_items
		j=j+1
		rs.update
		end if
		rs.close
	end if
next
word="Successfully edit Series Group ("&j&")."
action="location.href='"&beforepath&"'"
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