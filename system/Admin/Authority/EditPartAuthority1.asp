<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Authority/AuthorityCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.form("path")
query=request.form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
id=request.Form("id")
partnumber=request.Form("partnumber")
fromcode=replace(request.Form("fromitem")," ","")
afromcode=split(fromcode,",")
tocode=replace(request.Form("toitem")," ","")
atocode=split(tocode,",")

SQL="Select CODE,AUTHORIZED_PARTS_ID from OPERATORS where code in ('"&replace(tocode,",","','")&"')"
rs.open SQL,conn,1,3
if not rs.eof then
	while not rs.eof
		for i=0 to ubound(atocode)
			if rs("CODE")=atocode(i) then
				if instr(rs("AUTHORIZED_PARTS_ID"),id)=0 then
				rs("AUTHORIZED_PARTS_ID")=rs("AUTHORIZED_PARTS_ID")&","&id
				rs.update
				end if
			end if
		next
	rs.movenext
	wend
end if
rs.close

SQL="Select CODE,AUTHORIZED_PARTS_ID from OPERATORS where code in ('"&replace(fromcode,",","','")&"')"
rs.open SQL,conn,1,3
if not rs.eof then
	while not rs.eof
		for i=0 to ubound(afromcode)
			if rs("CODE")=afromcode(i) then
				if instr(rs("AUTHORIZED_PARTS_ID"),id&",")<>0 then
				rs("AUTHORIZED_PARTS_ID")=replace(rs("AUTHORIZED_PARTS_ID"),id&",","")
				rs.update
				end if
				if instr(rs("AUTHORIZED_PARTS_ID"),id)<>0 then
				rs("AUTHORIZED_PARTS_ID")=replace(rs("AUTHORIZED_PARTS_ID"),id,"")
				rs.update
				end if
			end if
		next
	rs.movenext
	wend
end if
rs.close

word="Suscessfully update "&partnumber&"'s authority"
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
<%=action%>
</script>
<body>

</body>
</html>
