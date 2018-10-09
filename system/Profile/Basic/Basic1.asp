<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
SQL="select * from USERS where USER_CODE='"&session("code")&"'"
rs.open SQL,conn,1,3
if not rs.eof then
rs("USER_NAME")=trim(request.Form("englishname"))
rs("USER_CHINESE_NAME")=trim(request.Form("chinesename"))
rs("MANAGER")=trim(request.Form("manager"))
rs("BACKUP_PERSON_CODE")=trim(request.Form("backup"))
	if request.Form("language")<>"" then
	rs("LANGUAGE")=request.Form("language")
	end if
rs("EMAIL")=trim(request.Form("email"))
rs.update
word="Successfully update information."
action="location.href='"&beforepath&"'"
end if
rs.close
SQL="delete from USERS_YIELD_ALERT where USER_CODE='"&session("code")&"'"
rs.open SQL,conn,1,3
for i=1 to 15
	if request.Form("prefix"&i)<>"" then
		SQL="Select * from USERS_YIELD_ALERT where USER_CODE='"&session("code")&"' and ALERT_NAME='"&request.Form("prefix"&i)&"' and ALERT_TYPE=1"
		rs.open SQL,conn,1,3
		if rs.eof then
		rs.addnew
		end if
		rs("USER_CODE")=session("code")
		rs("ALERT_NAME")=ucase(request.Form("prefix"&i))
		rs("ALERT_YIELD")=csng(request.Form("model_yield"&i)/100)
		rs("ALERT_TYPE")=1
		rs.update
		rs.close
	end if
	if request.Form("line"&i)<>"" then
		SQL="Select * from USERS_YIELD_ALERT where USER_CODE='"&session("code")&"' and ALERT_NAME='"&request.Form("line"&i)&"' and ALERT_TYPE=2"
		rs.open SQL,conn,1,3
		if rs.eof then
		rs.addnew
		end if
		rs("USER_CODE")=session("code")
		rs("ALERT_NAME")=ucase(request.Form("line"&i))
		rs("ALERT_YIELD")=csng(request.Form("line_yield"&i)/100)
		rs("ALERT_TYPE")=2
		rs.update
		rs.close
	end if
	if request.Form("family"&i)<>"" then
		SQL="Select * from USERS_YIELD_ALERT where USER_CODE='"&session("code")&"' and ALERT_NAME='"&request.Form("family"&i)&"' and ALERT_TYPE=3"
		rs.open SQL,conn,1,3
		if rs.eof then
		rs.addnew
		end if
		rs("USER_CODE")=session("code")
		rs("ALERT_NAME")=ucase(request.Form("family"&i))
		rs("ALERT_YIELD")=csng(request.Form("family_yield"&i)/100)
		rs("ALERT_TYPE")=3
		rs.update
		rs.close
	end if

next
scrap_line=request("scrap_line")
SQL="delete from USERS_LINE where USER_CODE='"&session("code")&"'"
rs.open SQL,conn,1,3
if scrap_line<>"" then
arr_scrap_line=split(replace(scrap_line," ",""),",")
	for i=0 to ubound(arr_scrap_line)
		SQL="Select * from USERS_LINE where USER_CODE='"&session("code")&"' and LINE_ID='"&arr_scrap_line(i)&"'"
		rs.open SQL,conn,1,3
		if rs.eof then
		rs.addnew
		end if
		rs("USER_CODE")=session("code")
		rs("LINE_ID")=arr_scrap_line(i)
		rs.update
		rs.close
	next
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