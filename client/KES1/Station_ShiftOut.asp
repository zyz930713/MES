<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<%
code=trim(request.Form("code"))
'get Job Number
if instr(trim(request.Form("jobnumber")),"-")=0 then
response.Redirect("Station1.asp?errorstring=Job Number is error, please scan sub job！<br>工单错误，请扫描子工单！")
end if
ajobnumber=split(ucase(trim(request.Form("jobnumber"))),"-")
jobnumber=""
for w=0 to ubound(ajobnumber)
	if w<>ubound(ajobnumber) then
	jobnumber=jobnumber&ajobnumber(w)&"-"
	else
		if instr(ajobnumber(w),"R")=0 then'is not rework job
			if isnumeric(ajobnumber(w)) then
			sheetnumber=cint(ajobnumber(w))
			jobtype="N"
			else
			response.Redirect("Station1.asp?errorstring=Job Number is error, please re-scan sub job！<br>工单错误，请重新扫描子工单！")
			end if
		else
			if isnumeric(right(ajobnumber(w),len(ajobnumber(w))-1)) then
			sheetnumber=cint(right(ajobnumber(w),len(ajobnumber(w))-1))
			jobtype="R"
			else
			response.Redirect("Station1.asp?errorstring=Job Number is error, please re-scan sub job！<br>工单错误，请重新扫描子工单！")
			end if
		end if
	end if
next
jobnumber=left(jobnumber,len(jobnumber)-1)
'check whether user is recorded operator.
SQL="select OPERATOR_NAME,AUTHORIZED_STATIONS_ID,AUTHORIZED_PARTS_ID,LOCKED,PRACTISED,PRACTISE_START_TIME,PRACTISE_END_TIME from OPERATORS where code='"&code&"'"
rs.open SQL,conn,1,3
if rs.eof then
codevalid=false
valid=false
word="Operator Code of "&code&" is not existed, please contact engineer.<br>"&code&" 不存在，请联系工程师。"
else
	if rs("LOCKED")="1" then
	codevalid=false
	valid=false
	word="Operator Code of "&code&" is locked, please contact engineer.<br>"&code&" 被锁定，请联系工程师。"
	else
		if rs("PRACTISED")="1" and (date()<rs("PRACTISE_START_TIME") or date()>rs("PRACTISE_END_TIME")) then
			codevalid=false
			valid=false
			word="Practicing period of Operator Code of "&code&" is expired, please contact engineer.<br>"&code&" 的实习期限已过期，请联系工程师。"
		else
		codevalid=true
		session("operator")=rs("OPERATOR_NAME")
		AUTHORIZED_STATIONS_ID=rs("AUTHORIZED_STATIONS_ID")
		AUTHORIZED_PARTS_ID=rs("AUTHORIZED_PARTS_ID")
		end if
	end if
end if
rs.close

if codevalid=true then
	SQL="select * from JOB where STATUS=0 and JOB_NUMBER='"&jobnumber&"' and SHEET_NUMBER='"&sheetnumber&"' and JOB_TYPE='"&jobtype&"'"
	rs.open SQL,conn,1,3
	if not rs.eof then
	rs("STATUS")=5
	rs("SHIFT_OUT_PERSON")=rs("SHIFT_OUT_PERSON")&code&","
	rs("SHIFT_OUT_TIME")=rs("SHIFT_OUT_TIME")&now&","
	rs.update
	word=jobnumber&"-"&repeatstring(sheetnumber,"0",3)&"工单被暂时停线。"
	else
	word=jobnumber&"-"&repeatstring(sheetnumber,"0",3)&"工单当前状态不是开放的，不能被被暂时停线。"
	end if
	rs.close
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=application("SystemName")%></title>
<link href="/CSS/GeneralKES1.css" rel="stylesheet" type="text/css">
<script language="javascript">
alert("<%=word%>");
location.href="Station1.asp";
</script>
</head>

<body>
</body>
<!--#include virtual="/WOCF/BOCF_Commit.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" -->