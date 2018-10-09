<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"
%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/KESKQ_OPEN.asp" -->
<!--#include virtual="/Functions/GetScrapCost.asp" -->
<!--#include virtual="/Functions/GetUserReportTo.asp" -->
<!--#include virtual="/Functions/GetUserAgent.asp" -->
<!--#include virtual="/Functions/GetUserInfo.asp" -->
<!--#include virtual="/Functions/SendJMail.asp" -->

<%
nid=trim(request("nid")&"")
nid="CH00000066"

SQL="select * from JOB_MASTER_SCRAP where NID='" & nid & "'"
rs.open SQL,conn,1,3
if rs.eof then
	response.Write("Error: No such scrap number!")
	response.End()		
end if
rs.close

ScrapCost=GetScrapCost(nid)

dim approve(5)
dim amount(5)
set rst=server.CreateObject("adodb.recordset")

strUserCode=""


SQL="select * from JOB_MASTER_SCRAP where NID='" & nid & "'"
rs.open SQL,conn,1,3
if not rs.eof then
	strUserCode=trim(rs("SCRAP_CODE")&"")
end if
rs.close

strUserEName=GetUserInfo(strUserCode,"ENGLISH_NAME")
strUserCName=GetUserInfo(strUserCode,"CHINESE_NAME")

strPRE=GetUserReportTo(strUserCode,"code")
strPREName=GetUserReportTo(strUserCode,"ename")
isPRE=false

SQL="select * from USERS where USER_CODE='" & strPRE & "' and APPROVAL_ROLE_ID='AR00000001'"
rs.open SQL,conn,1,3
if rs.eof then

	strPRE=GetUserReportTo(strPRE,"code")
else
	isPRE=true	
	
end if
rs.close

SQL="select * from USERS where USER_CODE='" & strPRE & "' and APPROVAL_ROLE_ID='AR00000001'"
rs.open SQL,conn,1,3
if rs.eof then

	strPRE=GetUserReportTo(strPRE,"code")
else
	isPRE=true
end if
rs.close

if isPRE=false then
	strPRE=GetUserReportTo(strUserCode,"code")
end if

'strApproveList=strPRE
strApproveList=""

SQL="select * from FORM where NID='FO00000006'"

rs.open SQL,conn,1,3

if not rs.eof then
	if trim(rs("approve1") & "")<>"" then
		approve(0)=trim(rs("approve1") & "")
		amount(0)=trim(rs("amount1") & "")
		if trim(rs("amount1") & "")="" then
			amount(0)=0
		end if
	end if
	if trim(rs("approve2") & "")<>"" then
		approve(1)=trim(rs("approve2") & "")
		amount(1)=trim(rs("amount2") & "")
		if trim(rs("amount2") & "")="" then
			amount(1)=0
		end if		
	end if
	
	if trim(rs("approve3") & "")<>"" then
		approve(2)=trim(rs("approve3") & "")
		amount(2)=trim(rs("amount3") & "")
		if trim(rs("amount3") & "")="" then
			amount(2)=0
		end if		
	end if
	if trim(rs("approve4") & "")<>"" then
		approve(3)=trim(rs("approve4") & "")
		amount(3)=trim(rs("amount4") & "")
		if trim(rs("amount4") & "")="" then
			amount(3)=0
		end if		
	end if
	if trim(rs("approve5") & "")<>"" then
		approve(4)=trim(rs("approve5") & "")
		amount(4)=trim(rs("amount5") & "")
		if trim(rs("amount5") & "")="" then
			amount(4)=0
		end if		
	end if
			
end if
rs.close

strApproveList=strPRE


strDirectManager=strPRE

strOtherManager=""

iGrade=1

for i=1 to 4
	'response.Write(ScrapCost & "  VS " & amount(i) & "  VS " & approve(i) & "<br>")
	if trim(approve(i)&"")<>"" and cdbl(ScrapCost)>=cdbl(amount(i)) then
		
		'strApproveList=strApproveList & "," & approve(i)
		
		strDirectManager=GetCurrentManager(strDirectManager,"code")
		
		SQL="select * from USERS where USER_CODE='" & strDirectManager & "' and APPROVAL_ROLE_ID='" & trim(approve(i)&"") & "'"
		rs.open SQL,conn,1,3
		if rs.eof then
			
			SQL="select * from USERS where APPROVAL_ROLE_ID='" & trim(approve(i)&"") & "' order by APPROVAL_ROLE_SEQ"
			rst.open SQL,conn,1,3
			if not rst.eof then
				strOtherManager=trim(rst("USER_CODE")&"")
				if trim(GetUserAgent(strOtherManager)&"")<>"" then
					strOtherManager=GetUserAgent(strOtherManager)
				end if
			end if
			rst.close
			strApproveList=strApproveList & "," & strOtherManager
		else
			strApproveList=strApproveList & "," & strDirectManager	
		end if
		rs.close
		iGrade=iGrade+1
	end if	
next

response.Write(strApproveList & "---" & iGrade)
response.End()

if iGrade=1 then
	LastApprover=strPRE
else
	LastApprover=mid(strApproveList,instrrev(strApproveList,",")+1)
end if


SQL="select * from JOB_MASTER_SCRAP where NID='" & nid & "'"
rs.open SQL,conn,1,3
if not rs.eof then

	rs("APPROVE_STATUS")=0
	rs("APPROVE_GRADE")=iGrade
	rs("APPROVER_CODE")=strApproveList
	rs("APPROVE_STEP")=0
	rs("CURRENT_APPROVER")=strPRE
	rs("LAST_APPROVER")=LastApprover
	rs.update	
end if
rs.close

set rst=nothing

strHTML="<p>Hello, " & strPREName & ",</p><p><a href=""#"">Please approve scrap request from " & strUserEName & ". 请签核由" & strUserCName & "提交的报废申请。</a></p><p>Best regards,<br>"&application("SystemName")&"</p>"
strReceiver="dickens.xu@knowles.com"
strSender="BarcodeSystem@knowles.com"
strTitle="Scrap requet: " & nid
SendJMail strSender,strReceiver,strTitle,strHTML

word="Your scrap request is submit! 你的报废申请已提交！"
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
<!--#include virtual="/WOCF/KESKQ_CLOSE.asp" -->