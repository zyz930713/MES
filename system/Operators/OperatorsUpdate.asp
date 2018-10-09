<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/KESKQ_OPEN.asp" -->
<!--#include virtual="/Functions/NID_SEQ.asp" -->
<%
server.ScriptTimeout=200
set rsL=server.CreateObject("adodb.recordset")
SQLHR="select Code,Name,EnglishName,Department,Line from personnel01 where EnglishName is not null and EnglishName<>'' order by code"
rskq.open SQLHR,conn_web,1,3
if not rskq.eof then
i=0
iNew=0
iUpdate=0
iName=0
iChinesName=0
iLine=0
iFactory=0
while not rskq.eof
	SQL="select NID,CODE,OPERATOR_NAME,OPERATOR_CHINESE_NAME,FACTORY_ID,LINE_ID from OPERATORS where CODE='"&rskq("Code")&"'"
	rs.open SQL,conn,3,3
	if rs.eof then
		rs.addnew
		rs("NID")="OP"&NID_SEQ("OPERATOR")
		rs("CODE")=rskq("Code")
		iNew=iNew+1
	end if
		if isnull(rs("OPERATOR_NAME")) or rs("OPERATOR_NAME")<>rskq("EnglishName") then
		rs("OPERATOR_NAME")=rskq("EnglishName")
		iName=iName+1
		end if
		if isnull(rs("OPERATOR_CHINESE_NAME")) or rs("OPERATOR_CHINESE_NAME")<>rskq("Name") then
		rs("OPERATOR_CHINESE_NAME")=rskq("Name")
		iChinesName=iChinesName+1
		end if
		if instr(rskq("Department"),"KE")<>0 then
		factory="FA00000002"
		end if
		if instr(rskq("Department"),"DELTEK")<>0 then
		factory="FA00000001"
		end if
		if instr(rskq("Department"),"VAM")<>0 then
		factory="FA00000003"
		end if
		if instr(rskq("Department"),"KES2")<>0 then
		factory="FA0000004"
		end if
		if isnull(rs("FACTORY_ID")) or rs("FACTORY_ID")<>factory then
		rs("FACTORY_ID")=factory
		iFactory=iFactory+1
		end if
			SQLL="select NID from LINE where LINE_NAME='"&rskq("Line")&"'"
			rsL.open SQLL,conn,1,3
			if not rsL.eof then
			line=rsL("NID")
			else
			line=""
			end if
			rsL.close
		if isnull(rs("LINE_ID")) or rs("LINE_ID")<>line then
		rs("LINE_ID")=line
		iLine=iLine+1
		end if
	rs.update
	iUpdate=iUpdate+1
	rs.close
rskq.movenext
i=i+1
wend
end if
rskq.close
set rsL=nothing
response.Write("在"&i&"记录中：新增"&iNew&";更新"&iUpdate&";<br>名字更新"&iName&"<br>中文名更新"&iChinesName&";<br>工厂更新"&iFactory&";<br>线别更新"&iLine)
%>
<!--#include virtual="/WOCF/KESKQ_CLOSE.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>Untitled Document</title>
</head>

<body>
</body>
</html>