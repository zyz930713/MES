<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/KESKQ_OPEN.asp" -->
<%
server.ScriptTimeout=200
SQL="select * from USERS order by USER_CODE"
rs.open SQL,conn,1,3
if not rs.eof then
i=0
iUpdate=0
iName=0
iChinesName=0
while not rs.eof
	SQLHR="select EnglishName,Name from Personnel01 where CODE='"&rs("USER_CODE")&"'"
	rsHR.open SQLHR,conn_web,1,3
	if not rsHR.eof then
		if isnull(rs("USER_NAME")) or rs("USER_NAME")<>rsHR("EnglishName") then
		rs("USER_NAME")=rsHR("EnglishName")
		iName=iName+1
		end if
		if isnull(rs("USER_CHINESE_NAME")) or rs("USER_CHINESE_NAME")<>rsHR("Name") then
		rs("USER_CHINESE_NAME")=rsHR("Name")
		iChinesName=iChinesName+1
		end if
	iUpdate=iUpdate+1
	end if
	rsHR.close
rs.update
rs.movenext
i=i+1
wend
end if
rs.close
response.Write("在"&i&"记录中：更新"&iUpdate&";<br>名字更新"&iName&"<br>中文名更新"&iChinesName)
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