<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
jobnumber=request.QueryString("jobnumber")
SQL="select 1,J.*,U.USER_CODE,U.USER_NAME,U.USER_CHINESE_NAME,U.EMAIL from JOB_MASTER J left join LINE L on J.LINE_NAME=L.LINE_NAME left join USERS U on L.LEADER=U.USER_CODE where JOB_NUMBER='"&jobnumber&"'"
rs.open SQL,conn,1,3
if not rs.eof then
	line_name=rs("LINE_NAME")
	factory_id=rs("FACTORY_ID")
	start_quantity=rs("START_QUANTITY")
	group_leader=rs("USER_CODE")
	group_leader_mail=rs("EMAIL")
	if session("language")="0" then
	word="Part Number: "&rs("PART_NUMBER_TAG")&"; Line Name: "&rs("LINE_NAME")&"; Start Quantity: "&rs("START_QUANTITY")&"; Group Leader: "&rs("USER_NAME")
	else
	word="型号: "&rs("PART_NUMBER_TAG")&"；线别："&rs("LINE_NAME")&"；开始数量: "&rs("START_QUANTITY")&"；领班：" &rs("USER_CHINESE_NAME")
	end if
else
	if session("language")="0" then
	word="Job has not produced on lines or is not exsited!"
	else
	word="该工单没有生产或不存在！"
	end if
end if
rs.close
if request.QueryString("objectname")="New Line" then
	'generate line list
	options=""
	i=1
	SQL="select LINE_NAME from LINE where LINE_NAME<>'"&line_name&"' and FACTORY_ID='"&factory_id&"' order by LINE_NAME"
	rs.open SQL,conn,1,3
	if not rs.eof then
	while not rs.eof
	options=options&"var newElem"&i&"=document.createElement('option');newElem"&i&".text='"&rs("LINE_NAME")&"';newElem"&i&".value='"&rs("LINE_NAME")&"';parent.form1.param"&request.QueryString("objectindex")&".options.add(newElem"&i&");"
	i=i+1
	rs.movenext
	wend
	end if
	rs.close
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>Untitled Document</title>
<script language="javascript" type="text/javascript">
parent.document.all.jobinfohtml.innerHTML="<%=word%>";
parent.document.all.note.value="<%=word%>";
parent.document.all.line_name.value="<%=line_name%>";
parent.document.all.start_quantity.value="<%=start_quantity%>";
parent.document.all.group_leader.value="<%=group_leader%>";
parent.document.all.group_leader_mail.value="<%=group_leader_mail%>";
<%=options%>
</script>
</head>

<body>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->