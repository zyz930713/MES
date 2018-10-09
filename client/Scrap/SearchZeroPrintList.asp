<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
pagename="SearchPrePrintList.asp"
jobnumber=trim(request("jobnumber"))
if jobnumber <> "" then
	SQL="SELECT * FROM SCRAP_PRINT A WHERE EXISTS (SELECT 1 FROM JOB_MASTER_SCRAP_PRE WHERE INSTR(A.PRINT_MEMBERS,NID)>0 AND JOB_NUMBER LIKE '"&jobnumber&"%')"
	rs.open SQL,conn,1,3	
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=application("SystemName")%></title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
</head>

<body bgcolor="#339966" onLoad="form1.jobnumber.focus();">
<form id="form1" method="post" name="form1">
<table border="1" width="800" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
	<tr>
		<Td height="20" colspan="8" class="t-t-DarkBlue"  align="center">Query Print List 查询打印清单</Td>
	</tr>	
	<tr align="center">
		<Td  width="130">Job Number 工单号</Td>
		<Td width="140" ><input type="text" id="jobnumber" name="jobnumber" value="<%=jobnumber%>"> 
		</Td>
		<Td align="left">	
			<input name="btnQuery" type="button"  id="btnQuery" onClick="if(form1.jobnumber.value){form1.submit();}" value="Query 查询">			
		</Td>
	</tr>
</table>
</form>

<%if jobnumber <> "" then%>
	<table width="800" border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000000" bordercolordark="#FFFFFF">
	<tr>
		<td class="today" width="50" ><div align="center">No 序号</div></td>
		<td class="today" width="100" ><div align="center">Print Id 打印编号</div></td>
		<td class="today"><div align="center">Scrap Id 报废单号</div></td>
	  </tr>
	<%
	if not rs.eof then
		i=1		
		while not rs.eof
	%>
		  <tr>
			<td><div align="center"><%=i%></div></td>
			<td><div align="center"><a href="/Scrap/PrePrintScrapList.asp?printid=<%=rs("NID")%>"><%=rs("NID")%></a></div></td>
			<td><%=formatlongstring(rs("PRINT_MEMBERS"),"<br>",100)%></td>
		  </tr>
	<%     i=1+1
			rs.movenext
		wend
	else
	%>
	<tr>
		<td colspan="3"><div align="center">No Records 没有记录</div></td>
	  </tr>
	<%end if%>
	</table>
<%rs.close
end if%>	
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->