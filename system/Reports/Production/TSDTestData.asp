<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>


<!--#include virtual="/WOCF/TSD_Open1.asp" -->



<%

code=request("code")
isQuery=Ture
if code <> "" then
	isQuery=true
	 sql="select  * from package a ,serial_index b where a.serial_id=b.serial_id and serialnumber = '"&code&"'" 
	rsTSD.open sql,connTSD,1,3
	
end if
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charsPRet=gb2312">
<title>Barcode System - Scan </title>
<!--#include virtual="/Language/Components/Lan_PageSplit.asp" -->

<script language="JavaScript" src="/Language/language.js" type="text/javascript"></script>
<script language=javascript src="/Components/sniffer.js" type=text/javascript></script>

<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<link href="/CSS/dynCalendar.css" rel="stylesheet" type="text/css">

</head>

<body onLoad="language_page();language(<%=session("language")%>);">

<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="10" class="t-c-greenCopy"><span id="inner_BrowseData"></span></td>
</tr>
<tr>
  <td height="20" colspan="10" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="50%"><span id="inner_User"></span>:
          <% =session("User") %></td>
      <td width="50%"><div align="right"><%if admin=true then%>
      <a href="/Admin/Packing_Plan/addPacking_Plan.asp?path=<%=path%>&query=<%=query%>" target="main" class="white"><span id="inner_AddRecord"></span></a>
      <%end if%></div></td>
    </tr>
  </table></td>
</tr>

<tr>
  <td height="20" class="t-t-Borrow"><div align="center">NO 序列</div></td>
 
  <td height="20" class="t-t-Borrow"><div align="center">2D Code 二维码</div></td>
  <td height="20" class="t-t-Borrow"><div align="center">Test ID</div></td>
  <td height="20" class="t-t-Borrow"><div align="center">Result 结果</td>
  <td height="20" class="t-t-Borrow"><div align="center">Defect 缺陷</div></td>
  <td height="20" class="t-t-Borrow"><div align="center">Criterion 标准</div></td>
  <td height="20" class="t-t-Borrow"><div align="center">Test Name 测试名称</div></td>
  <td height="20" class="t-t-Borrow"><div align="center">Machine 测试机</div></td>
  <td class="t-t-Borrow"><div align="center">Test Time 测试时间</div></td>
  

  </tr>
<%	
if isQuery then
i=1
while not rsTSD.eof
	result="Pass"
	trColor="#00CC00"
	defect="&nbsp;"
	criterion="&nbsp;"
	if rsTSD("testResult")="FAIL" then
		result="Fail"
		defect=rsTSD("failName")
		
			trColor="#FF0000"
		
	end if
%>
	<tr align="center" bgcolor="<%=trColor%>" >
		<td><%=i%></td>
		<td><%=rsTSD("serialNumber")%></td>
		<td><%=rsTSD("serial_id")%></td>
		<td><%=result%></td>
		<td><%=defect%></td>
		<td><%=criterion%></td>
		<td><%=rsTSD("testitem")%>&nbsp;</td>
		<td><%=rsTSD("testTime")%>&nbsp;</td>
		<td><%=rsTSD("testTime")%>&nbsp;</td>
	</tr>	
<%
i=i+1
rsTSD.movenext		
wend
end if
%>

</table>
<table width="100%" height="26">
<tr><td align="center"><input name="btnClose2" type="button"  id="btnClose2"  value="返回上一页面" onClick="javascript:history.go(-1);"></td>
</tr></table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/TSD_Close.asp" -->