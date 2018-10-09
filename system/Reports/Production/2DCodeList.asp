<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>


<!--#include virtual="/WOCF/BOCF_Open.asp" -->



<%

code=request("Barcode")



sql="select  a.snno,a.selecttype, a.sendarea,a.area, b.barcode, b.codestate ,b.creatdate,a.operatorcode,  b.acceptdate,a.acceptcode, b.returndate, b.enddate  from ptc_SN  a,PTC_Barcodeno b  where a.snno= b.snno  and b.barcode='"&code&"'"

rs.open SQL,conn,1,3

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
  <td height="20" colspan="13" class="t-c-greenCopy"><span id="inner_BrowseData"></span></td>
</tr>
<tr>
  <td height="20" colspan="13" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
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
<td height="20" class="t-t-Borrow"><div align="center">序号</div></td>
  <td height="20" class="t-t-Borrow"><div align="center">单号</div></td>
   <td height="20" class="t-t-Borrow"><div align="center">借出类型</div></td>
   <td height="20" class="t-t-Borrow"><div align="center">2D Code 二维码</div></td>
    <td height="20" class="t-t-Borrow"><div align="center">二维码状态</div></td>
 <td height="20" class="t-t-Borrow"><div align="center">借出区</div></td>
 <td height="20" class="t-t-Borrow"><div align="center">接收区</div></td>
 
  <td height="20" class="t-t-Borrow"><div align="center">借出时间</div></td>
  <td height="20" class="t-t-Borrow"><div align="center">操作人员</div></td>
  <td height="20" class="t-t-Borrow"><div align="center">接收时间</td>
  <td height="20" class="t-t-Borrow"><div align="center">归还时间</div></td>
  <td height="20" class="t-t-Borrow"><div align="center">操作人员</div></td>
  <td height="20" class="t-t-Borrow"><div align="center">NPI接收时间</div></td>
  

  
  

  </tr>
<%
i=1
if not rs.eof then

while not rs.eof  
%>
<tr>
  <td height="20"><div align="center">
    <% =(cint(session("strpagenum"))-1)*recordsize+i%>
  </div>  </td>
 
    <td height="20"><div align="center"><%= rs("SNNO") %></div></td>
	<td height="20"><div align="center"><%= rs("selecttype") %></div></td>
	<td><%= rs("Barcode")%>&nbsp;</td>
	<td><%= rs("Codestate")%>&nbsp;</td>
	<td><%= rs("sendarea")%>&nbsp;</td>
	<td><%= rs("area")%>&nbsp;</td>	
	<td><%= rs("CreatDate")%>&nbsp;</td>
	<td><%= rs("operatorcode")%>&nbsp;</td>
	<td><%= rs("acceptdate")%>&nbsp;</td>
	<td><%= rs("returndate")%>&nbsp;</td>
	<td><%= rs("acceptcode")%>&nbsp;</td>
	<td><%= rs("EndDate")%>&nbsp;</td>

   
  </tr>
<%
i=i+1
rs.movenext
wend
else
%>
  <tr>
    <td height="20" colspan="13" align="center"><span id="inner_Records"></span>&nbsp;</td>
  </tr>
<%end if
rs.close%>

</table>
<table width="1019" height="26">
<tr><td align="center"><input name="btnClose2" type="button"  id="btnClose2"  value="返回上一页面" onClick="javascript:history.go(-1);"></td>
</tr></table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->