<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Job/JobCheck.asp" -->
<!--#include virtual="/Job/IsDBA.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/Functions/GetStation.asp" -->
<!--#include virtual="/Functions/GetStationTransactionChange.asp" -->
<!--#include virtual="/Functions/GetStationOperator.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
fromdate=request("fromdate")
todate=request("todate")
line=trim(request("line"))

SQL="select O.CODE,O.OPERATOR_CHINESE_NAME *,M.ASSEMBLY_GOOD_QUANTITY from OPERATORS AS O, JOB_MASTER AS M where O.CODE=OPERATE_CODE "
//SQL="select * from OPERATORS order by CODE"
rs.open SQL,conn,1,3
%>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>Untitled Document</title>
<link href="../../CSS/General.css" rel="stylesheet" type="text/css" />
<style type="text/css">
<!--
a:link {
	color: #333333;
}
a:visited {
	color: #FFFFFF;
}
-->
</style></head>

<body>
<form name="form1" method="get" action="/Reports/Process/Operator/OperatorOutput.asp">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="7" class="t-c-greenCopy"><span id="inner_Search"></span></td>
  </tr>
  <tr>
    <td width="1%" height="20"><span id="inner_SearchJobNumber"></span> </td>
    <td width="29%" height="20"><input name="jobnumber" type="text" id="jobnumber" value="<%=CODE%>" /></td>
    <td width="28%"><input name="jobnumber2" type="text" id="jobnumber2" value="<%=OPERATOR_CHINESE_NAME%>"></td>
    <td width="1%"><span id="inner_SearchLine"></span></td>
    <td width="7%"><input name="line" type="text" id="line" value="<%=line%>" size="6"></td>
    <td width="1%"><span id="inner_SearchCreateTime"></span></td>
    <td width="33%"><span id="inner_SearchFrom"></span>
      <input name="fromdate" type="text" id="fromdate" value="<%=fromdate%>" size="10">
      <script language=JavaScript type=text/javascript>
	function calendar1Callback(date, month, year)
	{
	document.all.fromdate.value=year + '-' + month + '-' + date
	}
    calendar1 = new dynCalendar('calendar1', 'calendar1Callback');
                        </script>
&nbsp;<span id="inner_SearchTo"></span>
<input name="todate" type="text" id="todate" value="<%=todate%>" size="10">
<script language=JavaScript type=text/javascript>
	function calendar2Callback(date, month, year)
	{
	document.all.todate.value=year + '-' + month + '-' + date
	}
    calendar2 = new dynCalendar('calendar2', 'calendar2Callback');
                        </script>
&nbsp; </td>
    </tr>
  <tr>
    <td height="20"><span id="inner_SearchPlaner"></span></td>
    <td height="20"><input name="OPERATOR_NAME" type="text" id="OPERATOR_NAME" value="<%=OPERATOR_NAME%>"></td>
    <td><input name="OPERATOR_CODE" type="text" id="OPERATOR_CODE" value="<%=OPERATOR_CODE%>"></td>
    <td colspan="4"><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()"></td>
    </tr>
</table>
</form>
<table width="100%" border="0">
  <tr>
    <td bgcolor="#666666"><table width="100%" border="0" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF" class="t-b-midautumn">
      <tr>
        <td width="14%" height="24"><span>
          User Name :
      <% =session("User") %>
        </span></td>
        <td width="86%"><table width="20%" border="0" align="right">
            <tr>
              <td><div align="right"></div></td>
            </tr>
        </table></td>
      </tr>
    </table></td>
  </tr>
</table>
<br />
<table width="100%" border="1" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" bordercolorlight="#000099" bordercolordark="#FFFFFF">
      <tr>
        <td height="24" colspan="5" class="t-t-Borrow"><div align="left"><strong>test</strong></div></td>
        </tr>
      <tr>
        <td class="t-t-Borrow"><div align="center"></div></td>
        <td height="24" bgcolor="#77B6DB" class="t-t-Borrow"><div align="center">code</div></td>
        <td bgcolor="#77B6DB" class="t-t-Borrow">Name</td>
        <td bgcolor="#77B6DB" class="t-t-Borrow">line</td>
        <td bgcolor="#77B6DB" class="t-t-Borrow"><div align="center">UNIT</div></td>
      </tr>
      <%

if not rs.eof then

while not rs.eof
i=i+1
%>
      <tr>
        <td height="24"><div align="center">
          <% =(csng(session("strpagenum"))-1)*recordsize+i%>
        </div></td>
        <td height="24"><span class="blue" style="cursor:hand">¡¡<%=rs("CODE")%></span></td>
        <td><%=rs("OPERATOR_CHINESE_NAME")%></td>
        <td><%=rs("line")%></td>
        <td><div align="center">
            <div align="left"><span class="blue" style="cursor:hand">¡¡</span>
              <%
	rs.movenext
wend
rs.close
end if
%>
          </div></td>
      </tr>
    </table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
