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
ordername=request("ordername")
ordertype=request("ordertype")
factory=trim(request("factory"))
code=request("code")
operator_chinese_name=request("operator_chinese_name")

if ordername="" and ordertype="" then
order=" order by J.JOB_NUMBER desc"
else
order=" order by "&ordername&" "&ordertype
end if
where=""
if jobnumber<>"" then
where=where&" and J.JOB_NUMBER like '%"&jobnumber&"%'"
end if
if jobtype<>"" then
where=where&" and J.JOB_TYPE='"&jobtype&"'"
end if
if partnumber<>"" then
where=where&" and J.PART_NUMBER_TAG like '%"&partnumber&"%'"
end if
if code<>"" then
where=where&" and O.CODE='"&code&"'"
end if
if operator_chinese_name<>"" then
where=where&" and O.OPERATOR_CHINESE_NAME='"&operator_chinese_name&"'"
end if
if line<>"" then
where=where&" and lower(J.LINE_NAME)='"&lcase(line)&"'"
end if
if fromdate<>"" then
where=where&" and to_date(J.INPUT_TIME)>=to_date('"&fromdate&"','yyyy-mm-dd')"
end if
if todate<>"" then
where=where&" and to_date(J.INPUT_TIME)<=to_date('"&todate&"','yyyy-mm-dd')"
end if
if fromdate<>"" then
where=where&" and J.START_TIME>=to_date('"&fromdate&"','yyyy-mm-dd')"
elseif jobnumber="" then
fromdate=dateadd("d",-7,date())
where=where&" and J.START_TIME>=to_date('"&fromdate&"','yyyy-mm-dd')"
end if
if todate<>"" then
where=where&" and J.START_TIME<=to_date('"&todate&"','yyyy-mm-dd')"
end if
pagepara="&jobnumber="&jobnumber&"&code="&code&"&jobtype="&jobtype&"&partnumber="&partnumber&"&line="&line&"&factory="&factory&"&operator_chinese_name="&operator_chinese_name&"&timespan="&timespan&"&fromdate="&fromdate&"&todate="&todate&"&ordername="&ordername&"&ordertype="&ordertype
pagename="/Reports/Process/Operator/OperatorOutput.asp"
FactoryRight "J."

SELECT SUM(J.GOOD_QUANTITY) AS GOOD,O.CODE,O.OPERATOR_CHINESE_NAME FROM JOB_STATIONS J INNER JOIN OPERATORS O ON J.OPERATOR_CODE=O.CODE WHERE J.START_TIME>=TO_DATE AND J.CLOSE_TIME<=TO_DATE GROUP BY O.CODE,O.OPERATOR_CHINESE_NAME



'SQL="select SUM(GOOD_QUANTITY) from JOB_STATION where START_TIME>=fromdate and CLOSE_TIME<=todate¡±&where&factorywhereoutsideand&order

'SQL="select J.* from JOB_MASTER J inner join FACTORY F on J.FACTORY_ID=F.NID where J.JOB_NUMBER is not null "&where&factorywhereoutsideand&order

'SQL="select O.CODE,O.OPERATOR_CHINESE_NAME,L.LINE_NAME form JOB_STATION J inner join OPERATORS O on J.OPERATOR_CODE=O.CODE  inner join LINE L on O.LINE_ID=L.LINE_NAME where O.OPERATOR is not null"&where&factorywhereoutsideand&order


'SQL="select * from OPERATORS order by CODE"
session("SQL")=SQL
rs.open SQL,conn,1,3
%>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>OperatorOutput</title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<link href="/CSS/dynCalendar.css" rel="stylesheet" type="text/css">
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
    <td height="20" colspan="8" class="t-c-greenCopy"><span id="inner_Search"></span></td>
  </tr>
  <tr>
    <td width="0%" height="20">code</td>
    <td width="22%" height="20"><input name="jobnumber" type="text" id="jobnumber" value="<%=code%>" /></td>
    <td width="0%">operator_chinese_name</td>
    <td width="41%"><input name="jobnumber2" type="text" id="jobnumber2" value="<%=operator_chinese_name%>"></td>
    <td width="1%"><span id="inner_SearchLine"></span></td>
    <td width="5%"><input name="line" type="text" id="line" value="<%=line%>" size="6"></td>
    <td width="1%"><span id="inner_SearchCreateTime"></span></td>
    <td width="30%"><span id="inner_SearchFrom"></span>
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
    <td height="20"> </td>
    <td height="20"><input name="jobnumber" type="text" id="jobnumber" value="<%=jobnumber%>" />    </td>
    <td><span id="inner_SearchPartNumber"></span></td>
    <td><input name="partnumber" type="text" id="partnumber" value="<%=partnumber%>" /></td>
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
        <td height="24" colspan="5" class="t-t-Borrow"><div align="left"><strong>OperatorOutput</strong></div></td>
        </tr>
      <tr>
        <td class="t-t-Borrow"><div align="center"><span id="inner_NO"></span></div></td>
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
        <td height="24"><div align="center"><%=i%></div></td>
        <td height="24"><span class="blue" style="cursor:hand">¡¡<%=rs("CODE")%></span></td>
        <td><%=rs("OPERATOR_CHINESE_NAME")%></td>
        <td><%=rs("line")%></td>
        <td><div align="center">
            <div align="left"><span class="blue" style="cursor:hand">¡¡</span>
              <% SUM(ASSEMBLY_GOOD_QUANTITY) %><%
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
