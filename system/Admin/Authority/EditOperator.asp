<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Authority/AuthorityCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/Functions/GetEngineerRole.asp" -->
<%
pagename="NewEngineer.asp"
path=request.QueryString("path")
query=request.QueryString("query")
id=request.QueryString("id")
SQL="select * from OPERATORS where NID='"&id&"'"
rs.open SQL,conn,1,3
%>
<html>
<head>
<title>Create Role</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<link href="/CSS/dynCalendar.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Admin/Authority/OperatorFormCheck.js" type="text/javascript"></script>
<script language=javascript src="/Components/sniffer.js" type=text/javascript></script>
<script language=javascript src="/Components/dynCalendar.js" type=text/javascript></script>
<!--#include virtual="/Language/Admin/Authority/Lan_Operator.asp" -->
</head>

<body onLoad="language();">
<form name="form1" method="post" action="/Admin/Authority/EditOperator1.asp" onSubmit="return formcheck()">
  <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolorlight="#666666" bordercolordark="#FFFFFF">
    <tr> 
      <td height="20" colspan="2" class="t-c-greenCopy"><span id="inner_EditData"></span> </td>
    </tr>
    <tr> 
      <td height="20" colspan="2" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="50%" class="white"><span id="inner_User"></span>:
              <% =session("User") %></td>
          <td width="50%"><div align="right"></div></td>
        </tr>
      </table></td>
    </tr>
    <tr>
      <td height="20"><span id="inner_SearchCode"></span> <span class="red">*</span></td>
      <td height="20"><input name="code" type="text" id="code" value="<%=rs("CODE")%>"></td>
    </tr>
    <tr> 
      <td width="112" height="20"><span id="inner_EnglishName"></span> <span class="red">*</span></td>
      <td width="572" height="20"><input name="name" type="text" id="name" value="<%=rs("OPERATOR_NAME")%>">	  </td>
    </tr>
    <tr bordercolorlight="#000099">
      <td height="20"><span id="inner_ChineseName"></span></td>
      <td height="20"><input name="chinesename" type="text" id="chinesename" value="<%=rs("OPERATOR_CHINESE_NAME")%>"></td>
    </tr>
    <tr bordercolorlight="#000099">
      <td height="20"><span id="inner_factory"></span><span class="red">*</span></td>
      <td height="20"><select name="factory" id="factory">
          <option value=""></option>
		  <%FactoryRight ""%>
          <%= getFactory("OPTION",rs("FACTORY_ID"),factorywhereinside,"","") %>
      </select></td>
    </tr>
    <tr bordercolorlight="#000099">
      <td height="20"><span id="inner_Locked"></span></td>
      <td height="20"><input name="locked" type="checkbox" id="locked" value="1" <%if rs("LOCKED")="1" then%>checked<%end if%>></td>
    </tr>
    <tr bordercolorlight="#000099">
      <td height="20"><span id="inner_Practised"></span></td>
      <td height="20"><input name="practised" type="checkbox" id="practised" value="1" <%if rs("PRACTISED")="1" then%>checked<%end if%>></td>
    </tr>
    <tr bordercolorlight="#000099">
      <td height="20"><span id="inner_PeriodFrom"></span></td>
      <td height="20"><input name="fromdate" type="text" id="fromdate" value="<%=rs("PRACTISE_START_TIME")%>" size="10">
      <script language=JavaScript type=text/javascript>
	function calendar1Callback(date, month, year)
	{
	document.all.fromdate.value=year + '-' + month + '-' + date
	}
    calendar1 = new dynCalendar('calendar1', 'calendar1Callback');
    </script></td>
    </tr>
    <tr bordercolorlight="#000099">
      <td height="20"><span id="inner_PeriodTo"></span></td>
      <td height="20"><input name="todate" type="text" id="todate" value="<%=rs("PRACTISE_END_TIME")%>" size="10">
        <script language=JavaScript type=text/javascript>
	function calendar2Callback(date, month, year)
	{
	document.all.todate.value=year + '-' + month + '-' + date
	}
    calendar2 = new dynCalendar('calendar2', 'calendar2Callback');
                        </script></td>
    </tr>
    <tr> 
      <td height="20" colspan="2">&nbsp;</td>
    </tr>
    <tr> 
      <td height="20" colspan="2"><div align="center">
          <input name="id" type="hidden" id="id" value="<%=id%>">
          <input name="path" type="hidden" id="path" value="<%=path%>">
          <input name="query" type="hidden" id="query" value="<%=query%>">
          <input type="submit" name="btnOK" value="OK">
          &nbsp; 
          <input type="reset" name="btnReset" value="Reset">
      </div></td>
    </tr>
  </table>
</form>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->