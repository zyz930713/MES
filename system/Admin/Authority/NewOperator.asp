<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Authority/AuthorityCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/GetEngineerRole.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<%
pagename="NewOperator.asp"
path=request.QueryString("path")
query=request.QueryString("query")
%>
<html>
<head>
<title>Create Role</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<link href="/CSS/dynCalendar.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Admin/Authority/OperatorFormCheck.js" type="text/javascript"></script>
<script language="JavaScript" src="/Components/sniffer.js" type="text/javascript"></script>
<script language="JavaScript" src="/Components/dynCalendar.js" type="text/javascript"></script>
<!--#include virtual="/Language/Admin/Authority/Lan_Operator.asp" -->
<script language="JavaScript">
function preload()
{
document.form1.code.focus()
}
</script>
</head>

<body onLoad="preload();language();">
<form name="form1" method="post" action="/Admin/Authority/NewOperator1.asp" onSubmit="return formcheck()">
  <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolorlight="#666666" bordercolordark="#FFFFFF">
    <tr> 
      <td height="20" colspan="2" class="t-c-greenCopy"><span id="inner_AddData"></span></td>
    </tr>
    <tr> 
      <td height="20" colspan="2" class="t-c-greenCopy"><span id="inner_User"></span>:
        <% =session("User") %></td>
    </tr>
    <tr>
      <td height="20"><span id="inner_SearchCode"></span> <span class="red">*</span></td>
      <td height="20"><input name="code" type="text" id="code"></td>
    </tr>
    <tr> 
      <td width="113" height="20"><span id="inner_EnglishName"></span><span class="red">*</span></td>
      <td width="571" height="20"><input name="name" type="text" id="name">	  </td>
    </tr>
    <tr bordercolorlight="#000099">
      <td height="20"><span id="inner_ChineseName"></span></td>
      <td height="20"><input name="chinesename" type="text" id="chinesename"></td>
    </tr>
    <tr bordercolorlight="#000099">
      <td height="20"><span id="inner_factory"></span> <span class="red">*</span></td>
      <td height="20"><select name="factory" id="factory">
          <option value=""></option>
          <%= getFactory("OPTION","","","","") %>
      </select></td>
    </tr>
    <tr bordercolorlight="#000099">
      <td height="20"><span id="inner_Locked"></span></td>
      <td height="20"><input name="locked" type="checkbox" id="locked" value="1"></td>
    </tr>
    <tr bordercolorlight="#000099">
      <td height="20"><span id="inner_Practised"></span> </td>
      <td height="20"><input name="practised" type="checkbox" id="practised" value="1"></td>
    </tr>
    <tr bordercolorlight="#000099">
      <td height="20"><span id="inner_PeriodFrom"></span></td>
      <td height="20">      <input name="fromdate" type="text" id="fromdate" value="<%=date()%>" size="10">
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
      <td height="20">        <input name="todate" type="text" id="todate" value="<%=dateadd("m",3,date())%>" size="10">
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