<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetSeriesGroup.asp" -->
<!--#include virtual="/Functions/GetEngineer.asp" -->
<!--#include virtual="/Functions/GetLine.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
pagename="/Profile/Basic/Basic.asp"
set rs1=server.CreateObject("adodb.recordset")
SQL="select U.*,F.FACTORY_NAME from USERS U inner join FACTORY F on instr(U.FACTORY_ID,F.NID)>0 where U.USER_CODE='"&session("code")&"'"
rs.open SQL,conn,1,3
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<!--#include virtual="/Language/Profile/Basic/Lan_Basic.asp" -->
<script language="JavaScript" src="/Profile/Basic/FormCheck.js" type="text/javascript"></script>
</head>

<body onLoad="language()">
<%if not rs.eof then%>
<form name="form1" method="post" action="/Profile/Basic/Basic1.asp" onSubmit="return formcheck()">
<table width="100%" border="1" cellspacing="0" bordercolorlight="#006633" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="6" class="t-c-greenCopy"><span id="inner_Browse"></span></td>
  </tr>
  <tr class="t-c-GrayLight">
    <td height="20"><span id="inner_Code"></span></td>
    <td height="20"><% =rs("USER_CODE")%></td>
    <td><span id="inner_Name"></span>&nbsp;<span class="red">*</span></td>
    <td><input name="englishname" type="text" id="englishname" value="<% =rs("USER_NAME")%>"></td>
    <td><span id="inner_ChineseName"></span>&nbsp;<span class="red">*</span></td>
    <td><input name="chinesename" type="text" id="chinesename" value="<% =rs("USER_CHINESE_NAME")%>"></td>
  </tr>
  <tr class="t-c-GrayLight">
    <td height="20"><span id="inner_Factory"></span></td>
    <td height="20"><% =rs("FACTORY_NAME")%></td>
    <td><span id="inner_Backup"></span></td>
    <td><select name="backup" id="backup">
      <option value="">-- Select Backup --</option>
      <%= getEngineer("OPTION",rs("BACKUP_PERSON_CODE")," Where U.USER_CODE<>'"&rs("USER_CODE")&"' and Length(U.USER_CODE)=4 and U.USER_CODE not like '9%'"," order by U.USER_CODE","") %>
    </select></td>
    <td>NT Account <span class="red">*</span></td>
    <td><%=rs("NT_ACCOUNT")%></td>
  </tr>
    <tr class="t-c-GrayLight">
      <td height="20">Department</td>
      <td height="20"><select name="Department" id="Department">
        <option value="">-- Select Department --</option>
        <%'= getDepartment("OPTION",rs("manager")," Where U.USER_CODE<>'"&rs("USER_CODE")&"' and Length(U.USER_CODE)=4 and U.USER_CODE not like '9%'"," order by U.USER_CODE","") %>
      </select></td>
      <td>Manager</td>
      <td><select name="manager" id="manager">
        <option value="">-- Select Manager --</option>
        <%= getEngineer("OPTION",rs("manager")," Where U.USER_CODE<>'"&rs("USER_CODE")&"' and Length(U.USER_CODE)=4 and U.USER_CODE not like '9%'"," order by U.USER_CODE","") %>
      </select></td>
      <td><span id="inner_Language"></span>&nbsp;<span class="red">*</span></td>
      <td><select name="language">
        <option value="">--Select--</option>
        <option value="0" <%if rs("LANGUAGE")="0" then%>selected<%end if%>>English</option>
        <option value="1" <%if rs("LANGUAGE")="1" then%>selected<%end if%>>中文</option>
      </select></td>
    </tr>
    <tr class="t-c-GrayLight">
      <td height="20"><span id="inner_Email"></span>&nbsp;</td>
      <td height="20"><input name="email" type="text" id="email" value="<% =rs("EMAIL")%>" size="50"></td>
      <td>Scrap Note </td>
      <td colspan="3"><%=getLine("CHECKBOX_SCRAP",line,where," order by L.LINE_NAME",null)%></td>
    </tr>
    
    
    <tr class="t-c-GrayLight">
    <td height="20" colspan="6"><table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td><span id="inner_YieldAlertModel"></span></td>
          <td><table width="100%" border="1" cellpadding="0" cellspacing="0">
            <tr class="t-b-abort">
              <%for m=1 to 3%>
              <td class="t-b-abort"><div align="center"><span id="inner_AlertNo<%=m%>"></span></div></td>
              <td><div align="center"><span id="inner_AlertModelPrefix<%=m%>"></span></div></td>
              <td><div align="center"><span id="inner_AlertModelYield<%=m%>"></span></div></td>
              <%next%>
            </tr>
            <%i=1
	SQL1="select * from USERS_YIELD_ALERT where USER_CODE='"&session("code")&"' and ALERT_TYPE=1"
	rs1.open SQL1,conn,1,3
	for m=1 to 5%>
            <tr>
              <%for n=1 to 3%>
              <td><div align="center"><%=i%></div></td>
              <td><div align="center">
                <input name="prefix<%=i%>" type="text" id="prefix<%=i%>" value="<%if not rs1.eof then%><%=rs1("ALERT_NAME")%><%end if%>" size="6" maxlength="10">
                </div></td>
              <td><div align="center">&lt;=
                <input name="model_yield<%=i%>" type="text" id="model_yield<%=i%>" value="<%if not rs1.eof then%><%=csng(rs1("ALERT_YIELD"))*100%><%end if%>" size="2" maxlength="2">
                %</div></td>
              <%
	  if not rs1.eof then
	  rs1.movenext
	  end if
	  i=i+1
	  next%>
              </tr>
            <%next
	rs1.close%>
          </table></td>
          <td><span id="inner_YieldAlertLine"></span></td>
          <td><table width="100%" border="1" cellpadding="0" cellspacing="0">
            <tr class="t-b-abort">
              <%for m=4 to 6%>
              <td class="t-b-abort"><div align="center"><span id="inner_AlertNo<%=m%>"></span></div></td>
              <td><div align="center"><span id="inner_AlertLineName<%=m%>"></span></div></td>
              <td><div align="center"><span id="inner_AlertModelYield<%=m%>"></span></div></td>
              <%next%>
            </tr>
            <%i=1
	  SQL1="select * from USERS_YIELD_ALERT where USER_CODE='"&session("code")&"' and ALERT_TYPE=2"
	  rs1.open SQL1,conn,1,3
	  for m=1 to 5%>
            <tr>
              <%for n=1 to 3%>
              <td><div align="center"><%=i%></div></td>
              <td><div align="center">
                <input name="line<%=i%>" type="text" id="line<%=i%>" value="<%if not rs1.eof then%><%=rs1("ALERT_NAME")%><%end if%>" size="6" maxlength="6">
                </div></td>
              <td><div align="center">&lt;=
                <input name="line_yield<%=i%>" type="text" id="line_yield<%=i%>" value="<%if not rs1.eof then%><%=csng(rs1("ALERT_YIELD"))*100%><%end if%>" size="2" maxlength="2">
                %</div></td>
              <%
	  if not rs1.eof then
	  rs1.movenext
	  end if
	  i=i+1
	  next%>
              </tr>
            <%next
	  rs1.close%>
          </table></td>
        </tr>

        </table></td>
    </tr>
    <tr class="t-c-GrayLight">
      <td height="20" colspan="6"><table width="100%" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td><span id="inner_YieldAlertFamily"></span></td>
            <td><table width="100%" border="1" cellpadding="0" cellspacing="0">
              <tr class="t-b-abort">
                <%for m=7 to 9%>
                <td class="t-b-abort"><div align="center"><span id="inner_AlertNo<%=m%>"></span></div></td>
                <td><div align="center"><span id="inner_AlertFamilyName<%=m%>"></span></div></td>
                <td><div align="center"><span id="inner_AlertModelYield<%=m%>"></span></div></td>
                <%next%>
              </tr>
              <%i=1
	SQL1="select * from USERS_YIELD_ALERT where USER_CODE='"&session("code")&"' and ALERT_TYPE=3"
	rs1.open SQL1,conn,1,3
	for m=1 to 5%>
              <tr>
                <%for n=1 to 3%>
                <td><div align="center"><%=i%></div></td>
                <td><div align="center">
                  <select name="family<%=i%>" id="family<%=i%>">
                    <option value="">-- Family --</option>
                    <%FactoryRight "S."
				if not rs1.eof then
				seriesgroup=rs1("ALERT_NAME")
				else
				seriesgroup=null
				end if%>
                    <%=getSeriesGroup("OPTION",seriesgroup,factorywhereoutside," order by S.SERIES_GROUP_NAME",null)%>
                    </select>
                  </div></td>
                <td><div align="center">&lt;=
                  <input name="family_yield<%=i%>" type="text" id="family_yield<%=i%>" value="<%if not rs1.eof then%><%=csng(rs1("ALERT_YIELD"))*100%><%end if%>" size="2" maxlength="2">
                  %</div></td>
                <%
	  if not rs1.eof then
	  rs1.movenext
	  end if
	  i=i+1
	  next%>
                </tr>
              <%next
	rs1.close%>
            </table></td>
          </tr>

            </table></td>
    </tr>
    <tr class="t-c-GrayLight">
      <td height="20" colspan="6"><span id="inner_Role"></span>&nbsp;</td>
    </tr>
    <tr class="t-c-GrayLight">
    <td height="20" colspan="6">
      <%if isnull(rs("ROLES_ID"))=false and rs("ROLES_ID")<>"" then%>
      <%=replace(rs("ROLES_ID"),","," ; ")%>
      <%end if%>      &nbsp;</td>
    </tr>
  <tr class="t-c-GrayLight">
    <td height="20" colspan="6"><div align="center">
      <input name="path" type="hidden" id="path" value="<%=path%>">
      <input name="query" type="hidden" id="query" value="<%=query%>">
<input name="Update" type="submit" id="Update" value="Update">
      &nbsp;
      <input type="reset" name="Submit2" value="Reset">
    </div></td>
    </tr>
</table>
</form>
<%end if%>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->