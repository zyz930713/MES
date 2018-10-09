<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Functions/GetWelcomeSettingsByDay.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")

if request.QueryString("month")<>"" then
thismonth=cint(request.QueryString("month"))
else
thismonth=month(date())
end if

if request.QueryString("year")<>"" then
thisyear=cint(request.QueryString("year"))
else
thisyear=year(date())
end if
	
pagename="/Admin/LCD/WelcomeList.asp"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title></title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Admin/LCD/formcheck.js" type="text/javascript"></script>
</head>

<body>
<table width="770" border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000000" bordercolordark="#FFFFFF" b="1" bcolorlight="#666666" bcolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="7" class="t-c-greenCopy">
      <table width="100%" cellpadding="0" cellspacing="0" b="0">
        <tr>
          <td>&nbsp;
              <%
			  if thismonth=12 then
			  nextyear=thisyear+1
			  nextmonth=1
			  preyear=thisyear
			  premonth=thismonth-1
			  elseif thismonth=1 then
			  preyear=thisyear-1
			  premonth=12
			  nextyear=thisyear
			  nextmonth=thismonth+1
			  else
			  preyear=thisyear
			  premonth=thismonth-1
			  nextyear=thisyear
			  nextmonth=thismonth+1
			  end if
			  %>
		  <a href="<%=pagename%>?year=<%=preyear%>&month=<%=premonth%>" class="white">Previous Month</a></td>
          <td><div align="center">LCD Settings  of <%=monthconvert(thismonth)%></div></td>
          <td><div align="right"><a href="<%=pagename%>?year=<%=nextyear%>&month=<%=nextmonth%>" class="white">Next Month</a></div></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr class="t-b-blue">
    <td height="20" class="red"><div align="center">Sunday</div></td>
    <td height="20"><div align="center">Monday</div></td>
    <td height="20"><div align="center">Tuesday</div></td>
    <td height="20"><div align="center">Wednesday</div></td>
    <td height="20"><div align="center">Thursday</div></td>
    <td height="20"><div align="center">Friday</div></td>
    <td height="20" class="green"><div align="center">Saturday</div></td>
  </tr>
  <%
		dates=cdate(thisyear&"-"&thismonth&"-1")
		While (year(dates)=thisyear and month(dates)=thismonth)
		%>
  <tr>
    <%		bgimg=""
			s=""
			i=1
			while i<=7
			if dates=date() then
			currentclass="t-b-blue"
			else
			currentclass=""
			end if
			if weekday(dates)=i and year(dates)=thisyear and month(dates)=thismonth then
				s=GetWelcomeSettingsByDay(dates,path,query)
				if i=1 then
				bgimg="Red_"&day(dates)
				elseif i=7 then
				bgimg="Green_"&day(dates)
				else
				bgimg="Gray_"&day(dates)
				end if
			dates=dateadd("d","1",dates)
			else
			s="&nbsp;"
			end if
	%>
    <td width="110" height="110" valign="middle" class="<%=currentclass%>" background="/Images/Date/<%=bgimg%>.gif" ><div align="center"><%=s%></div></td>
    <%	
			i=i+1
			bgimg=""
			s=""
			wend
	%>
  </tr>
  	<%
		wend
	%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->