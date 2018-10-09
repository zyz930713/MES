<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<!--#include virtual="/Language/Lan_Main.asp" -->
</head>

<body onLoad="language()">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <%if request.QueryString("error")<>"" then%>
  <tr>
    <td height="20" class="red"><div align="center"><%=request.QueryString("error")%></div></td>
  </tr>
  <tr>
    <td height="20">&nbsp;</td>
  </tr>
  <%end if%>
  <tr> 
    <td height="20"><div align="center"><span id="inner_Welcome"></span></div></td>
  </tr>
  <tr>
    <td height="20">&nbsp;</td>
  </tr>
  <tr> 
    <td height="20"><table width="100%" border="0" cellspacing="1">
      <tr class="t-c-greenCopy">
        <td height="20" colspan="2"><span id="inner_Logon"></span></td>
        </tr>
      <tr class="t-c-GrayLight">
        <td height="20"><span id="inner_Code"></span></td>
        <td height="20"><% =session("code")%></td>
      </tr>
      <tr class="t-c-GrayLight">
        <td width="11%" height="20"><span id="inner_Name"></span></td>
        <td height="20"><% =session("user")%></td>
        </tr>
      <tr class="t-c-GrayLight">
        <td height="20"><span id="inner_Email"></span></td>
        <td height="20"><% =session("email")%></td>
      </tr>
      <tr class="t-c-GrayLight">
        <td height="20"><span id="inner_Factory"></span></td>
        <td height="20"><% =session("factory_name")%></td>
      </tr>
      <tr class="t-c-GrayLight">
        <td height="20"><span id="inner_Role"></span></td>
        <td width="89%" height="20">		
		<%if session("role")<>"" then
			roleName = session("role")
			if left(roleName,1)="," then
				roleName=right(roleName,len(roleName)-1)
			end if
			roleName = formatlongstring(replace(roleName,",",",&nbsp;"),"<br>",100)
			role_chinese_name=""
			SQL="select * from ENGINEER_ROLE where ROLE_NAME in ('"&replace(session("role"),",","','")&"')"
			rs.open SQL,conn,1,3
			while not rs.eof
				if rs("ROLE_CHINESE_NAME")<>"" then
					role_chinese_name=role_chinese_name&trim(rs("ROLE_CHINESE_NAME"))&","
				end if
				rs.movenext
			wend
			rs.close			
		end if%>
		<%=roleName%><br><%=role_chinese_name%>
		</td>
      </tr>
      <tr class="t-c-greenCopy">
        <td height="20" colspan="2"><span id="inner_Links"></span></td>
        </tr>
      <tr class="t-c-GrayLight">
        <td height="20">&nbsp;</td>		
        <td height="20">&nbsp;
		<!--
		<table width="200" border="1" cellspacing="0" cellpadding="0">
          <tr>
            <td nowrap><span id="inner_LinksScrap"></span></td>
            <td nowrap><span id="inner_LinksSubStoreConfirm"></span></td>
            <td nowrap><span id="inner_LinksScrapConfirm"></span></td>
            </tr>
          <tr>
            <td><span onClick="javascript:window.open('http://<%=application("ClientServerQA")%>/Scrap/Scrap1.asp?factory=<%=session("factory")%>')"><img src="/Images/Link_JobScrap.gif" width="60" height="60" style="cursor:hand"></span></td>
            <td><span onClick="javascript:window.open('http://<%=application("ClientServerQA")%>/SelectStoreConfirmationType.asp')"><img src="/Images/Link_JobStoreConfirm.gif" width="60" height="60" align="absmiddle" style="cursor:hand"></span></td>
            <td><span onClick="javascript:window.open('http://<%=application("ClientServerQA")%>/Scrap/Confirm1.asp')"><img src="/Images/Link_JobStoreConfirm.gif" width="60" height="60" align="absmiddle" style="cursor:hand"></span></td>
            </tr>
        </table>
		-->
		</td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="20">&nbsp;</td>
  </tr>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->