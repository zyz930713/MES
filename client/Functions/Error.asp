<%@ language="VBScript" %>
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Functions/SendJMail.asp" -->
<%
 
  'Option Explicit
  
  Const lngMaxFormBytes = 200

  Dim objASPError, blnErrorWritten, strServername, strServerIP, strRemoteIP
  Dim strMethod, lngPos, datNow, strQueryString, strURL, errortype, HTML, reciever, JMail,jobnumber

  If Response.Buffer Then
    Response.Clear
    Response.Status = "500 Internal Server Error"
    Response.ContentType = "text/html"
    Response.Expires = 0
  End If
  errortype=session("errortype")

  Set objASPError = Server.GetLastError
  
  httpHost = Request.ServerVariables("HTTP_HOST")
   
  jobnumber=session("JOB_NUMBER")&"-"&replace(session("JOB_TYPE"),"N","")&repeatstring(session("SHEET_NUMBER"),"0",3)
   
 
  
  HTML="<!DOCTYPE html PUBLIC '-//W3C//DTD XHTML 1.0 Transitional//EN' 'http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd'><html xmlns='http://www.w3.org/1999/xhtml'><head><meta http-equiv='Content-Type' content='text/html; charset=gb2312' /><title>"&jobnumber&" Error Information</title><link href='"&httpHost&"/CSS/General.css' rel='stylesheet' type='text/css' /></head><body><table width='410' cellpadding='3' cellspacing='5'><tr><td width='400' colspan='2'><p class='style1'>Technical info/技术信息</p><ul><li>Error type: /错误类型：<br>"
  Dim bakCodepage
  bakCodepage = Session.Codepage
  Session.Codepage = 936
  
 


  
  HTML=HTML&Server.HTMLEncode(objASPError.Category)
  If objASPError.ASPCode > "" Then
  HTML=HTML&Server.HTMLEncode(", " & objASPError.ASPCode)
  session("aerror")=objASPError.ASPCode
  HTML=HTML&Server.HTMLEncode(" (0x" & Hex(objASPError.Number) & ")" ) & "<br>"
  end if
  
  
   if instr(objASPError.Description,"TNS")>0 then
 
 HTML="<BODY><font size='20'>数据库数据备份，请过一小时再试！</font>"
 HTML=HTML&"</body></html>"
 else
 
  If objASPError.ASPDescription > "" Then
 ' Session.Codepage = bakCodepage
  HTML=HTML&Server.HTMLEncode(objASPError.ASPDescription) & "<br>"
  end if
  
 
  
  
  blnErrorWritten = False

  ' Only show the Source if it is available and the request is from the same machine as IIS
  If objASPError.Source > "" Then
    strServername = LCase(Request.ServerVariables("SERVER_NAME"))
    strServerIP = Request.ServerVariables("LOCAL_ADDR")
    strRemoteIP = Request.ServerVariables("REMOTE_ADDR")
    If (strServername = "localhost" Or strServerIP = strRemoteIP) And objASPError.File <> "?" Then
      HTML=HTML&Server.HTMLEncode(objASPError.File)
      If objASPError.Line > 0 Then
	  HTML=HTML&", at the " & objASPError.Line & " row <br>"
	  end if
      If objASPError.Column > 0 Then
	  HTML=HTML& ", at the " & objASPError.Column & " column"
      HTML=HTML&"<br>"
      HTML=HTML& Server.HTMLEncode(objASPError.Source) & "<br>"
	  end if
      If objASPError.Column > 0 Then
	  HTML=HTML& String((objASPError.Column - 1), "-") & "<br>"
      HTML=HTML& "</font>"
      blnErrorWritten = True
	  end if
    End If
  End If

  If Not blnErrorWritten And objASPError.File <> "?" Then
    HTML=HTML&Server.HTMLEncode(objASPError.File)
    If objASPError.Line > 0 Then
	HTML=HTML&Server.HTMLEncode(", at the " & objASPError.Line & " row")
	HTML=HTML&"<br>"
	end if
    If objASPError.Column > 0 Then
	HTML=HTML&", at the " & objASPError.Column & " column"
    HTML=HTML&"<br>"
	end if
	HTML=HTML& objASPError.Description &"<br>"
  End If
  HTML=HTML&"</li>"
  'HTML=HTML&"<p><li>Type of internet explorer/浏览器类型：<br>"
  'HTML=HTML&Server.HTMLEncode(Request.ServerVariables("HTTP_USER_AGENT"))&"</li><p>"
  HTML=HTML&"<li>页：<br>"
  strMethod = Request.ServerVariables("REQUEST_METHOD")

  HTML=HTML&strMethod & " "
  If strMethod = "POST" Then
  HTML=HTML&Request.TotalBytes & " bytes to "
  End If

  HTML=HTML&Request.ServerVariables("SCRIPT_NAME")
  lngPos = InStr(Request.QueryString, "|")

  If lngPos > 1 Then
  HTML=HTML&"?" & Server.HTMLEncode(Left(Request.QueryString, (lngPos - 1)))
  End If

  Response.Write "</li>"
  If strMethod = "POST" Then
  HTML=HTML&"<p><li>POST 数据:<br>"
  'HTML=HTML&replace(Server.HTMLEncode(Request.Form),"&amp;","<br>")
  HTML=HTML&"</li>"
  End If
  HTML=HTML&"<p><li>Time/时间：<br>"
  datNow = Now()
  HTML=HTML& Server.HTMLEncode(FormatDateTime(datNow, 1) & ", " & FormatDateTime(datNow, 3))

  HTML=HTML&"</p></font></li><p><li>Customized Error/自定义错误：<br>Computer/电脑："
  HTML=HTML&request.ServerVariables("REMOTE_HOST")&"<br>Job Number/工单号："
  HTML=HTML&jobnumber&"<br>Line Name/线别："
  HTML=HTML&session("LINE_NAME")&"<br>Current Station/当前站："
  HTML=HTML&session("CURRENT_STATION_NAME")&"<br>Operator/工号："
  HTML=HTML&session("CODE")&"<br>Error/错误"
  HTML=HTML&session("aerror")&"<br>"&errortype
  HTML=HTML&"</li></body></html>"

  Session.Codepage = bakCodepage
  HTML=HTML&"</td></tr></table>"
  end if
  reciever="Young.Li@knowles.com"
  
  SendJMail application("MailSender"),reciever,"Error notification from "&application("SystemName")&" <"&jobnumber&">",HTML

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">
<head>
<META NAME="ROBOTS" CONTENT="NOINDEX">
<title>Page cannot be displayed</title>
<META HTTP-EQUIV="Content-Type" Content="text-html; charset=gb2312">
<META NAME="MS.LOCALE" CONTENT="ZH-CN">
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
</head>

<body bgcolor="FFFFFF">
<table width="640" align="center" cellpadding="3" cellspacing="5">
  <tr>
    <td>Error occurred. System has sent notification to administrator.&nbsp;<br>
      Please wait and retry or return home page. <br>
    系统发生错误，已通知管理员。<br>
    请等待稍后再试或返回首页。</td>
  </tr>
  <tr><td><%=HTML %><BR>
 </td></tr>
  <tr>
    <td><div align="center">
      <input type="button" name="Button" value="Home Page/首页" onClick="javascript:location.href='/KEB/Station1.asp'">
    </div></td>
  </tr>
</table>
</body>
</html>