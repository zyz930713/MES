<%
function doEvent(eventname,jobnumber,series_name,series_group_name)
'set rsEv=server.CreateObject("adodb.recordset")
'SQL="select NID,EVENT_NOTE,EMAIL_BODY,EMAIL_HYPERLINK from EVENT where EVENT_NAME='"&eventname&"'"
'rsEv.open SQL,conn,1,3
'if not rsEv.eof then
'eventid=rsEv("NID")
'event_note=rsEv("EVENT_NOTE")
'email_body=rsEv("EMAIL_BODY")
'email_hyperlink=rsEv("EMAIL_HYPERLINK")
'end if 
'rsEv.close
'SQL="select EMAIL from USERS where EVENTS_ID like '%"&eventid&"%'"
'rsEv.open SQL,conn,1,3
'if not rsEv.eof then
'  set JMail=CreateObject("JMail.Message") 
'  JMail.ContentType = "text/html"
'  JMail.Charset ="gb2312"
'  JMail.From = "BarcodeSystem@knowles.com" 
'	  while not rsEv.eof
'	  JMail.AddRecipient rsEv("EMAIL")
'	  rsEv.movenext
'	  wend
'	  if event_note<>"" then
'	  event_note=replace(event_note,"&nbsp;"," ")
'	  end if
'  select case eventname
'  case "SERIES_GROUP_UPDATE"
'  JMail.Subject = "["&series_group_name&"] "&event_note
'  JMail.Body = "<a href='"&email_hyperlink&"'>["&series_group_name&"] "&email_body&"&nbsp;"&now()&"</a><br>"&application("SystemName")
'  end select
'  JMail.Send (application("MailServer"))
'  Set JMail = Nothing
'end if
'rsEv.close
'set rsEv=nothing
end function
%>