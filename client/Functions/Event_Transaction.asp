<%
function doEvent(eventname,jobnumber)
'set rsEv=server.CreateObject("adodb.recordset")
'SQL="select NID,EVENT_NOTE,EMAIL_BODY from EVENT where EVENT_NAME='"&eventname&"'"
'rsEv.open SQL,conn,1,3
'if not rsEv.eof then
'eventid=rsEv("NID")
'event_note=rsEv("EVENT_NOTE")
'email_body=rsEv("EMAIL_BODY")
'end if 
'rsEv.close
'SQL="select EMAIL from USERS where EVENTS_ID like '%"&eventid&"%'"
'rsEv.open SQL,conn,1,3
'if not rsEv.eof then
'  set JMail=CreateObject("JMail.Message") 
'  JMail.ContentType = "text/html"
'  JMail.Charset ="gb2312"
'  JMail.From = "Barcode.System@knowles.com" 
'	  while not rsEv.eof
'	  JMail.AddRecipient rsEv("EMAIL")
'	  rsEv.movenext
'	  wend
'	  if event_note<>"" then
'	  event_note=replace(event_note,"&nbsp;"," ")
'	  end if
'  JMail.Subject = "<"&jobnumber&">"&event_note
'  JMail.Body = "<"&jobnumber&">"&email_body
'  JMail.Send (application("MailServer"))
'  Set JMail = Nothing
'end if
'rsEv.close
'set rsEv=nothing
end function
%>