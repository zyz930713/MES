<%
function sendEmail(sender,reciever,subject,body)
	on error resume next
	set JMail=CreateObject("JMail.Message") 
	JMail.ContentType = "text/html"
	JMail.Charset ="gb2312"
	JMail.From = sender
	JMail.AddRecipient reciever
	'JMail.AddRecipient "Jack.zhang@knowles.com"
	JMail.Subject = subject
	JMail.Body = body
	session("aerror")=reciever
	JMail.Send (application("MailServer"))
	set JMail = Nothing
end function
%>