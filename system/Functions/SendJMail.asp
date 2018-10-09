<%
function SendJMail(sender,reciever,subject,body)
	on error resume next
	set JMail=CreateObject("JMail.Message") 
	JMail.ContentType = "text/html"
	JMail.Charset ="gb2312"
	JMail.From = sender
	if instr(reciever,";")>1 then
		array_mail=split(reciever,";")
		For i=0 to ubound(array_mail) 
			JMail.AddRecipient array_mail(i)
		Next
	else
		JMail.AddRecipient reciever
	end if
	JMail.Subject = subject
	JMail.Body = body
	JMail.Send (application("MailServer"))
	set JMail = Nothing
end function
%>