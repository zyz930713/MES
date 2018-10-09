<%
function sendEmail(sender,reciever,subject,body)
	on error resume next
	set JMail=server.CreateObject("JMail.Message") 
	JMail.ContentType = "text/html"
	JMail.Charset ="gbk"
	JMail.From = sender
	a_reciever=split(reciever,",")
	for SE=0 to ubound(a_reciever)
	this_reciever=this_reciever&";"&a_reciever(SE)
	session("aerror")=reciever
	JMail.AddRecipient a_reciever(SE)
	next
	JMail.Subject = subject
	JMail.Body = "<!DOCTYPE html PUBLIC '-//W3C//DTD XHTML 1.0 Transitional//EN' 'http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd'><html xmlns='http://www.w3.org/1999/xhtml'><head><meta http-equiv='Content-Type' content='text/html; charset=gb2312' /><title>"&jobnumber&" Error Information</title></head><body>"&body&"</body></html>"
	JMail.Send (application("MailServer"))
	set JMail = Nothing
end function
%>