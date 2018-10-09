 <%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->

<!--#include virtual="/Functions/FormatFunctions.asp" -->
<script type="text/javascript">
	window.opener=null;   
  window.open("","_self");  
	window.close();
</script>
<%
reportTime=replace(dateadd("d",-1,now()),"/","-")
fromTime=formatdate(reportTime,"dd.mm.yyyy")&" 07:15"
toTime=formatdate(now(),"dd.mm.yyyy")&" 07:14"
reportDay=longweekdayconvert(weekday(reportTime))&","&formatdate(reportTime,"dd.mm.yyyy")

mailTo="Young.li@knowles.com;Vivian.Huang@knowles.com;ivan.li@knowles.com"



mailContents="Test"
'response.Write(mailContents)

SendJMail application("MailSender"),mailTo,"Daily Production Report",mailContents



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


<!--#include virtual="/WOCF/BOCF_Close.asp" -->