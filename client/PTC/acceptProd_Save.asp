<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->


<%
PTCstate=request("PTCstate")
SNNO=request("SNNO")
SelectType=request("SelectType")
if PTCstate="accept" then
if SelectType="不归还" then

Istate="结束"
else


Istate="已接收"

end if
 
acceptDate=now()

conn.Execute("update PTC_SN set PTCstate='"&Istate&"',acceptDate='"&acceptDate&"' where SNNO='"&SNNO&"'")

response.Redirect("acceptProd.asp")

elseif PTCstate="Return" then
Istate="正在归还"
returnDate=now()

conn.Execute("update PTC_SN set PTCstate='"&Istate&"',returnDate='"&returnDate&"' where SNNO='"&SNNO&"'")

response.Redirect("acceptProd.asp")

elseif PTCstate="pending"  then

Istate="部分归还"
returnDate=now()

conn.Execute("update PTC_SN set PTCstate='"&Istate&"',returnDate='"&returnDate&"' where SNNO='"&SNNO&"'")

response.Redirect("acceptProd.asp")



else
Istate="拒收"
rejectDate=now()
conn.Execute("update PTC_SN set PTCstate='"&Istate&"',rejectDate='"&rejectDate&"' where SNNO='"&SNNO&"'")

response.Redirect("acceptProd.asp")

end if



%>