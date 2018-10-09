<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->


<%
PTCstate=request("PTCstate")
SNNO=request("SNNO")

BadNO=request("BadNo")



if PTCstate="End" then

Istate="�ѹ黹"
EndDate=now()

conn.Execute("update PTC_SN set PTCstate='"&Istate&"',EndDate='"&EndDate&"'  where SNNO='"&SNNO&"'")

response.Redirect("EndProd.asp")


end if



%>