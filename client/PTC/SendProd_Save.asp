<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->


<%
selectType=trim(request("selectType"))
txtSubJobList=trim(request("txtSubJobList"))
SNNO=trim(request("SNNO"))
BoxNO=trim(request("BoxNO"))
Sendarea=trim(request("Sendarea"))
LendNO=trim(request("LendNO"))
OperatorCode=trim(request("txtOperatorCode"))
acceptCode=trim(request("acceptCode"))
getCode=trim(request("getCode"))
txtComments=trim(request("txtComments"))
area=trim(request("area"))
PTC=trim(request("PTC"))
ExpectDate=trim(request("ExpectDate"))

if PTC="OK" then

SendDate=now()
PTCstate="等待接收"
conn.Execute("update PTC_SN set PTCstate='"&PTCstate&"',SendDate='"&SendDate&"' where SNNO='"&SNNO&"'")
response.Redirect("SendProd.asp")

else

set rs=server.CreateObject("adodb.recordset")
sql= "select * from  PTC_SN where SNNO='"&SNNO&"'"

rs.Open sql,conn,1,3
if rs.bof and rs.eof then

PTCstate="等待确认借出"
SendDate=now()
'LendNO=0
BADNO=0
RETURNNOJSQ=0
RETURNNO=0
conn.Execute("INSERT INTO PTC_SN (SNNO,BoxNo,BADNO,RETURNNOJSQ,RETURNNO,OperatorCode,Sendarea,SelectType,getCode,acceptCode,txtComments,PTCstate,area,SendDate,ExpectDate) values ('"&SNNO&"','"&BoxNO&"','"&BADNO&"','"&RETURNNOJSQ&"','"&RETURNNO&"','"&OperatorCode&"','"&Sendarea&"','"&SelectType&"','"&getCode&"','"&acceptCode&"','"&txtComments&"','"&PTCstate&"','"&area&"',sysdate,'"&ExpectDate&"')")
response.Redirect ("Barcode_add.asp?SNNO="&SNNO)
else
response.Redirect ("Barcode_add.asp?SNNO="&SNNO)

end if

end if



%>