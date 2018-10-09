<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%

Bstate=request("Bstate")	
SNNO=request("SNNO")	
if Bstate= "Delete"  then
	
sql="delete from PTC_SN where SNNO='"&SNNO&"'"

conn.execute sql

sqlP="delete from PTC_BarcodeNO where SNNO='"&SNNO&"'"

conn.execute sqlP
	
end if

response.Redirect("Admincode.asp")

%>