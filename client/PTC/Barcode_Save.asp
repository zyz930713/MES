<!--#include virtual="/WOCF/BOCF_Open.asp" -->

<%
Dcode=trim(request("Dcode"))
SNNO=trim(request("SNNO"))



set rs=server.CreateObject("adodb.recordset")
sql= "select * from PTC_BarcodeNO where SNNO='"&SNNO&"' and Barcode  in ("&Dcode&") " 



rs.Open sql,conn,1,3
if rs.bof and rs.eof then
tArr = Split(Dcode,",")
For i = 0 To UBound(tArr)
NDcode=(tArr(i))
if NDcode="'SANEND'" then

else

'CreatDate=now()

Set rsQ =  conn.Execute("INSERT INTO PTC_BarcodeNO (Barcode,SNNO) values (" & NDcode & ",'" & SNNO & "')")




set rs1=server.CreateObject("adodb.recordset")
sql= "select * from  PTC_SN where SNNO='"&SNNO&"'"
'response.Write(sql)
'
rs1.Open sql ,conn,1,3

LendNOA=rs1("LendNO")


LendNOA=LendNOA+1
'response.Write(LendNOA)
rs1("LendNO")=LendNOA

rs1.Update









end if

Next

response.Write("OK")
else



do while not rs.eof
'



AA=rs("Barcode")

ErrorName="¼ÇÂ¼ÖØ¸´!"

Set rsE = conn.Execute("INSERT INTO PTC_BarcodeError (SNNO,Barcode,ErrorName) values ('" & SNNO & "','" & AA & "','" & ErrorName & "')")



response.Write(AA+",")

rs.movenext
loop


end if


%>