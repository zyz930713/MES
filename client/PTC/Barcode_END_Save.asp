<!--#include virtual="/WOCF/BOCF_Open.asp" -->

<%
Dcode=trim(request("Dcode"))
SNNO=trim(request("SNNO"))


set rsJ=server.CreateObject("adodb.recordset")
sqlJsq= "select count(codestate)  as JSQ   from PTC_BarcodeNO where SNNO='"&SNNO&"' and  codestate='已归还'"

rsJ.Open sqlJsq,conn,1,3
if not(rsJ.bof and rsJ.eof) then
JSQA=rsJ("JSQ")

end if



codeK=""
tArr = Split(Dcode,",")  '以逗号为分隔符，转换成数组tArr
For i = 0 To UBound(tArr)  ' UBound(tArr) 遍历数组有多少个
NDcode=(tArr(i))           '得到数组中的每个值
if NDcode="'SANEND'" then  '如果是结束命令不做任何操作
else
set rs=server.CreateObject("adodb.recordset")
sql= "select * from PTC_BarcodeNO where SNNO='"&SNNO&"' and  Codestate='正在归还' and Barcode  in ("&NDcode&")" 
rs.Open sql,conn,1,3
if rs.bof and rs.eof then
code=mid(NDcode,2,17)    '去除两边的单引号
codeK=codeK&code&","
response.Write(code+",")
end if
end if
Next
if codeK<> "" then
response.End()
else
Codestate="已归还"
EndDate=now()
conn.Execute("update PTC_BarcodeNO set Codestate='"&Codestate&"',EndDate='"&EndDate&"' where SNNO='"&SNNO&"' and  Barcode  in ("&Dcode&")" )
response.Write("OK")






set rsJ=server.CreateObject("adodb.recordset")
sqlJsq= "select count(codestate)  as JSQ   from PTC_BarcodeNO where SNNO='"&SNNO&"' and  codestate='已归还'"

rsJ.Open sqlJsq,conn,1,3
if not(rsJ.bof and rsJ.eof) then
JSQ=rsJ("JSQ")

end if


set rsQ=server.CreateObject("adodb.recordset")
sqlQ= "select * from PTC_SN where SNNO='"&SNNO&"'"

rsQ.Open sqlQ,conn,1,3
if not(rsQ.bof and rsQ.eof) then
ReturnNOJSQ=rsQ("ReturnNOJSQ")

end if

ReturnNOJSQ=cint(ReturnNOJSQ)-(cint(JSQ)-cint(JSQA))



conn.Execute("update PTC_SN set ReturnNO='"&JSQ&"',ReturnNOJSQ='"&ReturnNOJSQ&"' where SNNO='"&SNNO&"'")

response.Write(JSQ)

end if






%>