<!--#include virtual="/WOCF/BOCF_Open.asp" -->

<%
Dcode=trim(request("Dcode"))
SNNO=trim(request("SNNO"))


set rsJ=server.CreateObject("adodb.recordset")
sqlJsq= "select count(codestate)  as JSQ   from PTC_BarcodeNO where SNNO='"&SNNO&"' and  codestate='�ѹ黹'"

rsJ.Open sqlJsq,conn,1,3
if not(rsJ.bof and rsJ.eof) then
JSQA=rsJ("JSQ")

end if



codeK=""
tArr = Split(Dcode,",")  '�Զ���Ϊ�ָ�����ת��������tArr
For i = 0 To UBound(tArr)  ' UBound(tArr) ���������ж��ٸ�
NDcode=(tArr(i))           '�õ������е�ÿ��ֵ
if NDcode="'SANEND'" then  '����ǽ���������κβ���
else
set rs=server.CreateObject("adodb.recordset")
sql= "select * from PTC_BarcodeNO where SNNO='"&SNNO&"' and  Codestate='���ڹ黹' and Barcode  in ("&NDcode&")" 
rs.Open sql,conn,1,3
if rs.bof and rs.eof then
code=mid(NDcode,2,17)    'ȥ�����ߵĵ�����
codeK=codeK&code&","
response.Write(code+",")
end if
end if
Next
if codeK<> "" then
response.End()
else
Codestate="�ѹ黹"
EndDate=now()
conn.Execute("update PTC_BarcodeNO set Codestate='"&Codestate&"',EndDate='"&EndDate&"' where SNNO='"&SNNO&"' and  Barcode  in ("&Dcode&")" )
response.Write("OK")






set rsJ=server.CreateObject("adodb.recordset")
sqlJsq= "select count(codestate)  as JSQ   from PTC_BarcodeNO where SNNO='"&SNNO&"' and  codestate='�ѹ黹'"

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