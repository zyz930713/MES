<!--#include virtual="/WOCF/BOCF_Open.asp" -->

<%
Dcode=trim(request("Dcode"))
SNNO=trim(request("SNNO"))
codeK=""
tArr = Split(Dcode,",")  '�Զ���Ϊ�ָ�����ת��������tArr
For i = 0 To UBound(tArr)  ' UBound(tArr) ���������ж��ٸ�
NDcode=(tArr(i))           '�õ������е�ÿ��ֵ
if NDcode="'SANEND'" then  '����ǽ���������κβ���
else
set rs=server.CreateObject("adodb.recordset")
sql= "select * from PTC_BarcodeNO where SNNO='"&SNNO&"' and  Codestate='�ѽ���' and Barcode  in ("&NDcode&")" 
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
Codestate="���ڹ黹"
ReturnDate=now()
conn.Execute("update PTC_BarcodeNO set Codestate='"&Codestate&"',ReturnDate='"&ReturnDate&"' where SNNO='"&SNNO&"' and  Barcode  in ("&Dcode&")" )
response.Write("OK")
set rsJ=server.CreateObject("adodb.recordset")
sqlJsq= "select count(codestate)  as JSQ  from PTC_BarcodeNO where SNNO='"&SNNO&"' and   codestate='���ڹ黹'"
rsJ.Open sqlJsq,conn,1,3
if not(rsJ.bof and rsJ.eof) then
JSQ=rsJ("JSQ")
end if
conn.Execute("update PTC_SN set ReturnNOJSQ='"&JSQ&"' where SNNO='"&SNNO&"'")
response.Write(JSQ)
end if

%>