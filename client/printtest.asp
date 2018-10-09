<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
set PrintCtl=server.createobject("PrintClass.PrintCtl")
'PrintName="10.6.99.87"
'PrintName="\\keb-dt303\OQC"
'ReturnCode=PrintCtl.PrintNormal(PrintName, "1","1234568", "584685", "10","333333",20,"")

'ReturnCode=PrintCtl.PrintOEMKEBBoxIDLabel("10.6.89.24","12222","111","222","111",remarks,INSPECTIONPQC)
'ReturnCode=PrintCtl.PrintKEBPEGATRONBIGCustomerLabel("10.6.88.140","1212","3233","231231","23232","2222","222","2222")

'ReturnCode=PrintCtl.PrintKEBCustomerLabel("10.6.98.140", "SPRJKEB472530A00034", "720", "677-00523", "240326300137", "0.566")
'ReturnCode=PrintCtl.PrintKEBPOTLabel("10.6.98.49", "WINSLET POT ASSY", "111121200006", "9711201309280001", "A", "2013-10-19", "3000", "HZ","HZ","HZ","1","Y1233","Y1333","ddd","dddddd")
ReturnCode=PrintCtl.PrintBoxLabel("10.6.89.24","FG37114072513001","240326300137","RWK37111429","720","","Foxconn")
ReturnCode="OK"
%>
<%=ReturnCode%>