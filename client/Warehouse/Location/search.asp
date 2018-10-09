<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
keyword=request.QueryString("names")
sql="select * from product_model where ROWNUM < = 10 and ITEM_NAME like '%"&keyword&"%'"
set rs=server.CreateObject("adodb.recordset")
rs.open sql,conn,1,1
Response.ContentType="text/xml"   
response.Write "<?xml version=""1.0"" encoding=""UTF-8"" ?>"
response.Write "<response>"
do while not rs.eof
response.Write "<content>"
response.Write "<name>"&rs("ITEM_NAME")&"</name>"
'response.Write "<userid>"&rs("userid")&"</userid>"
'response.Write "<startime>"&rs("startime")&"</startime>"
'response.Write "<endtime>"&rs("endtime")&"</endtime>"
response.Write "</content>"
rs.movenext
loop
response.Write "</response>"
rs.close
set rs=nothing
%>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
