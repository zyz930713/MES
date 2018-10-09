<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/Admin/CheckABC/ABCCheck.asp" -->
<!--#include virtual="/WOCF/PVS_Open.asp" -->
<%
path=request.Form("path")
query=request.Form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
paramSF=trim(request("paramSF"))
LLA=trim(request("LLA"))
ULA=trim(request("ULA"))
LLB=trim(request("LLB"))
ULB=trim(request("ULB"))
Product_Name=trim(request("Product_Name"))



SQL="select * from TempCheckSol where paramSF='"&paramSF&"' and Product_Name='"&Product_Name&"' "

rs.open SQL,conn,1,3
if rs.eof then
rs.addnew
rs("paramSF")=paramSF
rs("LLA")=LLA
rs("ULA")=ULA
rs("LLB")=LLB
rs("ULB")=ULB
rs("Product_Name")=Product_Name
rs.update
rs.close
word="Successfully save a New CheckABC config."
action="location.href='"&beforepath&"'"
else
word="CheckABC config of "&paramSF&" has existed, please input again."
action="history.back()"
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
</head>
<script language="javascript">
alert("<%=word%>");
<%=action%>;
</script>
<body>

</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->