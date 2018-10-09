<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
response.Charset="GB2312"
on error resume next
keyword = Trim(request("queryString"))
cstname = Trim(request("cstname"))

if keyword <> "" then
	set rs3=server.CreateObject("adodb.recordset")
	sql="select * from product_model where ROWNUM < = 10 and ITEM_NAME like '%"&VBsUnEscape(keyword)&"%'"
	rs3.open sql,conn,1,1
	if rs3.eof and rs3.bof then
		response.Write("Not found")
	else 
		do while not rs3.eof
		%>
		<li onClick="fill('<%=rs3("ITEM_NAME")%>');"><%=rs3("ITEM_NAME")%></li>
		<%
			rs3.movenext
		loop
	end if
	rs3.close
	set rs3 = nothing
end if 

if cstname <> "" then
	set rs3=server.CreateObject("adodb.recordset")
	sql="select * from Plan_Customer_Name_Config where ROWNUM < = 10 and PART_NUMBER like '%"&VBsUnEscape(cstname)&"%'"
	rs3.open sql,conn,1,1
	if rs3.eof and rs3.bof then
		response.Write("Not found")
	else 
		do while not rs3.eof
			response.write "<li onClick='fill("&rs3("CUSTOMER_NAME")&")';>"&rs3("CUSTOMER_NAME")&"</li>"
			rs3.movenext
		loop
	end if
	rs3.close
	set rs3 = nothing
end if

Function VBsUnEscape(str)  
dim i,s,c
s=""
For i=1 to Len(str)
	c=Mid(str,i,1)
	If Mid(str,i,2)="%u" and i<=Len(str)-5 Then
		If IsNumeric("&H" & Mid(str,i+2,4)) Then
			s = s & CHRW(CInt("&H" & Mid(str,i+2,4)))
			i = i+5
		Else
			s = s & c
		End If
	ElseIf c="%" and i<=Len(str)-2 Then
		If IsNumeric("&H" & Mid(str,i+1,2)) Then
			s = s & CHRW(CInt("&H" & Mid(str,i+1,2)))
			i = i+2
		Else
			s = s & c
		End If
	Else
		s = s & c
	End If
Next
VBsUnEscape = s
End Function
%>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->