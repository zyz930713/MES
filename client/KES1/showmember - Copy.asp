<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%

'unescape()函数

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


 on error resume next
 response.Charset="GB2312"

  dim keyword
  keyword=Trim(request.Form("queryString"))
  cstname=trim(request("cstname"))
  NO=trim(request("NO"))
  Linename=trim(request("Linename"))
  
  
   select case Linename
  
  case "8511"
     Linename="8411"
  case "8512"
   Linename="8412"
   case "8513"
   Linename="8413"
   case "8514"
   Linename="8414"
   case "8515"
   Linename="8415"
   
     case "8516"
   Linename="8416"
   
     case "8517"
   Linename="8417"
   
     case "8518"
   Linename="8418"
   
     case "5411"
	 
	 Linename="3711"
    case "5412"
	 
	 Linename="3712"
	 
	  case "5413"
	 
	 Linename="3713"
	 
	  case "5414"
	 
	 Linename="3714"
  end select
  
  
  
  
  if keyword<>"" then
  sqlA="select PRODUCT from RPT_DAILY_TARGET where Line='"&Linename&"'"
 
  rs.open sqlA,conn,1,3
  if rs.bof and rs.eof then
  else
  PRODUCT=rs("PRODUCT")
  end if
 ' response.Write(PRODUCT)
  rs.close
   set rs3=server.CreateObject("adodb.recordset")
  sql="select b.STATION_NO, b.station_description , c.description, c.supplier_name,b.item_name,b.station_description_EN from   material_station b  left join product_model c on (b.item_name= c.item_name) where b.STATION_DESCRIPTION_EN='"&PRODUCT&"'  and b.station_description like '%"&VBsUnEscape(keyword)&"%'"
   'sql="select * from (select  a.sms_id, b.station_description , c.description, c.supplier_name,a.item_name,b.station_description_EN  from material_config a, material_station b, product_model c where  a.station_no= b.station_no and a.item_name= c.item_name  and b.station_description like '%"&VBsUnEscape(keyword)&"%')  where station_description_EN ='Maple' "
'sql="select * from (select  a.sms_id, b.station_description , c.description, c.supplier_name,a.item_name  from material_config a, material_station b, product_model c where  a.station_no= b.station_no and a.item_name= c.item_name and b.station_description_EN ='Maple' ) where station_description like '%胶%'"
     'sql="select  a.sms_id, b.station_description , c.description, c.supplier_name,a.item_name  from material_config a, (select * from material_station where and station_description_EN ='Maple') b, product_model c where  a.station_no= b.station_no and a.item_name= c.item_name and b.station_description like '%"&VBsUnEscape(keyword)&"%' "
'response.Write(sql)
  rs3.open sql,conn,1,1
  if rs3.eof and rs3.bof then
  response.Write("没有该料号！")
 else 
 do while not rs3.eof
%>
<li onClick="fill('<%=rs3("station_description")%>','<%=rs3("item_name")%>','<%=rs3("description")%>','<%=rs3("STATION_NO")%>',<%=NO%>)"><%=rs3("station_description")%>--<%=rs3("description")%>--<%=rs3("supplier_name")%>--<%=rs3("item_name")%></li>
<% 
 rs3.movenext
 loop
end if
 rs3.close
 set rs3 = nothing
end if 





 %>


