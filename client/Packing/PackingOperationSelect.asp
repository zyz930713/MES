<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=application("SystemName")%></title>
</head>
<link href="/CSS/GeneralKES1.css" rel="stylesheet" type="text/css">
<link rel="Stylesheet" type="text/css" href="../Scripts/jquery-ui-1.9.2/css/ui-lightness/jquery-ui-1.9.2.custom.min.css" />

<body bgcolor="#339966">
<%
pack_type=request.Form("rad_pack_type")



action=request("action")

if action="OK"  then
   txt_pack_type=request.Form("txt_pack_type")
   txt_boxid=trim(request.Form("txt_boxid"))
   txt_customer=trim(request("txt_customer"))
   PRODUCT=trim(request("PRODUCT"))
	if txt_pack_type="" then
		response.Redirect("SelectPackType.asp")
	
	elseif txt_pack_type="TSCRAP" then
	      txt_pack_type="SCRAP"
	     if  txt_boxid<>"" then
		  response.Redirect "PackingOperationRWK.asp?txt_pack_type="&txt_pack_type&"&txt_boxid="&txt_boxid&"&txt_customer="&txt_customer
		  else
		  response.Redirect "PackingOperationRWK.asp?txt_pack_type="&txt_pack_type&"&txt_customer="&txt_customer
		 end if
	elseif txt_pack_type="SCRAP" then
	     if  txt_boxid<>"" then
		  response.Redirect "PackingOperationScrap.asp?txt_pack_type="&txt_pack_type&"&txt_boxid="&txt_boxid&"&txt_customer="&txt_customer
		  else
		  response.Redirect "PackingOperationScrap.asp?txt_pack_type="&txt_pack_type&"&txt_customer="&txt_customer
		 end if
	
	elseif txt_pack_type="IA" then
	     txt_pack_type="FG"
	     if  txt_boxid<>"" then
		  response.Redirect "PackingOperationIAFG.asp?txt_pack_type="&txt_pack_type&"&txt_boxid="&txt_boxid&"&txt_customer="&txt_customer
		  else
		  response.Redirect "PackingOperationIAFG.asp?txt_pack_type="&txt_pack_type&"&txt_customer="&txt_customer
		 end if
		 
	elseif txt_pack_type="RWKFG" then
	     txt_pack_type="FG"
	     if  txt_boxid<>"" then
		  response.Redirect "PackingOperationRWKFG1.asp?txt_pack_type="&txt_pack_type&"&txt_boxid="&txt_boxid&"&txt_customer="&txt_customer
		  else
		  response.Redirect "PackingOperationRWKFG1.asp?txt_pack_type="&txt_pack_type&"&txt_customer="&txt_customer
		 end if	 
		 
		 
		 elseif txt_pack_type="TKBFG" then
	     txt_pack_type="FG"
	     if  txt_boxid<>"" then
		  response.Redirect "PackingOperationTKBFG.asp?txt_pack_type="&txt_pack_type&"&txt_boxid="&txt_boxid&"&txt_customer="&txt_customer
		  else
		  response.Redirect "PackingOperationTKBFG.asp?txt_pack_type="&txt_pack_type&"&txt_customer="&txt_customer
		 end if	 
	 elseif txt_pack_type="CheckABC" then
	     txt_pack_type="FG"
	     if  txt_boxid<>"" then
		  response.Redirect "PackingOperationCheckABCD.asp?txt_pack_type="&txt_pack_type&"&txt_boxid="&txt_boxid&"&txt_customer="&txt_customer
		  else
		  response.Redirect "PackingOperationCheckABCD.asp?txt_pack_type="&txt_pack_type&"&txt_customer="&txt_customer
		 end if	 
	 elseif txt_pack_type="TSDFG" then
	     txt_pack_type="FG"
	     if  txt_boxid<>"" then
		  response.Redirect "PackingOperationTSDFG.asp?txt_pack_type="&txt_pack_type&"&txt_boxid="&txt_boxid&"&txt_customer="&txt_customer&"&PRODUCT="&PRODUCT
		  else
		  response.Redirect "PackingOperationTSDFG.asp?txt_pack_type="&txt_pack_type&"&txt_customer="&txt_customer&"&PRODUCT="&PRODUCT
		 end if	
	 elseif txt_pack_type="TSDWeekFG" then
	     txt_pack_type="FG"
	     if  txt_boxid<>"" then
		  response.Redirect "PackingOperationTSDWeekFG.asp?txt_pack_type="&txt_pack_type&"&txt_boxid="&txt_boxid&"&txt_customer="&txt_customer&"&PRODUCT="&PRODUCT
		  else
		  response.Redirect "PackingOperationTSDWeekFG.asp?txt_pack_type="&txt_pack_type&"&txt_customer="&txt_customer&"&PRODUCT="&PRODUCT
		 end if		
	 elseif txt_pack_type="TSDMarigoldWeekFG" then
	     txt_pack_type="FG"
	     if  txt_boxid<>"" then
		  response.Redirect "PackingOperationTSDWeekMarigold.asp?txt_pack_type="&txt_pack_type&"&txt_boxid="&txt_boxid&"&txt_customer="&txt_customer
		  else
		  response.Redirect "PackingOperationTSDWeekMarigold.asp?txt_pack_type="&txt_pack_type&"&txt_customer="&txt_customer
		 end if		  
	 elseif txt_pack_type="TSDMarigoldFG" then
	     txt_pack_type="FG"
	     if  txt_boxid<>"" then
		  response.Redirect "PackingOperationTSDFGMarigold.asp?txt_pack_type="&txt_pack_type&"&txt_boxid="&txt_boxid&"&txt_customer="&txt_customer
		  else
		  response.Redirect "PackingOperationTSDFGMarigold.asp?txt_pack_type="&txt_pack_type&"&txt_customer="&txt_customer
		 end if	
		  elseif txt_pack_type="TSDMarigoldTKB" then
	     txt_pack_type="FG"
	     if  txt_boxid<>"" then
		  response.Redirect "PackingTSDTKBFGMarigold.asp?txt_pack_type="&txt_pack_type&"&txt_boxid="&txt_boxid&"&txt_customer="&txt_customer
		  else
		  response.Redirect "PackingTSDTKBFGMarigold.asp?txt_pack_type="&txt_pack_type&"&txt_customer="&txt_customer
		 end if	 	
		  elseif txt_pack_type="TSDMarigoldRWK" then
	     txt_pack_type="FG"
	     if  txt_boxid<>"" then
		  response.Redirect "PackingTSDRWKFGMarigold.asp?txt_pack_type="&txt_pack_type&"&txt_boxid="&txt_boxid&"&txt_customer="&txt_customer
		  else
		  response.Redirect "PackingTSDRWKFGMarigold.asp?txt_pack_type="&txt_pack_type&"&txt_customer="&txt_customer
		 end if	 	 		  	
	 elseif txt_pack_type="TSDRWKFG" then
	     txt_pack_type="FG"
	     if  txt_boxid<>"" then
		  response.Redirect "PackingOperationTSDRWKFG.asp?txt_pack_type="&txt_pack_type&"&txt_boxid="&txt_boxid&"&txt_customer="&txt_customer
		  else
		  response.Redirect "PackingOperationTSDRWKFG.asp?txt_pack_type="&txt_pack_type&"&txt_customer="&txt_customer
		 end if	 
	 elseif txt_pack_type="TSDTKBFG" then
	     txt_pack_type="FG"
	     if  txt_boxid<>"" then
		  response.Redirect "PackingOperationTSDTKBFG.asp?txt_pack_type="&txt_pack_type&"&txt_boxid="&txt_boxid&"&txt_customer="&txt_customer
		  else
		  response.Redirect "PackingOperationTSDTKBFG.asp?txt_pack_type="&txt_pack_type&"&txt_customer="&txt_customer
		 end if	 	 
	 elseif txt_pack_type="SemiFG" then
	     txt_pack_type="FG"
	     if  txt_boxid<>"" then
		  response.Redirect "PackingOperationNOTango.asp?txt_pack_type="&txt_pack_type&"&txt_boxid="&txt_boxid&"&txt_customer="&txt_customer
		  else
		  response.Redirect "PackingOperationNOTango.asp?txt_pack_type="&txt_pack_type&"&txt_customer="&txt_customer
		 end if	 
	 elseif txt_pack_type="HVFG" then
	 txt_pack_type="FG"
	 	if  txt_boxid<>"" then
		  response.Redirect "PackingOperationHV.asp?txt_pack_type="&txt_pack_type&"&txt_boxid="&txt_boxid&"&txt_customer="&txt_customer
		  else
		  response.Redirect "PackingOperationHV.asp?txt_pack_type="&txt_pack_type&"&txt_customer="&txt_customer
		end if	 		 	 	 
	elseif txt_pack_type="Nbass" then
	 txt_pack_type="FG"
	 	if  txt_boxid<>"" then
		  response.Redirect "PackingOperationNbass.asp?txt_pack_type="&txt_pack_type&"&txt_boxid="&txt_boxid&"&txt_customer="&txt_customer
		  else
		  response.Redirect "PackingOperationNbass.asp?txt_pack_type="&txt_pack_type
		end if	 
	elseif txt_pack_type="797" then
	 txt_pack_type="FG"
	 	if  txt_boxid<>"" then
		  response.Redirect "PackingOperationOtherSite.asp?txt_pack_type="&txt_pack_type&"&txt_boxid="&txt_boxid&"&txt_customer="&txt_customer
		  else
		  response.Redirect "PackingOperationOtherSite.asp?txt_pack_type="&txt_pack_type&"&txt_customer="&txt_customer
		end if	 
	 elseif txt_pack_type="Exception" then
	     txt_pack_type="FG"
	     if  txt_boxid<>"" then
		  response.Redirect "PackingOperationException.asp?txt_pack_type="&txt_pack_type&"&txt_boxid="&txt_boxid&"&txt_customer="&txt_customer
		  else
		  response.Redirect "PackingOperationException.asp?txt_pack_type="&txt_pack_type&"&txt_customer="&txt_customer
		 end if	 
	end if
		 
end if



%>

<form id="form1" method="post"  action="PackingOperationSelect.asp?action=OK">
<input type="hidden" id="txt_pack_type"  name="txt_pack_type" value="<%=request("rad_pack_type")%>" />
  <table id="table1" width="800px" border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
    <thead>
      <tr>
        <td colspan="6" class="t-t-DarkBlue" align="center">Packing 输入包装箱号
			&nbsp;&nbsp;&nbsp;<a style="color:#6495ED;" href="SelectPackType.asp"></a></td>
      </tr>
      <Tr>
		<td height="20" colspan="6"><%=request.QueryString("word")%></td>
	</Tr>
      <tr align="center" >
        <td width="110">Box Id<span class="red"> *</span></td>
        <td width="207"><input type="text" id="txt_boxid" size="22" name="txt_boxid" value="" /></td>
		<td width="145">Customer 客户</td>
        <td width="106"><select id="txt_customer" name="txt_customer">
        <%if pack_type="SemiFG" then %>
         
		<%
        set rs_s=server.createobject("adodb.recordset")
        rs_s.open "select CUSTOMER_LABEL from CUSTOMER_MATERIAL where PRODUCT_TYPE in ('HV','NONTango') group by CUSTOMER_LABEL order by CUSTOMER_LABEL desc ",conn,1,3
        %>
        <%
        while not rs_s.eof%>
        <option value="<%=rs_s("CUSTOMER_LABEL")%>"><%=rs_s("CUSTOMER_LABEL")%></option>
        <%
        rs_s.movenext
        wend
        rs_s.close
        set rs_s=nothing
        %>
        <%elseif pack_type="HVFG" then 
        set rs_s=server.createobject("adodb.recordset")
        rs_s.open "select CUSTOMER_LABEL from CUSTOMER_MATERIAL where PRODUCT_TYPE in ('HV','NONTango') group by CUSTOMER_LABEL  order by CUSTOMER_LABEL desc  ",conn,1,3
        %>
        <%
        while not rs_s.eof%>
        <option value="<%=rs_s("CUSTOMER_LABEL")%>"><%=rs_s("CUSTOMER_LABEL")%></option>
        <%
        rs_s.movenext
        wend
        rs_s.close
        set rs_s=nothing
        %>
         <%elseif pack_type="Exception" then 
        set rs_s=server.createobject("adodb.recordset")
        rs_s.open "select CUSTOMER_LABEL from CUSTOMER_MATERIAL where PRODUCT_TYPE in ('HV','NONTango','Nbass') group by CUSTOMER_LABEL order by CUSTOMER_LABEL desc",conn,1,3
        %>
        <%
        while not rs_s.eof%>
        <option value="<%=rs_s("CUSTOMER_LABEL")%>"><%=rs_s("CUSTOMER_LABEL")%></option>
        <%
        rs_s.movenext
        wend
        rs_s.close
        set rs_s=nothing
        %>
        <%elseif pack_type="Nbass" then%>
        
        <%elseif pack_type="797" then
        set rs_s=server.createobject("adodb.recordset")
        rs_s.open "select CUSTOMER_LABEL from CUSTOMER_MATERIAL where PRODUCT_TYPE in ('HV','NONTango') group by CUSTOMER_LABEL  order by CUSTOMER_LABEL desc  ",conn,1,3
        %>
        <%
        while not rs_s.eof%>
        <option value="<%=rs_s("CUSTOMER_LABEL")%>"><%=rs_s("CUSTOMER_LABEL")%></option>
        <%
        rs_s.movenext
        wend
        rs_s.close
        set rs_s=nothing
        %>
         <%else%>
        <option value="Foxconn ">Foxconn</option>
        <option value="Pegatron">Pegatron</option>
        <%end if%>
        </select></td>
        <td width="107">项目        </td>
        <td width="111"><select id="PRODUCT" name="PRODUCT">
          
          <%
        set rs_s=server.createobject("adodb.recordset")
        rs_s.open "select PRODUCT from Line where  product is not null GROUP by product",conn,1,3
        %>
          <%
        while not rs_s.eof%>
          <option value="<%=rs_s("PRODUCT")%>"><%=rs_s("PRODUCT")%></option>
          <%
        rs_s.movenext
        wend
        rs_s.close
        set rs_s=nothing
        %>
        </select>
       </td>		        
      </tr>	  
    </thead>
    <tbody align="center">
    </tbody>
    <tfoot>	
      <tr>
        <td colspan="6" align="center" >&nbsp;</td>		
      </tr>
    </tfoot>
  </table>

  <center>
    <label>
    <input type="submit" name="Submit" value="提交" />
    </label>
  </center>
  <input type="hidden" id="computername" name="computername" />
</form>
</body>
</html>
<script language="JavaScript" src="/Functions/NoRefresh.js" type="text/javascript"></script>