<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
'response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->

<%
code=request.Form("txtOperatorCode")
session("code")=code

PTCstate=request("PTCstate")
box_id=request("box_id")
ISSUE_ID=request("ISSUE_ID")
ComputerName=request("computername")
district=","&GetdistrictName(ComputerName)
SQL="select * from Users where USER_CODE='"&code&"'"
set rs4=server.createobject("adodb.recordset")
rs4.open SQL,conn,1,3
if rs4.eof then	
	word="<span align='center' style='color:red;'>Operator Code of "&code&" is not permissions, please contact engineer.<br>"&code&" 不存在，请联系工程师。</span>"
	response.Redirect("Split_Box_index.asp?word="&word)
end if 

      
       set ISSUEI=server.CreateObject("adodb.recordset")
	   SQLI="SELECT * FROM ISSUE_RECORD WHERE ISSUE_ID='"+ISSUE_ID+"'"
       ISSUEI.open SQLI,conn,1,3
        if ISSUEI.eof and ISSUEI.bof then
			ISSUE_ID=""
			word="<span align='center' style='color:red;'>此异常号不存在！</span>"
			response.Redirect("Split_Box_index.asp?word="&word)	
		else 

		    CLOSE_TIME=ISSUEI("CLOSE_TIME")
			response.Write(CLOSE_TIME)
			if CLOSE_TIME<>""  then
				word="<span align='center' style='color:red;'>此异常已关闭不能进行操作，请重新登记异常操作！</span>"
				response.Redirect("Split_Box_index.asp?word="&word)	
			else
				Box_id=ISSUEI("Box_id")
				set rsGetBoxid=server.CreateObject("adodb.recordset")
				SQL="SELECT * FROM job_pack_DETAIL WHERE Box_id='"+Box_id+"'"
				rsGetBoxid.open SQL,conn,1,3
					if rsGetBoxid.eof and rsGetBoxid.bof then
					rsGetBoxid=""
					word="<span align='center' style='color:red;'>此箱号不存在，请换一个新的箱号</span>"
					response.Redirect("Split_Box_index.asp?word="&word)	
					else 
					WHREC_TIME=rsGetBoxid("WHREC_TIME")
					GET_TIME=rsGetBoxid("GET_TIME")
						if WHREC_TIME<>"" and  isnull(GET_TIME) then
						word="<span align='center' style='color:red;'>此箱号在大库房，请先领取</span>"
						response.Redirect("Split_Box_index.asp?word="&word)	
						end if 
					end if 
			end if		 
		end if 
		
		 


%>
<%
	
	
	
	function GetdistrictName(ComputerName)
		set rsdistrictName=server.CreateObject("adodb.recordset")
		SQL="SELECT * FROM COMPUTER_PRINTER_MAPPING WHERE UPPER(COMPUTER_NAME)='"+UCase(ComputerName)+"'"
	 
		rsdistrictName.open SQL,conn,1,3
	 	if(rsdistrictName.recordcount=0) then
			GetdistrictName=""
			exit function
		end if 
		    
		if (rsdistrictName(1)="") then
			GetdistrictName=""
			exit function
		end if 
		
		
		GetdistrictName=rsdistrictName(1)
	
	end function
%>
<!--#include virtual="/WOCF/BOCF_Commit.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>KEB Barcode System</title>
<link href="/CSS/GeneralKES1.css" rel="stylesheet" type="text/css">

<script language="javascript" type="text/javascript">


	function SaveData()
	{
	
	     if(confirm("确定要删除吗？"))
   { 
			document.form1.action="Packup_Save.asp?Action=DEL";
			document.form1.submit();
	}
	
	
	
	}

</script>



</head>
<body bgcolor="#339966">
<table width="80%"  border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <form id="form1" method="post" name="form1" action="Packup_Save.asp"  >
    <tr>
      <td height="20" colspan="2" class="t-t-DarkBlue"  align="center">(操作选择)</td>
    </tr>
    
	 <tr>
      <td colspan="2"><BR></td>
    </tr>
	
	
	
	
	
	<tr>
		<td><table width="95%" colspan="2" border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
          <tr>
            <td width="54%" align="center"></td>
            <td width="46%" align="center">></td>
          </tr>
         
        </table>
	  	  </td>
	</tr>
	<tr><td colspan="2">&nbsp;</td></tr>
	<tr>
		<Td colspan="2" align="center"><label>
		  
		</label>
		  
		  &nbsp;
			
			&nbsp;
			<input name="btnClose" type="button"  id="btnClose"  value="Close(关闭)" onClick="window.close();">	  </Td>
	</tr>
	 <input type="hidden" id="txtSubJobList" name="txtSubJobList" value=<%=SubJob%>>
	 <input type="hidden" id="txtOperatorCode" name="txtOperatorCode" value=<%=OperatorCode%>>
	   <input type="hidden" id="computername" name="computername" value=<%=ComputerName%>>
  </form>
</table>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->