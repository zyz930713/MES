<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
'response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->

<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<%
 Act=request("Act")
 
 if Act="OK" then
 
 Code=trim(request("Code"))
 DEFECT_CODE=request("DEFECT_CODE")
 
 
 sql="select job_number,get_sub_job_number(job_number, sheet_number) as subJob,DEFECT_CODE_ID from job_2d_code where CODE='"&Code&"'"
'response.Write(sql)
 rs.open sql,conn,1,3
 
if rs.eof then
response.Write("<script>window.alert('��ά�벻����!');document.form1.code.focus();</script>")
else

                        job_number=rs("job_number")
						set rsJ=server.createobject("adodb.recordset")
						sql="select job_number from job_master_store_pre where job_number='"&rs("job_number")&"' and instr(sub_job_numbers,'"&rs("subJob")&"')>0 "
						rsJ.open sql,conn,1,3
						if rsJ.eof then
					    response.Write("<script>window.alert('û����Ԥ���!');document.form1.code.focus();</script>") 
						end if
						rsJ.close

 
		  if isnull(rs("DEFECT_CODE_ID")) then
		  
				  set rs1=server.createobject("adodb.recordset")
				  
						  sqlD="select * from DEFECTCODE where DEFECT_CODE='"&DEFECT_CODE&"'"
						  
						  rs1.open sqlD,conn,1,3
								  if rs1.eof then
								  response.Write("<script>window.alert('ȱ�ݴ��벻����!');document.form1.code.focus();</script>")
								  else
								  NID=rs1("NID")
								  sql="update job_2d_CODE  set DEFECT_CODE_ID='"&NID&"' where CODE='"&Code&"'"
								  conn.execute(sql)
								  
								 response.Write("<script>window.alert('ȱ�ݴ�����������!');document.form1.code.focus();</script>") 
								  end if
						  else
						  DEFECT_CODE_ID=(rs("DEFECT_CODE_ID"))
						  response.Write("<script>window.alert('"&DEFECT_CODE_ID&"�Ѵ���DEFECT_CODE��');document.form1.code.focus();</script>")
						  
						  end if
				
				 end if
		 end if
%>


<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>KEB Barcode System</title>
<link href="/CSS/GeneralKES1.css" rel="stylesheet" type="text/css">




</head>

<body bgcolor="#339966"  onload= "javascript:document.all.Code.focus();">
<table width="47%"  border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <form id="form1" method="post" name="form1" action="Code_DEFECT_Code.asp?Act=OK"  >
    <tr>
      <td height="20" colspan="2" class="t-t-DarkBlue"  align="center">(��ά�벹ȱ�ݴ���)</td>
    </tr>
    
	 <tr>
      <td colspan="2"><BR>
	  
	  
	  </td>
    </tr>
	<tr>
		<td>
		<table width="65%" colspan="2" border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
			<tr>
			  <td>��ά��</td>
				<td width="80%"><input name="Code" type="text"  id="Code" value="" size="30"></td>
			</tr>
				<tr>
				<td width="20%">ȱ�ݴ���</td>
				<td><input  name="DEFECT_CODE" type="text" value="" size="30"></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr><td colspan="2">&nbsp;</td></tr>
	<tr>
		<Td colspan="2" align="center">&nbsp;&nbsp;<input type="submit" value="�ύ"></Td>
	</tr>
	 
	 <input type="hidden" id="txtOperatorCode" name="txtOperatorCode" value=<%=OperatorCode%>>
	   <input type="hidden" id="computername" name="computername" value=<%=ComputerName%>>
  </form>
</table>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Commit.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" -->