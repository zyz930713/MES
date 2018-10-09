<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<%

action=request.QueryString("action")
if(action="2") then
	JobNumber=request("txtJobNumber")
	SheetNumber=request("txtSheetNumber")
	Comments=request("txtComments")
	action=request("txtaction")
	set rsJobMismatch=server.createobject("adodb.recordset")
	SQL="SELECT * FROM FINAL_TEST_QTY_MISMATCH WHERE 1=2 "
	rsJobMismatch.open SQL,conn,1,3
	
	if(rsJobMismatch.recordcount=0 ) then
		if(JobNumber<>"" and SheetNumber<>"" and Comments<>"") then
			rsJobMismatch.addnew
			rsJobMismatch("JOB_NUMBER")=JobNumber
			rsJobMismatch("SHEET_NUMBER")=SheetNumber
			rsJobMismatch("COMMENTS")=session("code") &":"&Comments
			rsJobMismatch.update
			
			JobNumber=""
			SheetNumber=""
			Comments=""
			response.write("Save Successfully!")
		end if 
	end if 
end if 
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script>
	function SaveData()
	{
		if(document.getElementById("txtJobNumber").value=="")
		{
			window.alert("Please input Job Number!");
			return;
		}
		 
		 if(document.getElementById("txtSheetNumber").value=="")
		{
			window.alert("Please input Sheet Number!");
			return;
		}
		if(document.getElementById("txtComments").value=="")
		{
			window.alert("Please input Comments!");
			return;
		}
		form1.action="Mismath.asp?action=2";
		form1.submit();
	}
</script>

</head>

<body>
<form method="post" name="form1" target="_self">
  <table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
    <tr>
      <td height="20" colspan="10" class="t-c-greenCopy">Add a Mismatch Job</td>
    </tr>
  
    <tr>
      <td height="20" width="15%">Job Number <span class="red">*</span> </td>
      <td height="20" class="red" colspan="9">
        <div align="left">
          <input name="txtJobNumber" type="text" id="txtJobNumber" size="30" value="<%=JobNumber%>">
      </div></td>
    </tr>
	   <tr>
      <td height="20" width="15%">Sheet Number <span class="red">*</span> </td>
      <td height="20" class="red" colspan="9">
        <div align="left">
          <input name="txtSheetNumber" type="text" id="txtSheetNumber" size="30" value="<%=SheetNumber%>">
      </div></td>
    </tr>
    <tr>
      <td height="20">Comments<span class="red">*</span></td>
      <td height="20" colspan="9"><input name="txtComments" type="text" id="txtComments" size="30" value="<%=Comments%>"></td>
    </tr>
    <tr>
      <td height="20" colspan="10"><div align="center">
          <input type="button" name="button" value="Save" width="75px" onclick="SaveData()">
</div></td>
    </tr>
  </table>
</form>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Commit.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" -->