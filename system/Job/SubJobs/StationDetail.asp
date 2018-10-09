<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Job/JobCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
jobnumber=request.QueryString("jobnumber")
sheetnumber=request.QueryString("sheetnumber")
jobtype=request.QueryString("jobtype")
factory=request.QueryString("factory")
part_number=request.QueryString("part_number")
part_number_id=request.QueryString("part_number_id")
station_id=request.QueryString("station_id")
if request.QueryString("repeated_sequence")<>"" then
	repeated_sequence=request.QueryString("repeated_sequence")
else
	repeated_sequence=1
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<!--#include virtual="/Language/Job/SubJobs/Lan_StationDetail.asp" -->
</head>

<body onLoad="language()">
<form name="form1" method="post" action="/Job/SubJobs/StationDetail1.asp">
  <table border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td><table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
          <tr>
          <td height="20" colspan="6" nowrap class="t-t-Borrow"><span id="inner_ActionInfo"></span></td>
          </tr>
        <tr>
          <%
SQL="Select J.*,A.ACTION_NAME,A.ACTION_CHINESE_NAME from JOB_ACTIONS J right join ACTION A on J.ACTION_ID=A.NID where J.JOB_NUMBER='"&jobnumber&"' and J.SHEET_NUMBER='"&sheetnumber&"' and J.JOB_TYPE='"&jobtype&"'and J.STATION_ID='"&station_id&"' and J.REPEATED_SEQUENCE='"&repeated_sequence&"'"
'response.Write(SQL)
rs.open SQL,conn,1,3
  if not rs.eof then
  actions_empty=false
  i=1
	 while not rs.eof
  		for j=1 to 3
			if not rs.eof then%>
          <td height="20" nowrap class="today"><%= rs("ACTION_NAME") %>&nbsp;<%= rs("ACTION_CHINESE_NAME") %>(<%=rs("ACTION_ID")%>)
              <input name="action_id<%=i%>" type="hidden" value="<%=rs("ACTION_ID")%>"></td>
          <td height="20"><input name="action_value<%=i%>" type="text" value="<%= rs("ACTION_VALUE")%>" size="15"></td>
          <%rs.movenext
		  	i=i+1
		  	else%>
          <td height="20" nowrap class="today">&nbsp;</td>
          <td height="20">&nbsp;</td>
          <%
			end if
		next%>
        </tr>
        <%
  	wend
  else
  actions_empty=true
  end if
  rs.close
  if actions_empty=true then
  	SQL="select ACTIONS_INDEX from STATION where NID='"&station_id&"'"
	rs.open SQL,conn,1,3
	if not rs.eof then
	actions_index=rs("ACTIONS_INDEX")
	end if
	rs.close
  	SQL="select * from ACTION where NID in ('"&replace(actions_index,",","','")&"')"
	rs.open SQL,conn,1,3
	if not rs.eof then
	i=1
		while not rs.eof
			for j=1 to 3
				if not rs.eof then%>
			  <td height="20" nowrap class="today"><%= rs("ACTION_NAME") %>&nbsp;<%= rs("ACTION_CHINESE_NAME") %>
				  <input name="action_id<%=i%>" type="hidden" value="<%=rs("NID")%>"></td>
			  <td height="20"><input name="action_value<%=i%>" type="text" value="" size="15"></td>
			  <%rs.movenext
				i=i+1
				else%>
			  <td height="20" nowrap class="today">&nbsp;</td>
			  <td height="20">&nbsp;</td>
			  <%
				end if
			next%>
        </tr>
        <%
		wend
	end if
	rs.close
  end if%>
      </table>
	  <input name="actionidcount" type="hidden" id="actionidcount" value="<%=i-1%>">
</td>
    </tr>
    <tr>
      <td><table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
 <%
  set rsD=server.CreateObject("adodb.recordset")
  SQL="Select STATION_ENTER_DEFECTCODE from STATION where NID='"&station_id&"'"
  rs.open SQL,conn,1,3
  if not rs.eof then
	if rs("STATION_ENTER_DEFECTCODE")<>"" then
	station_enter_defectcode=replace(rs("STATION_ENTER_DEFECTCODE"),",","','")
	end if
  end if
  rs.close
  if factory="FA00000001" then
  partwhere=" and D.APPLIED_PART_ID like '%"&part_number_id&"%' "
  end if
  SQL="Select D.NID,D.DEFECT_CODE,D.DEFECT_NAME,D.DEFECT_CHINESE_NAME,S.STATION_NAME,S.TRANSACTION_TYPE from DEFECTCODE D left join STATION S on D.STATION_ID=S.NID where D.STATION_ID in ('"&station_enter_defectcode&"')"&partwhere
 

  rs.open SQL,conn,1,3
  if not rs.eof then
  %>
          <tr>
          <td height="20" colspan="6" nowrap class="t-t-Borrow"><span id="inner_DefectcodeInfo"></span></td>
          </tr>
        <tr>
  <%
  k=1
	 while not rs.eof
  		for l=1 to 3
			if not rs.eof then
			SQLD="Select DEFECT_QUANTITY from JOB_DEFECTCODES where JOB_NUMBER='"&jobnumber&"' and SHEET_NUMBER='"&sheetnumber&"' and STATION_ID='"&station_id&"' and DEFECT_CODE_ID='"&rs("NID")&"'"
 
			rsD.open SQLD,conn,1,3
			if not rsD.eof then
				defect_quantity=rsD("DEFECT_QUANTITY")
			else
				defect_quantity=""
			end if
			rsD.close
			%>
          <td height="20" class="today" title="<%=rs("NID") %>"><%=k%>.<%= rs("DEFECT_NAME") %>&nbsp;<%= rs("DEFECT_CHINESE_NAME") %>&nbsp;(<%= rs("DEFECT_CODE") %>)
            <input name="defectcode_id<%=k%>" type="hidden" value="<%=rs("NID")%>"></td>
			<input name="trans_type<%=k%>" type="hidden" value="<%=rs("TRANSACTION_TYPE")%>"></td>
          <td height="20"><input name="defectcode_value<%=k%>" type="text" value="<%= defect_quantity%>" size="6"></td>
          <%rs.movenext
		  	k=k+1
		  	else%>
          <td height="20" nowrap class="today">&nbsp;</td>
          <td height="20">&nbsp;</td>
          <%
			end if
		next%>
        </tr>
        <%
  	wend
  end if
  rs.close
  set rsD=nothing%>
      </table>
      <input name="defectidcount" type="hidden" id="defectidcount" value="<%=k-1%>">
</td>
    </tr>
    <tr>
      <td><table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
        <tr>
          <td height="20" nowrap><div align="center">
              <input name="path" type="hidden" id="path" value="<%=path%>">
              <input name="query" type="hidden" id="query" value="<%=query%>">
              <input name="jobnumber" type="hidden" id="jobnumber" value="<%=jobnumber%>">
			  <input name="sheetnumber" type="hidden" id="sheetnumber" value="<%=sheetnumber%>">
              <input name="jobtype" type="hidden" id="jobtype" value="<%=jobtype%>">
              <input name="station_id" type="hidden" id="station_id" value="<%=station_id%>">
              <input name="repeated_sequence" type="hidden" id="repeated_sequence" value="<%=repeated_sequence%>">
              <input name="Update" type="submit" id="Update" value="Update" <%if admin=false then%>disabled<%end if%>>
&nbsp;
        <input name="Reset" type="reset" id="Reset" value="Reset">
          </div></td>
        </tr>
      </table></td>
    </tr>
  </table>
</form>
</body>
</html>
