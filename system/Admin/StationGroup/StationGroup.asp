<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Station/StationCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/Functions/GetSection.asp" -->
<!--#include virtual="/Functions/GetAction.asp" -->
<!--#include virtual="/Functions/GetStation.asp" -->
<!--#include virtual="/Functions/GetDefectCodeGroup.asp" -->
<!--#include virtual="/Functions/GetSeries.asp" -->
<!--#include virtual="/Functions/GetFamily.asp" -->
<!--#include virtual="/Functions/GetMaterialCategory.asp" -->
<!--#include virtual="/Functions/NID_SEQ.asp" -->
<%
action=request.QueryString("Action")
isEdit=request.QueryString("Action1")


STATION_GROUP_ID=request("txtStationGroupID")
StationGroupname=request("txtstationgroupname")
StationGroupChineseName=request("txtStationGroupChineseName")

if(STATION_GROUP_ID="") then
	STATION_GROUP_ID=request.QueryString("STATION_GROUP_ID")
end if 

if(isEdit="edit") then
	STATION_GROUP_ID=request.QueryString("STATION_GROUP_ID")
	set rsSTATION_GROUP=server.createobject("adodb.recordset")
	SQL="SELECT * FROM STATION_GROUP WHERE STATION_GROUP_ID='"+STATION_GROUP_ID+"' AND STATUS='1'"
	rsSTATION_GROUP.open SQL,conn,1,3
	if rsSTATION_GROUP.recordcount>0 then
		StationGroupname=rsSTATION_GROUP("STATION_GROUP_NAME")
		StationGroupChineseName=rsSTATION_GROUP("STATION_GROUP_CHINESE_NAME")
	end if 
end if 


if(action="2") then
	isEdit=request("isEdit")
	'STATION_GROUP Define
	if(isEdit="edit") then
		set rsStationGroupDisable=server.createobject("adodb.recordset")
		SQL="UPDATE STATION_GROUP SET  STATUS='0' WHERE STATION_GROUP_NAME='"+ucase(StationGroupname)+"'"
		rsStationGroupDisable.open SQL,conn,1,3
		
		set rsStationGroup=server.createobject("adodb.recordset")
		SQL="INSERT INTO STATION_GROUP"
		SQL=SQL+"(STATION_GROUP_ID,STATION_GROUP_NAME,STATION_GROUP_CHINESE_NAME,CREATE_TIME,CREATE_USER_CODE,STATUS)"
		SQL=SQL+"VALUES ('"+STATION_GROUP_ID+"','"+ucase(StationGroupname)+"','"+ucase(StationGroupChineseName)+"',SYSDATE,'"+session("code")+"','1')"
		rsStationGroup.open SQL,conn,1,3
		response.write("<script>alert('Save Successfully!');location.href='StationGroupList.asp'</script>")
	else
		STATION_GROUP_ID="SG"&NID_SEQ("STATION_GROUP_SEQ")
		SQL="SELECT * FROM STATION_GROUP WHERE STATION_GROUP_NAME='"+ucase(StationGroupname)+"' AND STATUS='1'"
		set rsStationGroupExist=server.createobject("adodb.recordset")
		rsStationGroupExist.open SQL,conn,1,3
		if rsStationGroupExist.recordcount>0 then
				response.write("<script>alert('Station Group:"+StationGroupname+" already exists!');</script>")
		else
				set rsStationGroup2=server.createobject("adodb.recordset")
				SQL="INSERT INTO STATION_GROUP"
				SQL=SQL+"(STATION_GROUP_ID,STATION_GROUP_NAME,STATION_GROUP_CHINESE_NAME,CREATE_TIME,CREATE_USER_CODE,STATUS)"
				SQL=SQL+"VALUES ('"+STATION_GROUP_ID+"','"+ucase(StationGroupname)+"','"+ucase(StationGroupChineseName)+"',SYSDATE,'"+session("code")+"','1')"
				rsStationGroup2.open SQL,conn,1,3
				response.write("<script>alert('Save Successfully!');location.href='StationGroupList.asp'</script>")
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
<script language="JavaScript" src="/Functions/FromTo.js" type="text/javascript"></script>
<script language="JavaScript" src="/Admin/Station/FormCheck.js" type="text/javascript"></script>
<script>
	function SaveData()
	{
		if(document.getElementById("txtstationgroupname").value=="")
		{
			window.alert("Please input Station Group Name!");
			return;
		}
		if(document.getElementById("txtStationGroupChineseName").value=="")
		{
			window.alert("Please input Station Group Chinese Name!");
			return;
		}
		form1.action="StationGroup.asp?action=2";
		form1.submit();
	}
</script>

</head>

<body>

<form method="post" name="form1" target="_self">
  <table width="100%"  border="2" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
    <tr>
	<%if isEdit<>"edit" then%>
      <td height="20" colspan="10" class="t-c-greenCopy">Add a New Station Group</td>
	  <%else%>
	  <td height="20" colspan="10" class="t-c-greenCopy">Edit a New Station Group</td>
	 <%end if %>
    </tr>
 
    <tr>
      <td height="20" width="15%">Station Group Name <span class="red">*</span> </td>
      <td height="20" class="red" colspan="9">
        <div align="left">
          <input name="txtstationgroupname" type="text" id="txtstationgroupname" size="30" value="<%=StationGroupname%>" <%if isEdit="edit" then response.write "readonly" end if %>>
      </div></td>
    </tr>
    <tr>
      <td height="20">Station Group Chinese Name  <span class="red">*</span></td>
      <td height="20" colspan="9"><input name="txtStationGroupChineseName" type="text" id="txtStationGroupChineseName" size="30" value="<%=StationGroupChineseName%>" <%if action="edit" then response.write "readonly" end if %>></td>
    </tr>
	
    <tr>
      <td height="20" colspan="10"><div align="center">
          <input type="button" name="button" value="Save" width="75px" onclick="SaveData()">
		   <input type="hidden" id="txtStationGroupID" name="txtStationGroupID" value="<%=STATION_GROUP_ID%>">
		    <input type="hidden" id="isEdit" name="isEdit" value="<%=isEdit%>">
&nbsp;
        <input type="reset" name="Submit7" value="Return" width="75px" onclick="location.href='StationGroupList.asp'">
</div></td>
    </tr>
  </table>
</form>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Commit.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" -->