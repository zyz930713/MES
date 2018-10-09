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
<!--#include virtual="/Functions/GetSeries.asp" -->
<!--#include virtual="/Functions/GetFamily.asp" -->
<!--#include virtual="/Functions/GetMaterialCategory.asp" -->
<!--#include virtual="/Functions/NID_SEQ.asp" -->
<%

action=request("Action")
action1=request("Action1")
STATION_GROUP_ID=request("stationgroup")
family=request("family")
Series=request("Series")
SubSeries=request("SubSeries")
model=request("model")
UVS=request("txtUVS")
'CategoryArr=cstr(replace(request.Form("toitem")," ",""))
Station_Index=replace(request.Form("toitem200")," ","")


wherestr1=" where 1=2"
wherestr2=" where 1=2"
wherestr3=" where 1=2"

if(action1="edit") then
	StationGroupMappingID=request.QueryString("MAPPING_ID")
	SQL="SELECT * FROM STATION_GROUP_UVS_SETTING WHERE MAPPING_ID='"+StationGroupMappingID+"'"
	set rsUVS2=server.createobject("adodb.recordset")
	rsUVS2.open SQL,conn,1,3
	if(rsUVS2.recordcount>0) then
		UVS=rsUVS2("UVS")
		family=rsUVS2("FAMILY_ID")
		Series=rsUVS2("SERIES_ID")
		SubSeries=rsUVS2("SUBSERIES_ID")
		model=rsUVS2("MODEL_NAME")
		STATION_GROUP_ID=rsUVS2("STATION_GROUP_ID")
	end if 
	
end if 
if not isnull(family) and family<>"" then
	 	wherestr1=" where Family_id='"&family&"'"
end if

if not isnull(Series) and Series<>"" then
	 	wherestr2=" where Series_id='"&Series&"'"
end if

if  not isnull(SubSeries) and SubSeries<>"" then
	 	wherestr3=" where Series_id='"&SubSeries&"'"
end if
	
if(action="2") then
		StationGroupMappingID=request("txtStationGroupMappingID")
		if(StationGroupMappingID="") then
			StationGroupMappingID="SM"&NID_SEQ("STATION_GROUP_MAPPING_SEQ")
		else
			STATION_GROUP_ID=request("txtStationGroupID")
		end if 
 
	
			set rsSTATION_GROUP_UVS=server.createobject("adodb.recordset")
			SQL="DELETE STATION_GROUP_UVS_SETTING WHERE MAPPING_ID='"+StationGroupMappingID+"'"
			rsSTATION_GROUP_UVS.open SQL,conn,1,3
			
			'set rsSTATION_GROUP_BOM=server.createobject("adodb.recordset")
			'SQL="DELETE STATION_GROUP_BOM_SETTING WHERE MAPPING_ID='"+StationGroupMappingID+"'"
			'rsSTATION_GROUP_BOM.open SQL,conn,1,3

			
			set rsSTATION_GROUP_STATION_MAPPING=server.createobject("adodb.recordset")
			SQL="DELETE STATION_GROUP_STATION_MAPPING WHERE MAPPING_ID='"+StationGroupMappingID+"'"
			rsSTATION_GROUP_STATION_MAPPING.open SQL,conn,1,3
	

		set rsUVS=server.createobject("adodb.recordset")
		SQL="INSERT INTO STATION_GROUP_UVS_SETTING"
		SQL=SQL+"(MAPPING_ID,STATION_GROUP_ID,FAMILY_ID,SERIES_ID,SUBSERIES_ID,MODEL_NAME,UVS)"
		SQL=SQL+"VALUES ('"+StationGroupMappingID+"','"+STATION_GROUP_ID+"','"+family+"','"+Series+"','"+SubSeries+"','"+model+"','"+UVS+"')"
		rsUVS.open SQL,conn,1,3
		
		'BOM
		'CategorySignleArr=split(CategoryArr,",")
		'for j=0 to ubound(CategorySignleArr)
		'	if(CategorySignleArr(j)<>"") then
		'		set rsBOM=server.createobject("adodb.recordset")
		'		SQL="INSERT INTO STATION_GROUP_BOM_SETTING"
		'		SQL=SQL+"(MAPPING_ID,STATION_GROUP_ID,FAMILY_ID,SERIES_ID,SUBSERIES_ID,MODEL_NAME,MATERIAL_CATEGORY_ID,MATERIAL_CATEGORY_SEQUENCE_NO)"
		'		SQL=SQL+"VALUES ('"+StationGroupMappingID+"','"+STATION_GROUP_ID+"','"+family+"','"+Series+"','"+SubSeries+"','"+model+"','"+CategorySignleArr(j)+"',"+cstr(j)+")"
		'		rsBOM.open SQL,conn,1,3
				'response.write SQL&"<br>"
		'	end if 
		'next 

	Station_IndexArr=split(Station_Index,",")
	for j=0 to ubound(Station_IndexArr)
		if(Station_IndexArr(j)<>"") then
			set rsStation=server.createobject("adodb.recordset")
			SQL="INSERT INTO STATION_GROUP_STATION_MAPPING"
			SQL=SQL+"(MAPPING_ID,STATION_GROUP_ID,FAMILY_ID,SERIES_ID,SUBSERIES_ID,MODEL_NAME,STATION_ID,STATION_SEQUENCE_NO)"
			SQL=SQL+"VALUES ('"+StationGroupMappingID+"','"+STATION_GROUP_ID+"','"+family+"','"+Series+"','"+SubSeries+"','"+model+"','"+Station_IndexArr(j)+"',"+cstr(j)+")"
			rsStation.open SQL,conn,1,3
			'response.write SQL&"<br>"
		end if 
	next 
	response.write("<script>alert('Save Successfully!');</script>")
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
	function LoadNextItem()
	{
		
		form1.action="StationGroupSetting.asp?Action=LoadNextItem"
		form1.submit();
	
	}
	
	function SaveData()
	{
		if(document.getElementById("txtUVS").value=="")
		{
			window.alert("Please input UVS!");
			return;
		}
		
		if(isNaN(document.getElementById("txtUVS").value))
		{
			alert("Please input numeric for UVS!");
			return;
		}
		
		
		//if(document.getElementById("toitem").options.length==0)
		//{
			//window.alert("Please select Material Category !");
			//return;
		//}
		if(document.getElementById("toitem200").options.length==0)
		{
			window.alert("Please select Stations for this Station Group !");
			return;
		}
		
		//for(var j=0;j<document.getElementById("toitem").options.length;j++)
//		{
//			document.getElementById("toitem").options[j].selected=true;
//		}
		
		
		for(var j=0;j<document.getElementById("toitem200").options.length;j++)
		{
			document.getElementById("toitem200").options[j].selected=true;
		}
		
		
		
		form1.action="StationGroupSetting.asp?action=2";
		form1.submit();
	}
	
	 
</script>

</head>

<body>

<form method="post" name="form1" target="_self">
  <table width="100%"  border="2" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
    <tr>
      <td height="20" colspan="10" class="t-c-greenCopy">Station Group Setting</td>
    </tr>
 
    <tr>
      <td height="20" width="15%">Station Group Name <span class="red">*</span> </td>
      <td height="20" class="red" colspan="9">
        <div align="left">
		<%
			set rsSTATION_GROUP=server.createobject("adodb.recordset")
			SQL="SELECT * FROM STATION_GROUP WHERE STATUS='1' order by STATION_GROUP_NAME"
			rsSTATION_GROUP.open SQL,conn,1,3
			
		%>
          <select name="stationgroup" id="stationgroup" style="width:110px" <%if action1="edit" then response.write "disabled='disabled'"%> >
		  	<%while not rsSTATION_GROUP.eof%>
			<option value="<%=rsSTATION_GROUP("STATION_GROUP_ID")%>" <%if STATION_GROUP_ID=rsSTATION_GROUP("STATION_GROUP_ID") then response.write "Selected" end if%>><%=rsSTATION_GROUP("STATION_GROUP_NAME")%></option>
			<%
				rsSTATION_GROUP.movenext
				wend
			%>
			 </select>  
      </div></td>
    </tr>
    <tr>
		 <td height="20">Family  Name </td>
		 <td height="20">
			<select name="family" id="family" ONCHANGE="LoadNextItem()">
			<option value="0">-- Select Family --</option>
			<%= getFamily("OPTION",family,factorywhereoutside,"","") %>
			 </select>  
		</td>
	  
	  
		<td>Series Name</td>
		 <td height="20"><select name="Series" id="Series" ONCHANGE="LoadNextItem()">
			<option value="">-- Select Series --</option>
			 <%= getSeries("OPTION",Series,wherestr1,"","") %>
		 </select>  </td>
	  
		<td>Sub Series Name</td>
		<td height="20">
			<select name="SubSeries" id="SubSeries" ONCHANGE="LoadNextItem()">
				<option value="">-- Select Sub Series --</option>
				 <%= getSubSeries("OPTION",SubSeries,wherestr2,"","") %>
			</select>  
		 </td>
	   <Td>Model </Td>
		 <td height="20" colpspan="3">
			<select name="model" id="model">
				 <option value="">-- Select Model --</option>
					<%= getModel("OPTION",model,wherestr3,"","") %>
			</select>  
		  </td>
	</tr>
	 <Td>UVS<span class="red">*</span> </Td>
  	 <td height="20" colspan="9">
	  	<input name="txtUVS" type="text" id="txtUVS" size="30" value="<%=UVS%>">
	  </td>
  </tr>
  

  	<tr>
      <td height="20" colspan="10">Station Setting</td>
    </tr> 
		
	
	<tr>
    <td height="20">Include Stations</td>
    <td height="20" colspan="9">
	<div align="center">
      <table  border="1" align="left" cellpadding="0" cellspacing="0" bordercolorlight="#73A2EE" bordercolordark="#FFFFFF">
        <tr>
          <td height="20" class="t-t-Borrow"> <div align="center">Available Stations </div></td>
          <td height="20"><div align="center">&nbsp;</div></td>
          <td height="20" class="t-t-Borrow"><div align="center">Selected Stations <span id="selectedinsert"></span></div></td>
          <td><div align="center">&nbsp;</div></td>
        </tr>
        <tr>
          <td rowspan="7">
		 	  <%
			  	SQL="select b.*,f.FACTORY_NAME from STATION_GROUP_STATION_MAPPING a, station_new b,FACTORY F "
				SQL=SQL+" where a.station_id=b.nid and b.FACTORY_ID=F.NID and A.MAPPING_ID='"+StationGroupMappingID+"' ORDER BY STATION_SEQUENCE_NO"
				set rsSaveStation=server.createobject("adodb.recordset")
				rsSaveStation.open SQL,conn,1,3
				StationIDStr=""
				while not rsSaveStation.eof
					StationIDStr=StationIDStr+"'"+rsSaveStation("nid")+"',"
					rsSaveStation.movenext
				wend
				if len(StationIDStr)>1 then
					StationIDStr=left(StationIDStr,len(StationIDStr)-1)
				else
					StationIDStr="''"
				end if 
			  %>
			  
			  
		  	<select name="fromitem200" size="10" multiple id="fromitem200">
				<%if StationIDStr<>"''" then%>
					<%= getStation_New2(true,"OPTION",""," and S.NID not in("+StationIDStr+")"," order by S.FACTORY_ID,S.STATION_NUMBER","","") %>
				<%else%>
					<%= getStation_New2(true,"OPTION",""," and 1=1"," order by S.FACTORY_ID,S.STATION_NUMBER","","") %>
				<%end if %>
			</select>
			</td>
          <td>
		  	<div align="center"> 
		  		<img src="/Images/Button_Add.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onclick="item_move(document.form1.fromitem200,document.form1.toitem200);selectedcount200()">
		  	</div>
		  </td>
          <td rowspan="7">		   
		  	<select name="toitem200" size="10" multiple id="toitem200" ondblclick="javascript:setAction();">
				<%
					if rsSaveStation.recordcount>0 then
						rsSaveStation.movefirst
						while not rsSaveStation.eof
							response.write "<option value='"&rsSaveStation("nid")&"'>"& rsSaveStation("STATION_NUMBER")&"&nbsp;"&rsSaveStation("STATION_NAME")&"&nbsp;("&rsSaveStation("STATION_CHINESE_NAME")&") - "&rsSaveStation("FACTORY_NAME")&"</option>"
							rsSaveStation.movenext
						wend
					end if 
				%>
			</select>
				
		  </td>
          <td>
		  	<div align="center">
				<img src="/Images/Button_Up.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onclick="item_up(document.form1.toitem200)"> 
			</div>
			</td>
        </tr>
        <tr>
          <td><div align="center"></div></td>
          <td><div align="center"></div></td>
        </tr>
        <tr>
          <td>
		  	<div align="center"> 
		  		<img src="/Images/Button_Remove.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onclick="item_move(document.form1.toitem200,document.form1.fromitem200);selectedcount200()">
		  	</div>
		  </td>
          <td>
		  	<div align="center"> 
				<img src="/Images/Button_Down.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onclick="item_down(document.form1.toitem200)"> 
			</div>
		  </td>
        </tr>
        <tr>
          <td><div align="center"></div></td>
          <td><div align="center"></div></td>
        </tr>
        <tr>
          <td><div align="center"> <img src="/Images/Button_Add_All.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onclick="item_all(document.form1.fromitem200,document.form1.toitem200);selectedcount200()">
		  </div></td>
          <td><div align="center"> <img src="/Images/Button_To_Top.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onclick="item_top(document.form1.toitem200)"> </div></td>
        </tr>
        <tr>
          <td><div align="center"></div></td>
          <td><div align="center"></div></td>
        </tr>
        <tr>
          <td>
		  	<div align="center">
				<img src="/Images/Button_Remove_All.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onclick="item_all(document.form1.toitem200,document.form1.fromitem200);selectedcount200()">
		  	</div>
		  </td>
          <td>
		  	<div align="center">
				<img src="/Images/Button_To_Bottom.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onclick="item_bottom(document.form1.toitem200)"> 
			</div>
			</td>
        </tr>

      </table>
    </div>
	</td>
    </tr>

    <tr>
      <td height="20" colspan="10"><div align="center">
          <input type="button" name="button" value="Save" width="75px" onclick="SaveData()">
		  <input type="hidden" id="txtStationGroupID" name="txtStationGroupID" value="<%=STATION_GROUP_ID%>">
		  <input type="hidden" id="txtStationGroupMappingID" name="txtStationGroupMappingID" value="<%=StationGroupMappingID%>">
&nbsp;
        <input type="reset" name="Submit7" value="Return" width="75px" onclick="location.href='StationGroupSettingList.asp'">
		<input type="hidden" id="Action1" name="Action1" value="<%=Action1%>">
		
</div></td>
    </tr>
  </table>
</form>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Commit.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" -->