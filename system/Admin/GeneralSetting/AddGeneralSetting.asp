<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/NID_SEQ.asp" -->
<%
	guid=request.QueryString("id")
	Action=request.QueryString("Action")
	
	title="<span id='inner_AddData'></span>"
	SubSeriesName=request("txtSubSeriesName")
	RetestYield=request("txtRetestYield")
	'RetestRatio=request("txtRetestRatio")
	RetestRatio=0
	Remark1=request("txtRemark1")
	SlappingHoldDays=request("txtSlappingHoldDays")
	Remark2=request("txtRemark2")
	SpecialExamine1=request("txtSpecialExamine1")
	AQL1=request("txtAQL1")
	SpecialExamine2=request("txtSpecialExamine2")
	AQL2=request("txtAQL2")
	SpecialExamine3=request("txtSpecialExamine3")
	AQL3=request("txtAQL3")
	SpecialExamine4=request("txtSpecialExamine4")
	AQL4=request("txtAQL4")
	Suffix=request("txtSuffix")
	CurrentAQL=request("txtCurrentAQL")
	MaxAQL=request("txtMaxAQL")
	MinAQL=request("txtMinAQL")
	Owner=request("txtOwner")
 	Appearance=request("txtAppearance")
	Performance=request("txtPerformance")
	MoreOQC=request("MoreOQC")
    
	
	 
	 
	
	if(Action="1") then   'add a new 
		title="<span id='inner_AddData'></span>"
		SQL="select * from GENERAL_SETTING where SUBSERIESNAME='"&SubSeriesName&"' AND IsExpired='0' AND ModelName=''"		
		rs.open SQL,conn,1,3
		if(rs.recordcount=0) then
			rs.addnew
				rs("GNID")="GS"&NID_SEQ("GENERAL")
				rs("SubSeriesName")=SubSeriesName
				rs("MODELNAME")=""
				rs("RETESTYIELD")=RetestYield
				rs("RETESTRATIO")=RetestRatio
				rs("REMARK1")=REMARK1
				rs("SLAPPINGHOLDDAYS")=SLAPPINGHOLDDAYS
				rs("REMARK2")=REMARK2
				rs("SPECIALEXAMINE1")=SPECIALEXAMINE1
				rs("AQL1")=AQL1
				rs("SPECIALEXAMINE2")=SPECIALEXAMINE2
				rs("AQL2")=AQL2
				rs("SPECIALEXAMINE3")=SPECIALEXAMINE3
				rs("AQL3")=AQL3
				rs("SPECIALEXAMINE4")=SPECIALEXAMINE4
				rs("AQL4")=AQL4
				rs("SUFFIX")=SUFFIX
				rs("CURRENTAQL")=CURRENTAQL
				rs("MAXAQL")=MAXAQL
				rs("MINAQL")=MinAQL
				rs("VALIDSTARTTIME")=now()
				rs("ISEXPIRED")="0"
				rs("LASTUPDATECODE")=session("code")
				rs("LSTUPDATEDATETIME")=now()
				rs("OWNER")=Owner
				rs("APPEARANCE")=Appearance
				rs("PERFORMANCE")=Performance
				rs("MOREOQC")=MoreOQC
			rs.update
			
			set rsG=server.createobject("adodb.recordset")
			SQL="Insert into GENERAL_SETTING (GNID,SubSeriesName,MODELNAME,RETESTYIELD,RETESTRATIO,REMARK1,SLAPPINGHOLDDAYS,REMARK2,"
			SQL=SQL+"SPECIALEXAMINE1,AQL1,SPECIALEXAMINE2,AQL2,SPECIALEXAMINE3,AQL3,SPECIALEXAMINE4,AQL4,SUFFIX,CURRENTAQL,MAXAQL,"
			SQL=SQL+"MINAQL,VALIDSTARTTIME,VALIDENDTIME,ISEXPIRED,LASTUPDATECODE,LSTUPDATEDATETIME,OWNER,APPEARANCE,PERFORMANCE,MOREOQC)  " 
			SQL=SQL+"SELECT 'GS'|| lpad(GENERAL_SEQ.nextval,8,'0'),'"&SubSeriesName &"',PM.ITEM_NAME,'"&RetestYield&"','"&RetestRatio&"','"&REMARK1&"',"
			SQL=SQL+"'"&SLAPPINGHOLDDAYS&"','"&REMARK2&"','"&SPECIALEXAMINE1&"','"&AQL1&"','"&SPECIALEXAMINE2&"','"&AQL2&"','"&SPECIALEXAMINE3&"',"
			SQL=SQL+"'"&AQL3&"','"&SPECIALEXAMINE4&"','"&AQL4&"','"&SUFFIX&"','"&CURRENTAQL&"','"&MAXAQL&"','"&MinAQL&"','"&now()&"','',0,"
			SQL=SQL+"'"&session("code")&"','"&now()&"','"&OWNER&"','"&Appearance&"','"&Performance&"','"&MoreOQC&"' from PRODUCT_MODEL PM ,SUBSERIES SS "
			SQL=SQL+"WHERE PM.SERIES_ID=SS.NID AND SS.SUBSERIES_NAME ='"&SubSeriesName&"'"	
			'response.Write SQL
			'response.End()		
			rsG.open SQL,conn,1,3
			 
			
			response.write "<script>window.alert('Add one setting successfully!');location.href='GeneralSetting.asp';</script>"
		else
			response.write "<script>window.alert('The Sub Series exists already!')</script>"
		end if
	end if
	
	if Action="2" then 'get a record
		title="<span id='inner_EditData'></span>"
		SQL="Select * from GENERAL_SETTING where GNID='"&guid&"'"
		set rsU=server.createobject("adodb.recordset")
		rsU.open SQL,conn,1,3
		if rsU.recordcount>0 then
			SubSeriesName=rsU("SubSeriesName")
			RetestYield=rsU("RetestYield")
			RetestRatio=rsU("RetestRatio")
			Remark1=rsU("Remark1")
			Remark2=rsU("Remark2")
			SlappingHoldDays=rsU("SlappingHoldDays")
			SpecialExamine1=rsU("SpecialExamine1")
			AQL1=rsU("AQL1")
			SpecialExamine2=rsU("SpecialExamine2")
			AQL2=rsU("AQL2")
			SpecialExamine3=rsU("SpecialExamine3")
			AQL3=rsU("AQL3")
			SpecialExamine4=rsU("SpecialExamine4")
			AQL4=rsU("AQL4")
			Suffix=rsU("Suffix")
			CurrentAQL=rsU("CurrentAQL")
			MaxAQL=rsU("MaxAQL")
			MinAQL=rsU("MinAQL")
		 	Owner=rsU("Owner")
			Appearance=rsU("APPEARANCE")
			Performance=rsU("PERFORMANCE")
			MoreOQC=rsU("MOREOQC")
		end if
	end if
	
	if Action="3" then    'edit 
		SQL="Select * from GENERAL_SETTING where GNID='"&request.Form("txtid")&"'"
		set rsS=server.createobject("adodb.recordset")
		rsS.open SQL,conn,1,3
		if rsS.recordcount>0 then
			rsS("SubSeriesName")=SubSeriesName
			rsS("MODELNAME")=""
			if(RetestYield<>"") then
				rsS("RETESTYIELD")=RetestYield
			end if 
			if(RetestRatio<>"") then
				rsS("RETESTRATIO")=RetestRatio
			end if 
			if(REMARK1<>"") then
				rsS("REMARK1")=REMARK1
			end if 
			if(SLAPPINGHOLDDAYS<>"") then
				rsS("SLAPPINGHOLDDAYS")=SLAPPINGHOLDDAYS
			end if 
			if(REMARK2<>"") then
				rsS("REMARK2")=REMARK2
			end if 
			if(SPECIALEXAMINE1<>"") then
				rsS("SPECIALEXAMINE1")=SPECIALEXAMINE1
			end if 
			if(AQL1<>"") then
				rsS("AQL1")=AQL1
			end if 
			if(SPECIALEXAMINE2<>"") then
				rsS("SPECIALEXAMINE2")=SPECIALEXAMINE2
			end if 
			if(AQL2<>"") then
				rsS("AQL2")=AQL2
			end if 
			if(SPECIALEXAMINE3<>"") then
				rsS("SPECIALEXAMINE3")=SPECIALEXAMINE3
			end if 
			if(AQL3<>"") then
				rsS("AQL3")=AQL3
			end if 
			if(SPECIALEXAMINE4<>"") then
				rsS("SPECIALEXAMINE4")=SPECIALEXAMINE4
			end if 
			if(AQL4<>"") then
				rsS("AQL4")=AQL4
			end if 
			if(SUFFIX<>"") then
				rsS("SUFFIX")=SUFFIX
			end if 
			if(CURRENTAQL<>"") then
				rsS("CURRENTAQL")=CURRENTAQL
			end if 
			if(MAXAQL<>"") then
				rsS("MAXAQL")=MAXAQL
			end if 
			if(MinAQL<>"") then
				rsS("MINAQL")=MinAQL
			end if 
			rsS("VALIDSTARTTIME")=now()
			rsS("ISEXPIRED")="0"
			rsS("LASTUPDATECODE")=session("code")
			rsS("LSTUPDATEDATETIME")=now()
			if(Owner<>"") then
				rsS("OWNER")=Owner
			end if 
			if(Appearance<>"") then
				rsS("APPEARANCE")=Appearance
			end if 
			if(Performance<>"") then
				rsS("PERFORMANCE")=Performance
			end if 
			if(MoreOQC<>"0") then
				rsS("MOREOQC")=MoreOQC
			end if 
			rsS.update
			set rsS2=server.createobject("adodb.recordset")
			sql="update GENERAL_SETTING set LASTUPDATECODE='"+session("code")+"'"
			if (RetestYield<>"") then
				sql=sql+",RETESTYIELD='"+RetestYield+"'"
			end if 
			if (RetestRatio<>"") then
				sql=sql+",RETESTRATIO='"+RetestRatio+"'"
			end if 
			if (owner<>"") then
				sql=sql+",owner='"+Owner+"'"
			end if 
			if (REMARK1<>"") then
				sql=sql+",REMARK1='"+REMARK1+"'"
			end if 
			if (SLAPPINGHOLDDAYS<>"") then
				sql=sql+",SLAPPINGHOLDDAYS='"+SLAPPINGHOLDDAYS+"'"
			end if 
			if (REMARK2<>"") then
				sql=sql+",REMARK2='"+REMARK2+"'"
			end if 
			if (SPECIALEXAMINE1<>"") then
				sql=sql+",SPECIALEXAMINE1='"+SPECIALEXAMINE1+"'"
			end if 
			if (AQL1<>"") then
				sql=sql+",AQL1='"+AQL1+"'"
			end if 
			if (SPECIALEXAMINE2<>"") then
				sql=sql+",SPECIALEXAMINE2='"+SPECIALEXAMINE2+"'"
			end if 
			if (AQL2<>"") then
				sql=sql+",AQL2='"+AQL2+"'"
			end if 
			if (SPECIALEXAMINE3<>"") then
				sql=sql+",SPECIALEXAMINE3='"+SPECIALEXAMINE3+"'"
			end if 
			if (AQL3<>"") then
				sql=sql+",AQL3='"+AQL3+"'"
			end if 
			if (SPECIALEXAMINE4<>"") then
				sql=sql+",SPECIALEXAMINE4='"+SPECIALEXAMINE4+"'"
			end if 
			if (AQL4<>"") then
				sql=sql+",AQL4='"+AQL4+"'"
			end if 
			if (SUFFIX<>"") then
				sql=sql+",SUFFIX='"+SUFFIX+"'"
			end if 
			if (CURRENTAQL<>"") then
				sql=sql+",CURRENTAQL='"+CURRENTAQL+"'"
			end if 
			if (MAXAQL<>"") then
				sql=sql+",MAXAQL='"+MAXAQL+"'"
			end if 
			if (MinAQL<>"") then
				sql=sql+",MINAQL='"+MinAQL+"'"
			end if 
			if (Appearance<>"") then
				sql=sql+",APPEARANCE='"+Appearance+"'"
			end if 
			if (Performance<>"") then
				sql=sql+",PERFORMANCE='"+Performance+"'"
			end if 
			if (MoreOQC<>"0") then
				sql=sql+",MOREOQC='"+MoreOQC+"'"
			end if 
			sql=sql+" where SubSeriesName='"+SubSeriesName+"'"
			rsS2.open sql,conn,1,3
			response.write "<script>window.alert('Edit one setting successfully!')</script>"
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
<script language="JavaScript" src="/Admin/Part/FormCheck.js" type="text/javascript"></script>
<script language="JavaScript" src="/Language/language.js" type="text/javascript"></script>
<script>
	function SaveData()
	{
		if (document.getElementById("txtSubSeriesName")!="")
		{
	
			if(document.getElementById("Action").value=="2")
			{
			 
				document.form1.action="AddGeneralSetting.asp?Action=3"
			}
			else
			{
				document.form1.action="AddGeneralSetting.asp?Action=1"
			}
			 //alert(document.form1.action);
			document.form1.submit();
		}
		else
		{
			window.alert("Please input SubSeries Name!");
			return;
		}
		
	}
</script>
<link href="../../CSS/General.css" rel="stylesheet" type="text/css">
</head>

<body onload="language(<%=session("language")%>);">
<div id="insert_frame" style="position:absolute;visibility: hidden;"><iframe id="insert_iframe" src=""></iframe></div>
<form  method="post" name="form1"  >
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="2" class="t-c-greenCopy"><%=title%></td>
</tr>
    <tr> 
      <td height="20" colspan="2" class="t-c-greenCopy"><span id="inner_User"></span>:
        <% =session("User") %></td>
    </tr>
<tr>
  <td width="160px" height="20"><span id="td_SubSeries"></span><span class="red">*</span> </td>
  <td width="600px" height="20" class="red">
      <div align="left">
        <input name="txtSubSeriesName" type="text" id="txtSubSeriesName" size="50" value="<%=SubSeriesName%>" <% if Action="2" then response.write "readonly"end if  %>>
      </div></td>
    </tr>
	<!--
  <tr>
    <td height="20">Retest Yield</td>
    <td height="20"><input name="txtRetestYield" type="text" id="txtRetestYield" size="50" value="<%=RetestYield%>"></td>
  </tr>
  <tr>
    <td height="20">Retest Ratio</td>
    <td height="20"><input name="txtRetestRatio" type="text" id="txtRetestRatio" size="50" value="<%=RetestRatio%>"></td>
  </tr>
  <tr>
  	<td height="20">Remark</td>
	<td height="20">
		 <input name="txtRemark1" type="text" id="txtRemark1" size="95"  value="<%=Remark1%>" >
	</td>
  </tr>
  
  <tr>
    <td height="20">Slapping Hold Days</td>
    <td height="20"><input name="txtSlappingHoldDays" type="text" id="txtSlappingHoldDays" size="20" value="<%=SlappingHoldDays%>"></td>
  </tr>
   
 
  <tr>
    <td height="20">Remark</td>
    <td height="20"><input name="txtRemark2" type="text" id="txtRemark2" size="95" value="<%=Remark2%>" ></td>
  </tr>
  <tr>
    <td height="20">Special Examine 1</td>
    <td height="20"><input name="txtSpecialExamine1" type="text" id="txtSpecialExamine1" size="50" value="<%=SpecialExamine1%>"></td>
  </tr>
  <tr>
    <td height="20">AQL1</td>
    <td height="20"><input name="txtAQL1" type="text" id="txtAQL1" value="<%=AQL1%>"></td>
  </tr>
  <tr>
    <td height="20">Special Examine 2</td>
    <td height="20"><input name="txtSpecialExamine2" type="text" id="txtSpecialExamine2" size="50" value="<%=SpecialExamine2%>"></td>
  </tr>
  <tr>
    <td height="20">AQL2</td>
    <td height="20"><input name="txtAQL2" type="text" id="txtAQL2" value="<%=AQL2%>" ></td>
  </tr>
  <tr>
    <td height="20">Special Examine 3</td>
    <td height="20"><input name="txtSpecialExamine3" type="text" id="txtSpecialExamine3" size="50" value="<%=SpecialExamine3%>"></td>
  </tr>
  <tr>
    <td height="20">AQL3</td>
    <td height="20"><input name="txtAQL3" type="text" id="txtAQL3" value="<%=AQL3%>"></td>
  </tr>
  <tr>
    <td height="20">Special Examine 4</td>
    <td height="20"><input name="txtSpecialExamine4" type="text" id="txtSpecialExamine4" size="50" value="<%=SpecialExamine4%>"></td>
  </tr>
  <tr>
    <td height="20">AQL4</td>
    <td height="20"><input name="txtAQL4" type="text" id="txtAQL4" value="<%=AQL4%>"></td>
  </tr>
   <tr>
    <td height="20">Suffix</td>
    <td height="20"><input name="txtSuffix" type="text" id="txtSuffix" value="<%=Suffix%>"></td>
  </tr>
  -->
   <tr>
    <td height="20"><span id="td_CurrentAQL"></span></td>
    <td height="20"><input name="txtCurrentAQL" type="text" id="txtCurrentAQL" value="<%=CurrentAQL%>"></td>
  </tr>

   <tr>
    <td height="20"><span id="td_MaxAQL"></span></td>
    <td height="20"><input name="txtMaxAQL" type="text" id="txtMaxAQL" value="<%=MaxAQL%>"></td>
  </tr>
   <tr>
    <td height="20"><span id="td_MinAQL"></td>
    <td height="20"><input name="txtMinAQL" type="text" id="txtMinAQL" value="<%=MinAQL%>" ></td>
  </tr>
    <!--
  <tr>
    <td height="20">Owner</td>
    <td height="20"><input name="txtOwner" type="text" id="txtOwner" value="<%=Owner%>" ></td>
  </tr>
  
   <tr>
    <td height="20">Appearance</td>
    <td height="20"><input name="txtAppearance" type="text" id="txtAppearance" value="<%=Appearance%>" ></td>
  </tr>
  
   <tr>
    <td height="20">Performance</td>
    <td height="20"><input name="txtPerformance" type="text" id="txtPerformance" value="<%=Performance%>" ></td>
  </tr>
  
   <tr>
    <td height="20">Is Twice OQC</td>
    <td height="20">
		<select name="MoreOQC" id="MoreOQC">
		<option value="0">-</option>
        <option value="1">1</option>
		<option value="2">2</option>
   		 </select>
	</td>
  </tr>
   -->
 
  <tr>
    <td height="20" colspan="2"><div align="center">
      <input name="stationscount" type="hidden" id="stationscount">
	  <input type="hidden" name="Action" value="<%=Action%>"  />
	  <input type="hidden" name="txtid" value="<%=guid%>"  />
      <input name="path" type="hidden" id="path" value="<%=path%>">
      <input name="query" type="hidden" id="query" value="<%=query%>">
      <input type="button" name="btnOK" value="OK" onclick="SaveData()" >
&nbsp;
<input type="reset" name="btnReset" value="Reset">
</div></td>
    </tr>
</table>
</form>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->