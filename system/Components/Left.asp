<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<!--#include virtual="/Language/Components/Lan_Left.asp" -->
</head>

<body onLoad="language()">
<table width="150" border="1" cellpadding="0" cellspacing="0" bordercolorlight="#6666FF" bordercolordark="#FFFFFF">
  <tr class="t-c-greenCopy">
    <td height="20"><span style="cursor:hand" onClick="javascript:window.open('/Default.asp','_parent')"><span id="inner_Home"></span></span></td>
  </tr>
  <%tabi=1
  if instr(session("role"),",JOB_ADMINISTRATOR")<>0 or instr(session("role"),",JOB_READER")<>0 then%>
  <tr class="t-c-greenCopy">
    <td height="20"><span style="cursor:hand" onClick="tabshow(<%=tabi%>)"><img src="/Images/Treeimg/plus.gif" name="tabimg<%=tabi%>" width="9" height="9">&nbsp;<span id="inner_MonitorJobs"></span></span></td>
  </tr>
  <tbody id="tab<%=tabi%>" style="display:none">
  <tr class="lightblue-t-t">
    <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('/Job/Job/Job.asp')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_JobsList"></span></span></td>
  </tr>
  <tr class="lightblue-t-t">
    <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('/Job/SubJobs/Job.asp')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_SubJobs"></span></span></td></tr>
  <tr class="lightblue-t-t">
      <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('/Job/Store/StoreRecords/PreStoreRecords.asp')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_PreStoreRecords"></span></span></td>
  </tr>
<!--
    <tr class="lightblue-t-t">
      <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('/Job/Store/StoreRecords/StoreChangeRecords.asp')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_StoreChangeRecords"></span></span></td>
  </tr>
-->
</tr>
    <tr class="lightblue-t-t">
      <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('/Job/Scrap/ScrapRecords/PreScrapRecords.asp')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_PreScrapRecords"></span></span></td>
    </tr>
  <tr class="lightblue-t-t">
    <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('/Job/Material/Material.asp')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_JobMaterial"></span></span></td>
  </tr>
  <tr class="lightblue-t-t">
    <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('/Job/JobPriority/JobPriority.asp')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_JobPriority"></span></span></td>
  </tr>
  <%if instr(session("role"),",JOB_ADMINISTRATOR")<>0 then %> 
  <!--
  <tr class="lightblue-t-t">
    <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('/Job/Eng/GetUnit.asp')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_Get_Unit"></span></span></td>
  </tr>
  -->
 <tr class="lightblue-t-t">
    <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('/Job/Subjobs/BatchHold.asp')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_BatchHold"></span></span></td>
  </tr> 
  <tr class="lightblue-t-t">
    <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('/Job/Subjobs/BatchRelease.asp')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_BatchRelease"></span></span></td>
  </tr>
  <%end if%>
  
  <tr class="lightblue-t-t">
      <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('/Reports/Process/KanBan/HoldReport.asp')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_HoldReport"></span></span></td>
    </tr>
	<tr class="lightblue-t-t">
      <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('<%=application("KEB_BPS_System_DoNet")%>web/JOB2DCODE/JOB2DCODE.aspx?language=<%=session("language")%>')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_2DCodeReport"></span></span></td>
    </tr>
	<tr class="lightblue-t-t">
      <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('/Reports/Production/WIIReport.asp')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_WII"></span></span></td>
    </tr>
	
	<tr class="lightblue-t-t">
      <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('/Reports/ChartReport/DefectCodeDistribution.asp')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_DefectCodeDistribution"></span></span></td>
    </tr>
  </tbody>
  <%tabi=tabi+1
  end if%>
 
  <tr class="t-c-greenCopy">
    <td height="20"><span style="cursor:hand" onClick="tabshow(<%=tabi%>)"><img src="/Images/Treeimg/plus.gif" name="tabimg<%=tabi%>" width="9" height="9">&nbsp;<span id="inner_WarehouseReport"></span></span></td>
  </tr>
   <tbody id="tab<%=tabi%>" style="display:none">
   <tr class="lightblue-t-t">
      <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('<%=application("KEB_BPS_System_DoNet")%>web/JOB_PACK_DETAIL/JOB_PACK_DETAIL.aspx?language=<%=session("language")%>')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_PackBoxReport"></span></span></td>
    </tr>
	
	<tr class="lightblue-t-t">
      <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('/Reports/Production/DeliveryReport.asp')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_DeliveryReport"></span></span></td>
    </tr>
    <tr class="lightblue-t-t">
      <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('/Reports/Production/RMAReport.asp')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_RMAReport"></span></span></td>
    </tr>
    <%if instr(session("role"),",PACKING_PLAN_ADMINISTRATOR")<>0 then%>
	<tr class="lightblue-t-t">
    <td height="20"><span style="cursor:hand" onClick="redirect('/Admin/PACKING_PLAN/Packing_Plan.asp')">&nbsp;&nbsp;&nbsp;&nbsp;<img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_PackingPlan"></span></span></td>
    </tr>
     
   
	<%end if
	
	if instr(session("role"),",PACKING_PLAN_ADMINISTRATOR")<>0 then%>
	<tr class="lightblue-t-t">
    <td height="20"><span style="cursor:hand" onClick="redirect('/Admin/PACKING_PLAN_KIPC/Packing_Plan_KIPC.asp')">&nbsp;&nbsp;&nbsp;&nbsp;<img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_PackingPlanKIPC"></span></span></td>
    </tr>
	<%end if%>
    
 <!--    <tr class="lightblue-t-t">
    <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('/OBI/JobClosing.asp')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_OBIJobClosing"></span></span></td>
  </tr>    
       <tr class="lightblue-t-t">
    <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('/OBI/JobClosingScrap.asp')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_OBIJobClosingScrap"></span></span></td>
  </tr>     
  <tr class="lightblue-t-t">
    <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('/OBI/JobSysClosing.asp')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_OBIJobSysClosing"></span></span></td>
  </tr>
  
  -->
  </tbody>
  <%tabi=tabi+1
  %>
 
  <tr class="t-c-greenCopy">
    <td height="20"><span style="cursor:hand" onClick="tabshow(<%=tabi%>)"><img src="/Images/Treeimg/plus.gif" name="tabimg<%=tabi%>" width="9" height="9">&nbsp;<span id="inner_CustomizationReport"></span></span></td>
  </tr>
  <tbody id="tab<%=tabi%>" style="display:none">
	<tr class="lightblue-t-t">
	  <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('/Reports/Cycle/JobCycle/JobCycle.asp')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_JobCycle"></span></span></td>
    </tr>
	<tr class="lightblue-t-t">
	  <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('/Reports/Cycle/JobLead/JobLead.asp')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_JobLead"></span></span></td>
    </tr>
    
    <tr class="lightblue-t-t">
	  <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirectA('/Reports/Production/IntervalReportAsia.asp')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_PrdReport"></span></span></td>
    </tr>
    <tr class="lightblue-t-t">
      <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('<%=application("KEB_BPS_System_DoNet")%>web/JOB_PACK/JOB_PACK.aspx?language=<%=session("language")%>')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_CheckBoxList"></span></span></td>
    </tr>
    <tr class="lightblue-t-t">
      <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('<%=application("KEB_BPS_System_DoNet")%>web/PTC_LIST/PTC_LIST.aspx?language=<%=session("language")%>')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_PTCList"></span></span></td>
    </tr>
     <tr class="lightblue-t-t">
      <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('<%=application("KEB_BPS_System_DoNet")%>web/EMR_PACK_DETAIL/EMR_PACK_DETAIL_LIST.aspx?language=<%=session("language")%>')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_EMRList"></span></span></td>
    </tr>
      <tr class="lightblue-t-t">
      <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('<%=application("KEB_BPS_System_DoNet")%>web/EMR_DISPATCH/EMR_DISPATCH.aspx?language=<%=session("language")%>')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_EMRManage"></span></span></td>
    </tr>	
     <tr class="lightblue-t-t">
      <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('<%=application("KEB_BPS_System_DoNet")%>web/MATERIAL_HV_RECORD/MATERIAL_HV_RECORD.aspx?language=<%=session("language")%>')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_Material_Record"></span></span></td>
    </tr>	
      <%if instr(session("role"),",SUPPLIERMATERIAL_ADMINISTRATOR")<>0 then%>
	 <tr class="lightblue-t-t">
      <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('<%=application("KEB_BPS_System_DoNet")%>web/SUPPLIERMATERIAL/SUPPLIERMATERIAL_LIST.aspx?language=<%=session("language")%>')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_Supplier_Material"></span></span></td>
    </tr>	
	<%end if%>		
  </tbody>
  <%tabi=tabi+1
  %>
  <tr class="t-c-greenCopy">
    <td height="20"><span style="cursor:hand" onClick="tabshow(<%=tabi%>)"><img src="/Images/Treeimg/plus.gif" name="tabimg<%=tabi%>" width="9" height="9">&nbsp;<span id="inner_ChartReport"></span></span></td>
  </tr>
  <tbody id="tab<%=tabi%>" style="display:none">
	<tr class="lightblue-t-t">  
	  <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('/Reports/NewReport/DefectReport.asp')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_NewDefectReport"></span></span></td>
    </tr>
   <tr class="lightblue-t-t">
      <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('/Reports/NewReport/FinalReport.asp')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_NewFinalYieldReport"></span></span></td>
   </tr>
  
   <tr class="lightblue-t-t">  
	  <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('/Reports/MATERIAL/MATERIAL.asp')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_MATERIAL"></span></span></td>
    </tr>
	
	<tr class="lightblue-t-t">  
	  <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('/Reports/Yield/JobYield/JobYield.asp')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_NewJobYield"></span></span></td>
    </tr>	
	<tr class="lightblue-t-t">  
	  <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('/Reports/QAReport/OQCPPM.asp')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_OQCPPM"></span></span></td>
    </tr>
	
		 <tr class="lightblue-t-t">  
	  <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('/Reports/Production/DailyStationYield.asp')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_DailyStationYieldNew"></span></span></td>
    </tr>
		 <tr class="lightblue-t-t">  
	  <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('/Reports/Production/WeeklyYield_OP.asp')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_WeeklyYield_OP"></span></span></td>
    </tr>
<!--	
	<tr class="lightblue-t-t">  
	  <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('/Reports/NewReport/FinancialReport.asp')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_FinancialReport"></span></span></td>
    </tr>
-->
  </tbody>	
  <%
  tabi=tabi+1
  %>
  <tr class="t-c-greenCopy">
    <td height="20"><span style="cursor:hand" onClick="tabshow(<%=tabi%>)"><img src="/Images/Treeimg/plus.gif" name="tabimg<%=tabi%>" width="9" height="9">&nbsp;<span id="inner_SystemSettings"></span></span></td>
  </tr>
  <tbody id="tab<%=tabi%>" style="display:none">
  <%if instr(session("role"),",FACTORY_ADMINISTRATOR")<>0 or instr(session("role"),",FACTORY_READER")<>0 then%>
  <tr class="lightblue-t-t">
    <td height="20">&nbsp;   &nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('/Admin/Factory/Factory.asp')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_Factories"></span></span></td>
  </tr>
    <%end if
	if instr(session("role"),",SECTION_ADMINISTRATOR")<>0 or instr(session("role"),",SECTION_READER")<>0 then%>
<tr class="lightblue-t-t">
    <td height="20">&nbsp;&nbsp; &nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('/Admin/Section/Section.asp')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_Sections"></span></span></td>
  </tr>
    <%end if
	if instr(session("role"),",LINE_ADMINISTRATOR")<>0 or instr(session("role"),",LINE_READER")<>0 then%>
    <tr class="lightblue-t-t">
      <td height="20"><span style="cursor:hand" onClick="redirect('/Admin/Line/Line.asp')">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_Lines"></span></span></td>
    </tr>
	<%end if
	if instr(session("role"),",FINANCE_SERIESGROUP_ADMINISTRATOR")<>0 or instr(session("role"),",FINANCE_SERIESGROUP_READER")<>0 then%>
	<!--
    <tr class="lightblue-t-t">
      <td height="20"><span style="cursor:hand" onClick="redirect('/Admin/FinanceSeriesGroup/SeriesGroup.asp')">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_FinanceSeriesGroup"></span></span></td>
    </tr>
	<tr class="lightblue-t-t">
      <td height="20"><span style="cursor:hand" onClick="redirect('/Admin/FinanceSeries/Series.asp')">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_FinanceSeries"></span></span></td>
    </tr>
	-->
	<%end if	
	if instr(session("role"),",SERIESGROUP_ADMINISTRATOR")<>0 or instr(session("role"),",SERIESGROUP_READER")<>0 or instr(session("role"),",SERIESGROUP_CAPACITY") then%>
    <tr class="lightblue-t-t">
      <td height="20"><span style="cursor:hand" onClick="redirect('/Admin/Family/SeriesGroup.asp')">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_Family_New"></span></span></td>
    </tr>
	
	<%end if
	if instr(session("role"),",SERIESGROUP_ADMINISTRATOR")<>0 or instr(session("role"),",SERIESGROUP_READER")<>0 or instr(session("role"),",SERIESGROUP_CAPACITY") then%>
    <tr class="lightblue-t-t">
      <td height="20"><span style="cursor:hand" onClick="redirect('/Admin/Series_New/SeriesGroup.asp')">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_Series_New"></span></span></td>
    </tr>
	
	<%end if
	if instr(session("role"),",SERIESGROUP_ADMINISTRATOR")<>0 or instr(session("role"),",SERIESGROUP_READER")<>0 or instr(session("role"),",SERIESGROUP_CAPACITY") then%>
    <tr class="lightblue-t-t">
      <td height="20"><span style="cursor:hand" onClick="redirect('/Admin/SubSeries_New/Series.asp')">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_SubSeries_New"></span></span></td>
    </tr>
	<%end if
	if instr(session("role"),",MODEL_ADMINISTRATOR")<>0 or instr(session("role"),",MODEL_READER")<>0 then%>
    <tr class="lightblue-t-t">
      <td height="20"><span style="cursor:hand" onClick="redirect('/Admin/Model/Model.asp')">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_Models"></span></span></td>
    </tr>
	<%end if
	if instr(session("role"),",CUSTOMER_MATERIAL_ADMINISTRATOR")<>0 then%>
	 <tr class="lightblue-t-t">
      <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('<%=application("KEB_BPS_System_DoNet")%>web/CUSTOMER_MATERIAL/CUSTOMER_MATERIAL_LIST.aspx?language=<%=session("language")%>')">&nbsp;<img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_Customer_Material"></span></span></td>
    </tr>	
	<%end if
	if instr(session("role"),",GlueWIRE_ADMINISTRATOR")<>0 or instr(session("role"),",GlueWIRE_READER")<>0 then%>
    <tr class="lightblue-t-t">
      <td height="20"><span style="cursor:hand" onClick="redirect('/Admin/GlueWIRE/GlueWIRE_Station_List.asp')">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_GlueWIRE"></span></span></td>
    </tr>
	<%end if
	
	if instr(session("role"),",CheckABC_ADMINISTRATOR")<>0 or instr(session("role"),",CheckABC_READER")<>0 then%>
    <tr class="lightblue-t-t">
      <td height="20"><span style="cursor:hand" onClick="redirect('/Admin/CheckABC/CheckABCList.asp')">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src="/Images/Arrow.gif" width="10" height="10">&nbsp;Check ABC</span></td>
    </tr>
	<%end if
	
	if instr(session("role"),",PACKING_ADMINISTRATOR")<>0 then%>
	
    <tr class="lightblue-t-t">
      <td height="20"><span style="cursor:hand" onClick="redirect('/Admin/PACKING/PackingTest.asp')">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_PackingTest"></span></span></td>
    </tr>
	<%end if
	
	if instr(session("role"),",RPT_Daily_Target_ADMINISTRATOR")<>0 or instr(session("role"),",RPT_Daily_Target_READER")<>0 then
	%>
   <tr class="lightblue-t-t">
    <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('/Admin/RPT_Daily_Target/RPT_DAILY_TARGET.asp')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_DailyTarget"></span></span></td>
  </tr>
  <%end if
	if instr(session("role"),",PART_ADMINISTRATOR")<>0 or instr(session("role"),",PART_READER")<>0 then
	%>
   <tr class="lightblue-t-t">
    <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('/Admin/Part/Routing.asp')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_PartsRoutine_New"></span></span></td>
  </tr>
  <%end if
	if instr(session("role"),",STATION_ADMINISTRATOR")<>0 or instr(session("role"),",STATION_READER")<>0  then
	%>
   <tr class="lightblue-t-t">
    <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('/Admin/Station/station_new.asp')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_Stations_New"></span></span></td>
  </tr>  
  <!--
  <tr class="lightblue-t-t">
    <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('/Admin/StationGroup/StationGroupList.asp')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_Stations_Group"></span></span></td>
  </tr>
  
  <tr class="lightblue-t-t">
    <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('/Admin/StationGroup/StationGroupSettingList.asp')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_Stations_Group_Setting"></span></span></td>
  </tr>  
  -->
  <%end if
  	if instr(session("role"),",ACTION_ADMINISTRATOR")<>0 or instr(session("role"),",ACTION_READER")<>0  then
  %>
   <tr class="lightblue-t-t">
    <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('/Admin/Action/Action_New.asp')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_Actions_New"></span></span></td>
  </tr>
   <%end if
   if instr(session("role"),",DEFECTCODE_ADMINISTRATOR")<>0 or instr(session("role"),",DEFECTCODE_READER")<>0 then%>
   <!--
  <tr class="lightblue-t-t">
    <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('/Admin/DefectCodeGroup/DefectCodeGroup.asp')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_DefectCodeGroups"></span></span></td>
  </tr>
  --> 
   <tr class="lightblue-t-t">
    <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('/Admin/DefectCode/DefectCode_New.asp')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_DefectCodes_New"></span></span></td>
  </tr>
  <%end if
  if instr(session("role"),",GENERAL_SETTINGS_ADMINISTRATOR")<>0 then
  %>
     <tr class="lightblue-t-t">
    <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('/Admin/GeneralSetting/GeneralSetting.asp')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_GeneralSetting"></span></span></td>
  </tr>
  <%end if
  if instr(session("role"),",MACHINE_ADMINISTRATOR")<>0 or instr(session("role"),",MACHINE_READER")<>0 then%>
   <!--
  <tr class="lightblue-t-t">
    <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('/Admin/Machine/Machine.asp')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_Machines"></span></span></td>
  </tr>
  -->
  <%end if
  if instr(session("role"),",TESTDEFECTCODE_ADMINISTRATOR")<>0 or instr(session("role"),",TESTDEFECTCODE_READER")<>0 then%>
  <tr class="lightblue-t-t">
    <td height="20"><span style="cursor:hand" onClick="redirect('/Admin/TestDefectCode/TestDefectCode.asp')"> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_TestDefectCodes"></span></span></td>
  </tr>
  <%end if
  if instr(session("role"),",GROUP_ADMINISTRATOR")<>0 or instr(session("role"),",GROUP_READER")<>0 then%>

  <tr class="lightblue-t-t">
    <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('/Admin/Group/Group.asp')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_Groups"></span></span></td>    
  </tr>

  <tr class="lightblue-t-t">
    <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('/Admin/Group/HoldGroup.asp')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_HoldGroups"></span></span></td>
  </tr>
  <%end if
  if instr(session("role"),",AUTHORITY_ADMINISTRATOR")<>0 or instr(session("role"),",AUTHORITY_READER")<>0 then%>
  <tr class="lightblue-t-t">
    <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('/Admin/Authority/Operator.asp')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_Authority"></span></span></td>
  </tr>
  <%end if
  if instr(session("role"),",PART_ADMINISTRATOR")<>0 then%>
  <tr class="lightblue-t-t">
    <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('/Admin/TrayLink/TrayLink.asp')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_TrayLink"></span></span></td>
  </tr>
  <%end if%>   
  </tbody>
  <%tabi=tabi+1
  if instr(session("role"),",SYSTEM_ADMINISTRATOR")<>0 then%>
  <tr class="t-c-greenCopy">
    <td height="20"><span style="cursor:hand" onClick="tabshow(<%=tabi%>)"><img src="/Images/Treeimg/plus.gif" name="tabimg<%=tabi%>" width="9" height="9">&nbsp;<span id="inner_SystemAdministation"></span></span></td>
  </tr>
  <tbody id="tab<%=tabi%>" style="display:none">
  <!--
  <tr class="lightblue-t-t">
    <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('/System/Operator/Operator.asp')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_Operators"></span></span></td>
  </tr>
  
  <tr class="lightblue-t-t">
    <td height="20">&nbsp;&nbsp;&nbsp; &nbsp;<span style="cursor:hand" onClick="redirect('/Operators/OperatorsUpdate.asp')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_UpdateOperators"></span></span></td>
  </tr>
  -->
  <tr class="lightblue-t-t">
    <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('/System/Engineer/Engineer.asp')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_Engineers"></span></span></td>
  </tr>
  <!--
  <tr class="lightblue-t-t">
    <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('/System/Engineer/EngineerUpdate.asp')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_UpdateEngineers"></span></span></td>
  </tr>
  
  <tr class="lightblue-t-t">
    <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('/System/UserGroup/UserGroup.asp')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_UserGroups"></span></span></td>
  </tr>
  -->
  <tr class="lightblue-t-t">
    <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('/System/Role/Role.asp')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_Roles"></span></span></td>
  </tr>
    <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('/System/ComputerSet/ComputerSet.asp')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_ComputerSet"></span></span></td>
  </tr>  
   <!--
    <tr class="lightblue-t-t">
    <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('/System/ApprovalRole/Role.asp')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_ApprovalRoles"></span></span></td>
  </tr>

    <tr class="lightblue-t-t">
    <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('/System/Event/Event.asp')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_Events"></span></span></td>
  </tr>
    <tr class="lightblue-t-t">
    <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('/System/Task/Task.asp')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_Tasks"></span></span></td>
  </tr>
     <tr class="lightblue-t-t">
    <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('/System/Form/Form.asp')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_Forms"></span></span></td>
  </tr>
    <tr class="lightblue-t-t">
    <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('/System/AutoDBJob/DBJob.asp')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_DatabaseJob"></span></span></td>
  </tr>
    <tr class="lightblue-t-t">
    <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('/System/JobRecycle/JobRecycle.asp')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_JobRecycle"></span></span></td>
  </tr>
      <tr class="lightblue-t-t">
    <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('/System/SubJobRecycle/SubJobRecycle.asp')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_SubJobRecycle"></span></span></td>
  </tr>
    <tr class="lightblue-t-t">
      <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('/System/Audit/Audit.asp')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_Audit"></span></span></td>
    </tr>
	  -->
  </tbody>
  <%tabi=tabi+1
  end if
  %>
  <!--
  <tr class="t-c-greenCopy">
    <td height="20"><span style="cursor:hand" onClick="tabshow(<%=tabi%>)"><img src="/Images/Treeimg/plus.gif" name="tabimg<%=tabi%>" width="9" height="9">&nbsp;<span id="inner_Profile"></span></span></td>
  </tr>
  <tbody id="tab<%=tabi%>" style="display:none">
  <tr class="lightblue-t-t">
    <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('/Profile/Basic/Basic.asp')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_Basic"></span></span></td>
  </tr>
  <tr class="lightblue-t-t">
    <td height="20">&nbsp;&nbsp;&nbsp; &nbsp;<span style="cursor:hand" onClick="redirect('/Profile/MyTask/MyTask.asp')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_MyTask"></span></span></td>
  </tr>
  <tr class="lightblue-t-t">
    <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('/Profile/MyForm/MyForm.asp')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_MyForm"></span></span></td>
  </tr>
  <tr class="lightblue-t-t">
    <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('/Profile/MyForm/FormBoard.asp')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_FormBoard"></span></span></td>
  </tr>
    <tr class="lightblue-t-t">
    <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('/Profile/Basic/PrintPassword.asp')"><img src="/Images/Arrow.gif" width="10" height="10">&nbsp;<span id="inner_PrintPassword"></span></span></td>
  </tr>
  </tbody>

    <tr class="t-c-greenCopy">
      <td height="20"><span style="cursor:hand" onClick="tabshow(<%=tabi%>)"><img src="/Images/Treeimg/plus.gif" name="tabimg<%=tabi%>" width="9" height="9">&nbsp;<span id="inner_Help"></span></span></td>
    </tr>
	<tbody id="tab<%=tabi%>" style="display:none">
    <tr class="lightblue-t-t">
      <td height="20">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="cursor:hand" onClick="redirect('/Help/System_Settings_Manual.doc')"><img src="/Images/Arrow.gif" width="10" height="10">System Settings Manual</span></td>
    </tr>
	</tbody>
	  -->
	<%'tabi=tabi+1%>
    <tr class="t-c-greenCopy">
      <td height="20"><%if cint(session("language"))=0 then%><span style="cursor:hand" onClick="javascript:window.open('/Language.asp?language=1','_parent')">中文版</span><%else%><span style="cursor:hand" onClick="javascript:window.open('/Language.asp?language=0','_parent')">English</span><%end if%></td>
    </tr>
    <tr class="t-c-greenCopy">
    <td height="20"><span style="cursor:hand" onClick="javascript:window.open('/Logoff.asp','_parent')"><span id="inner_LogOff"></span></span></td>
  </tr>
</table>
</body>
</html>
<!--#include virtual="/Functions/TableControl.asp" -->
