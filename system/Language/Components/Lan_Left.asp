<script language="javascript">
var strHome="Home|��ҳ"
var arrHome=strHome.split("|")
//-- Job Monitor
var strMonitorJobs="Monitor Jobs|��ع���" 
var arrMonitorJobs=strMonitorJobs.split("|")
var strJobsList="Master Jobs List|�������б�" 
var arrJobsList=strJobsList.split("|")
var strSubJobs="Sub Jobs List|�ӹ����б�" 
var arrSubJobs=strSubJobs.split("|")

var strSubJobsHistory="Sub Jobs List History|��ʷ�ӹ����б�" 
var arrSubJobsHistory=strSubJobsHistory.split("|")

 
var strJobDayCycle="Job Cycle|�ӹ�����ת" 
var arrJobDayCycle=strJobDayCycle.split("|")
var strJobsProgress="Jobs Progress|��������" 
var arrJobsProgress=strJobsProgress.split("|")
var strStoreRecords="Store Records|����¼" 
var arrStoreRecords=strStoreRecords.split("|")
var strPreStoreRecords="Pre Store Records|Ԥ����¼" 
var arrPreStoreRecords=strPreStoreRecords.split("|")
var strStoreChangeRecords="Store Change Records|�������¼" 
var arrStoreChangeRecords=strStoreChangeRecords.split("|")
var strPreScrapRecords="Pre Scrap Records|Ԥ���ϼ�¼" 
var arrPreScrapRecords=strPreScrapRecords.split("|")
var strScrapRecords="Scrap Records|���ϼ�¼" 
var arrScrapRecords=strScrapRecords.split("|")
var strScrapChangeRecords="Scrap Change Records|���ϱ����¼" 
var arrScrapChangeRecords=strScrapChangeRecords.split("|")
var strJobMaterial="Material Search|���ϲ�ѯ" 
var arrJobMaterial=strJobMaterial.split("|")
var strJobMaterial2="Material Dispatch|���ϲ�ѯ" 
var arrJobMaterial2=strJobMaterial2.split("|")
var strJobAction="Action Track|����׷��" 
var arrJobAction=strJobAction.split("|")
var strErrorJobs="Error Jobs|���󹤵�" 
var arrErrorJobs=strErrorJobs.split("|")
var strEEEJobs="EEE Jobs|���鹤��" 
var arrEEEJobs=strEEEJobs.split("|")
var strShift="Shift Manage|��ο���" 
var arrShift=strShift.split("|")
var strLabour="Labour Manage|�Ͷ�������" 
var arrLabour=strLabour.split("|")
var strCustomizationReport="Customization Report|���Ʊ���" 
var arrCustomizationReport=strCustomizationReport.split("|")
var strJobCycle="Job Cycle|������ת" 
var arrJobCycle=strJobCycle.split("|")
var strJobLead="Job Lead|�������" 
var arrJobLead=strJobLead.split("|")
var strFamilyCycle="Family Cycle|������ת" 
var arrFamilyCycle=strFamilyCycle.split("|")
var strUnCloseJob="UnClose Job|δ�ع���"
var arrUnCloseJob=strUnCloseJob.split("|")
var strBatchHold="Hold Job|Hold����" 
var arrBatchHold=strBatchHold.split("|")

var strBatchRelease="Release Job|Release����" 
var arrBatchRelease=strBatchRelease.split("|") 
var strUnInstore="UnInstore Job|δ��⹤��" 
var arrUnInstore=strUnInstore.split("|") 

var strJobPriority="Job Priority|�������ȼ�" 
var arrJobPriority=strJobPriority.split("|") 

var strPiecePartAutoScrap="Piece Parts Auto Scrap|ԭ�����Զ�����" 
var arrPiecePartAutoScrap=strPiecePartAutoScrap.split("|") 


var strPiecePartManualScrap="Piece Parts Manual Scrap|ԭ�����ֶ�����" 
var arrPiecePartManualScrap=strPiecePartManualScrap.split("|") 


var strMisMatch="Final Test Mismatch Reg|��ƥ�������Ǽ�" 
var arrMisMatch=strMisMatch.split("|") 



var strSubLineJob="Sub Line Job|Sub�߹���" 
var arrSubLineJob=strSubLineJob.split("|") 

var strSubLineBOM="Sub Line BOM|Sub��BOM" 
var arrSubLineBOM=strSubLineBOM.split("|") 


var strLABORRATE_OH="Labor Rate&OH Setting|Labor Rate��OH�趨" 
var arrLABORRATE_OH=strLABORRATE_OH.split("|") 

 

//--OBI
var strOBI="ERP Integration|ERP����" 
var arrOBI=strOBI.split("|")
var strOBIJobClosing="Job Auto (Store)|�����Զ�����⣩" 
var arrOBIJobClosing=strOBIJobClosing.split("|")
var strOBIJobClosingScrap="Job Auto (Scrap)|�����Զ������ϣ�" 
var arrOBIJobClosingScrap=strOBIJobClosingScrap.split("|")
var strOBIJobClosingChange="Job Auto (Store Change)|�����Զ����������" 
var arrOBIJobClosingChange=strOBIJobClosingChange.split("|")
var strOBIJobClosingScrapChange="Job Auto (Scrap Change)|�����Զ������ϱ����" 
var arrOBIJobClosingScrapChange=strOBIJobClosingScrapChange.split("|")
var strOBIJobSysClosing="Job Auto (Error)|�����Զ�������" 
var arrOBIJobSysClosing=strOBIJobSysClosing.split("|")
//--Yield Report
var strDailyYieldReport="Daily Yield Report|ÿ�����ʱ���"
var arrDailyYieldReport=strDailyYieldReport.split("|")
var strSeriesYield="Series Yield|ϵ������" 
var arrSeriesYield=strSeriesYield.split("|")
var strJobYield="Job Yield|��������" 
var arrJobYield=strJobYield.split("|")
var strDailyFinanceYield="Daily Finance Yield|ÿ�ղ�������" 
var arrDailyFinanceYield=strDailyFinanceYield.split("|")
var strDailyStationYield="Daily Station Yield|ÿ��վ������" 
var arrDailyStationYield=strDailyStationYield.split("|")
var strDailyOperatorYield="Daily Operator Yield|ÿ�ղ���������" 
var arrDailyOperatorYield=strDailyOperatorYield.split("|")
var strDailyMachineYield="Daily Machine Yield|ÿ�ջ�������" 
var arrDailyMachineYield=strDailyMachineYield.split("|")
var strDailyLineYield="Daily Line Yield|ÿ�ղ�������" 
var arrDailyLineYield=strDailyLineYield.split("|")
var strDailyStockYield="Daily Stock Yield|ÿ���������" 
var arrDailyStockYield=strDailyStockYield.split("|")
var strPartYield="Part Yield|�ͺ�����" 
var arrPartYield=strPartYield.split("|")
var strMachineYield="Machine Yield|��������" 
var arrMachineYield=strMachineYield.split("|")
var strStationYield="Station Yield|վ������" 
var arrStationYield=strStationYield.split("|")
var strLineYield="Line Yield|�߱�����" 
var arrLineYield=strLineYield.split("|")
//--weekly yiled
var strWeeklyYieldReport="Weekly Yield Report|ÿ�����ʱ���" 
var arrWeeklyYieldReport=strWeeklyYieldReport.split("|")
var strWeeklyFinanceYield="Weekly Finance Yield|ÿ�ܲ�������" 
var arrWeeklyFinanceYield=strWeeklyFinanceYield.split("|")
var strFinalFamilyYield="Final Family Yield|���ռ�������" 
var arrFinalFamilyYield=strFinalFamilyYield.split("|")
var strFinalSeriesYield="Final Series Yield|����ϵ������" 
var arrFinalSeriesYield=strFinalSeriesYield.split("|")
var strFinalPartYield="Final Part Yield|�����ͺ�����" 
var arrFinalPartYield=strFinalPartYield.split("|")
//--Process
var strDailyLineWIP="Daily Line WIP|ÿ���߱�WIP" 
var arrDailyLineWIP=strDailyLineWIP.split("|")
var strDailyFamilyWIP="Daily Family WIP|ÿ�ռ���WIP" 
var arrDailyFamilyWIP=strDailyFamilyWIP.split("|")
var strJobInformation="Job Information|������Ϣ" 
var arrJobInformation=strJobInformation.split("|")
var strStoreDefectcode="Store Defectcode|���ȱ��" 
var arrStoreDefectcode=strStoreDefectcode.split("|")
var strProcessReport="Process Report|���̱���" 
var arrProcessReport=strProcessReport.split("|")
var strFamilyLineLost="Family Line Lost|���������ʧ" 
var arrFamilyLineLost=strFamilyLineLost.split("|")
var strFamilyScrap="Family Scrap|���屨��" 
var arrFamilyScrap=strFamilyScrap.split("|")
var strFailureRatio="Failure Ratio|ȱ�ݱ���" 
var arrFailureRatio=strFailureRatio.split("|")

var strHoldReport="Job Hold Report|����Hold����" 
var arrHoldReport=strHoldReport.split("|")
var strJobCycleTime="Job Cycle Time|������ת����" 
var arrJobCycleTime=strJobCycleTime.split("|")
var strNONEFIFO="NONE FIFO|NONE FIFO" 
var arrNONEFIFO=strNONEFIFO.split("|")
var strJobOverdue="Job Overdue Report|������ʱ����" 
var arrJobOverdue=strJobOverdue.split("|")


var strOperatorOutput="Operator Output|����������" 
var arrOperatorOutput=strOperatorOutput.split("|")
var strPartWIP="Part WIP|�ͺ�WIP" 
var arrPartWIP=strPartWIP.split("|")
var strLineWIP="Line WIP|�߱�WIP" 
var arrLineWIP=strLineWIP.split("|")
var strOutput="Output|����" 
var arrOutput=strOutput.split("|")
var strJobStore="Job Store|������ⱨ��" 
var arrJobStore=strJobStore.split("|")
var strStoreRecords="Store Records|����¼" 
var arrStoreRecords=strStoreRecords.split("|")
var strStockRecords="Store List|����嵥" 
var arrStockRecords=strStockRecords.split("|")
var strStoreChange="Store Change|�������¼" 
var arrStoreChange=strStoreChange.split("|")
var strStoreNoYield="Store No Yield|�����������" 
var arrStoreNoYield=strStoreNoYield.split("|")
var strScrapReport="Scrap Report|���ϱ���" 
var arrScrapReport=strScrapReport.split("|")
var strJobScrap="Job Scrap|��������" 
var arrJobScrap=strJobScrap.split("|")
var strScrapChange="Scrap Change|���ϱ����¼" 
var arrScrapChange=strScrapChange.split("|")
//--Chart Report
var strChartReport="Yield Report|���ʱ���" 
var arrChartReport=strChartReport.split("|")
var strChartProduction="Prodution Chart|����ͼ��" 
var arrChartProduction=strChartProduction.split("|")
var strChartWIP="WIP Chart|����Ʒͼ��" 
var arrChartWIP=strChartWIP.split("|")












var strChartFinalFamilyYieldCompare_Daily="First past yield compare Chart_Daily|First Past Yield�նԱ�ͼ��" 
var arrChartFinalFamilyYieldCompare_Daily=strChartFinalFamilyYieldCompare_Daily.split("|")

var strChartLineLost="Line Lost Chart|������ʧͼ��" 
var arrChartLineLost=strChartLineLost.split("|")
var strChartQuantityScrap="Scrap Chart (Quantity)|����ͼ�� ��������" 
var arrChartQuantityScrap=strChartQuantityScrap.split("|")


var strNewYieldReport="FPY Report|һ�����ʱ���" 
var arrNewYieldReport=strNewYieldReport.split("|")


var strNewFinalYieldReport3="Final Yield Detail Report|����������ϸ����" 
var arrNewFinalYieldReport3=strNewFinalYieldReport3.split("|")

var strNewFinalYieldReport="Final Yield  Report|�������ʱ���" 
var arrNewFinalYieldReport=strNewFinalYieldReport.split("|")


var strNewDefectReport="FPY Report|һ�����ʱ���" 
var arrNewDefectReport=strNewDefectReport.split("|")


var strNewDefectReportWeekly="Weekly FPY Report|һ�������ܱ���" 
var arrNewDefectReportWeekly=strNewDefectReportWeekly.split("|")

var strNewDefectReportWeekly2="Weekly FPY Report(Day Stack_up)|һ�������ܱ���(���ۼ�)" 
var arrNewDefectReportWeekly2=strNewDefectReportWeekly2.split("|")



var str_FPYWeeklyChart="Weekly FPY Chart Report(Day Stack_up)|һ�������ܱ���ͼ��(���ۼ�)" 
var arr_FPYWeeklyChart=str_FPYWeeklyChart.split("|")

var str_FPYDailyChart="Daily FPY Chart Report|һ�������ձ���ͼ��" 
var arr_FPYDailyChart=str_FPYDailyChart.split("|")

var str_Rework="Rework Summary Report|Rework ���౨��" 
var arr_Rework=str_Rework.split("|")

var str_ReworkSlappingYield="ReworkSlapping Yield Report|ReworkSlapping����" 
var arr_ReworkSlappingYield=str_ReworkSlappingYield.split("|")



var str_NewJobYield="Job Yield|��������" 
var arr_NewJobYield=str_NewJobYield.split("|")
 
 var str_FY_Daily_TrackingReport="FY Daily Tracking|�����ʸ��ٱ���" 
var arr_FY_Daily_TrackingReport=str_FY_Daily_TrackingReport.split("|")

 var str_GEYieldReport="GE Yield Report|GE ���ʱ���" 
var arr_GEYieldReport=str_GEYieldReport.split("|")

var str_FinancialReport="Financial Report|��������" 
var arr_FinancialReport=str_FinancialReport.split("|")

var str_StationGroupFinanceYield="Station Group Financial Yield|վ�����������" 
var arr_StationGroupFinanceYield=str_StationGroupFinanceYield.split("|")

var str_WIPGroup="WIP & Movement Report|WIP����ת����" 
var arr_WIPGroup=str_WIPGroup.split("|")

 
var str_WIPReport="Before Test WIP Report|����ǰWIP ����" 
var arr_WIPReport=str_WIPReport.split("|")

var str_NewWIP="After Test WIP Report|���Ժ�WIP����" 
var arr_NewWIP=str_NewWIP.split("|")


var str_RetestQAHistory="Retest QA History|Retest QA��ʷ����" 
var arr_RetestQAHistory=str_RetestQAHistory.split("|")



//--Tracking
var strTrackingReport="Tracking Report|׷�ٱ���" 
var arrTrackingReport=strTrackingReport.split("|")
var strJobSchedule="Job Schedule|�����ƻ�" 
var arrJobSchedule=strJobSchedule.split("|")
var strJobTracking="Job Tracking|����׷��" 
var arrJobTracking=strJobTracking.split("|")
//System Settings
var strSystemSettings="System Settings|ϵͳ����" 
var arrSystemSettings=strSystemSettings.split("|")
var strFactories="Factories|����" 
var arrFactories=strFactories.split("|")
var strSections="Sections|��������" 
var arrSections=strSections.split("|")
var strLines="Lines|�߱�" 
var arrLines=strLines.split("|")
var strFinanceSeriesGroup="Finance Series Group|����ϵ����" 
var arrFinanceSeriesGroup=strFinanceSeriesGroup.split("|")
var strFinanceSeries="Finance Series|����ϵ��" 
var arrFinanceSeries=strFinanceSeries.split("|")





var strSubSeriesNew="SubSeries|��ϵ��" 
var arrSubSeriesNew=strSubSeriesNew.split("|")
var strSeriesNew="Series|ϵ��" 
var arrSeriesNew=strSeriesNew.split("|")
var strFamilyNew="Family|����" 
var arrFamilyNew=strFamilyNew.split("|")




var strModels="Models|�ͺ�" 
var arrModels=strModels.split("|")

var strGlueWIRE="Glue&WIRE Station Config|��ˮ&ͭ�߹�λ����" 
var arrGlueWIRE=strGlueWIRE.split("|")

var strPackingTest="Packing Config|��װ����" 
var arrPackingTest=strPackingTest.split("|")

var strPackingPlan="Packing Plan|��װ�ƻ�" 
var arrPackingPlan=strPackingPlan.split("|")

var strPackingPlanKIPC="KIPC Packing Plan|KIPC ��װ�ƻ�" 
var arrPackingPlanKIPC=strPackingPlanKIPC.split("|")
 
var strPartsRoutine="Parts Routine|�Ƴ�" 
var arrPartsRoutine=strPartsRoutine.split("|")
var strStations="Stations|վ��" 
var arrStations=strStations.split("|")

var strStationGroups="Station Group|վ����" 
var arrStationGroups=strStationGroups.split("|")

var strStationGroupSettings="Station Group Setting|վ�����趨" 
var arrStationGroupSettings=strStationGroupSettings.split("|")



var strActions="Actions|����" 
var arrActions=strActions.split("|")
var strDefectCodes="Defect Codes|ȱ�ݴ���" 
var arrDefectCodes=strDefectCodes.split("|")

var strStations1="Stations|վ��" 
var arrStations1=strStations1.split("|")
var strActions1="Actions|����" 
var arrActions1=strActions1.split("|")
var strDefectCodes1="Defect Codes|ȱ�ݴ���" 
var arrDefectCodes1=strDefectCodes1.split("|")
var strPartsRoutine1="Parts Routing|�Ƴ�" 
var arrPartsRoutine1=strPartsRoutine1.split("|")

var strGeneralSetting="General Setting|һ������" 
var arrGeneralSetting=strGeneralSetting.split("|")




var strDefectCodeGroups="Defect Code Groups|ȱ�ݴ�����" 
var arrDefectCodeGroups=strDefectCodeGroups.split("|")
var strMachines="Machines|����" 
var arrMachines=strMachines.split("|")
 
var strGroups="Groups|��" 
var arrGroups=strGroups.split("|")
var arrHoldGroups=Array("Hold Groups","HoldȺ��")
var strAuthority="Authority|Ȩ��" 
var arrAuthority=strAuthority.split("|")
 
var strDual="Finance Dual|������" 
var arrDual=strDual.split("|")
var strLCD="LCD|LCD" 
var arrLCD=strLCD.split("|")
//system administration
var strSystemAdministation="System Administation|ϵͳ����" 
var arrSystemAdministation=strSystemAdministation.split("|")
var strOperators="Operators|����Ա" 
var arrOperators=strOperators.split("|")
var strUpdateOperators="Update Operators|���²���Ա��Ϣ" 
var arrUpdateOperators=strUpdateOperators.split("|")
var strEngineers="Engineers|����ʦ" 
var arrEngineers=strEngineers.split("|")
var strUpdateEngineers="Update Engineers|���¹���ʦ��Ϣ" 
var arrUpdateEngineers=strUpdateEngineers.split("|")
var strUserGroups="User Groups|�û���" 
var arrUserGroups=strUserGroups.split("|")
var strRoles="Roles|��ɫ" 
var arrRoles=strRoles.split("|")
var strApprovalRoles="Approval Roles|������ɫ" 
var arrApprovalRoles=strApprovalRoles.split("|")
var strEvents="Events|�¼�" 
var arrEvents=strEvents.split("|")
var strTasks="System Task|ϵͳ����" 
var arrTasks=strTasks.split("|")
var strForms="System Form|ϵͳ��" 
var arrForms=strForms.split("|")
var strDatabaseJob="Database Job|���ݿ�����" 
var arrDatabaseJob=strDatabaseJob.split("|")
var strJobRecycle="Master Job Recycle|������Ͱ" 
var arrJobRecycle=strJobRecycle.split("|")
var strSubJobRecycle="Sub Job Recycle|�ӹ�������Ͱ" 
var arrSubJobRecycle=strSubJobRecycle.split("|")
var strAudit="Audit|���" 
var arrAudit=strAudit.split("|")
var strHelp="Help|����" 
var arrHelp=strHelp.split("|")
var strProfile="Profile|��������" 
var arrProfile=strProfile.split("|")
var strBasic="Basic|������Ϣ" 
var arrBasic=strBasic.split("|")
var strMyTask="My Task|�ҵ��Զ�������" 
var arrMyTask=strMyTask.split("|")
var strMyForm="My Form|�ҵı�" 
var arrMyForm=strMyForm.split("|")
var strFormBoard="Form Board|������" 
var arrFormBoard=strFormBoard.split("|")
var strLogOff="Log Off|ע��" 
var arrLogOff=strLogOff.split("|")

var strPrintPassword="Print Password|��ǩ��ӡ����" 
var arrPrintPassword=strPrintPassword.split("|")

var strGetUnit="Engineer Sample|����ʦ���" 
var arrGetUnit=strGetUnit.split("|")
var strOQC="Release to OQC|ת��OQC" 
var arrOQC=strOQC.split("|")


var strChangeModel="Change Model History|����ת�ͺż�¼" 
var arrChangeModel=strChangeModel.split("|")

var strOQCPPM="OQC PPM|OQC PPM����" 
var arrOQCPPM=strOQCPPM.split("|")


var strDailyStationYieldNew="Station Daily Yield|Station������" 
var arrDailyStationYieldNew=strDailyStationYieldNew.split("|")

var strWeeklyYield_OP="Operator Weekly OutPut|����Ա�ܲ�������" 
var arrWeeklyYield_OP=strWeeklyYield_OP.split("|")

var strRetestYield="Retest Yield|Retest ���ʱ���" 
var arrRetestYield=strRetestYield.split("|")


var strSlappingRate="Slapping Rate|�Ĵ���" 
var arrSlappingRate=strSlappingRate.split("|")

var strRetestIn="Retest In|RetestͶ��" 
var arrRetestIn=strRetestIn.split("|")


var strJobStatus="Job Status|����״̬" 
var arrJobStatus=strJobStatus.split("|")


var strAutoSolder="Auto Solder Program|�Զ�Solder����" 
var arrAutoSolder=strAutoSolder.split("|")

var strJobPriority="Job Priority|Job ���ȼ�" 
var arrJobPriority=strJobPriority.split("|")

var strJobType="Job Type|Job ����" 
var arrJobType=strJobType.split("|")
var arr2DCodeReport=["2D Code Report","��ά�뱨��"]
var arrWII=["WII","WII"]
var arrPrdReport=["Interval Production Report","��������"]
var arrCheckBoxList=["Check Box List","������"]
var arrEMRList=["Pot Membrane List","Pot Ĥ���ϲ�ѯ"]
var arrPTCList=["PTC List","��Ʒ׷�ݺͿ���"]
var arrEMRManage=["EMR Dispatch Manage","���Ϸ������"]
var arrMaterial_Record=["Material Record","���ϼ�¼��ѯ"]
var arrSupplier_Material=["Supplier Material Manage","��Ӧ����������"]
var arrCustomer_Material=["Customer Lable Manage","�ͻ���ǩ����"]
var arrPackingBoxReport=["Packing Report","��װ����"]
var arrStoreReport=["Store Report","��ⱨ��"]
var arrInventoryReport=["Inventory Report","��汨��"]
var strTrayLink="TrayLink|Tray��ϵ" 
var arrTrayLink=strTrayLink.split("|")

var strDefectCodeDistribution="Defect Code Distribution|ȱ�ݴ���ֲ�" 
var arrDefectCodeDistribution=strDefectCodeDistribution.split("|")
var arrComputerSet=["Client Register","�ͻ���ע��"]
var arrDeliveryReport=["Delivery Report","��������"]
var arrDailyTarget=["Daily Target","����Ŀ������"]
var arrMATERIAL=["MATERIAL Yield","ԭ�������ʹ�ϵ��"]
var arrRMA=["RMA Report","���˱���"]
var arrWarehouseReport=["Warehouse Report","�ֿⱨ��"]
function language()
{
try{inner_Home.innerText=arrHome[<%=session("language")%>]}catch(e){}
//--Job Monitor 
try{inner_MonitorJobs.innerText=arrMonitorJobs[<%=session("language")%>]}catch(e){} 
try{inner_JobsList.innerText=arrJobsList[<%=session("language")%>]}catch(e){}

try{inner_SubJobs_History.innerText=arrSubJobsHistory[<%=session("language")%>]}catch(e){}
 

try{inner_SubJobs.innerText=arrSubJobs[<%=session("language")%>]}catch(e){}
try{inner_JobsProgress.innerText=arrJobsProgress[<%=session("language")%>]}catch(e){}
try{inner_StoreRecords.innerText=arrStoreRecords[<%=session("language")%>]}catch(e){}
try{inner_PreStoreRecords.innerText=arrPreStoreRecords[<%=session("language")%>]}catch(e){}
try{inner_StoreChangeRecords.innerText=arrStoreChangeRecords[<%=session("language")%>]}catch(e){}
try{inner_PreScrapRecords.innerText=arrPreScrapRecords[<%=session("language")%>]}catch(e){}
try{inner_ScrapRecords.innerText=arrScrapRecords[<%=session("language")%>]}catch(e){}
try{inner_ScrapChangeRecords.innerText=arrScrapChangeRecords[<%=session("language")%>]}catch(e){}
try{inner_JobMaterial.innerText=arrJobMaterial[<%=session("language")%>]}catch(e){}
try{inner_JobMaterial2.innerText=arrJobMaterial2[<%=session("language")%>]}catch(e){}
try{inner_JobAction.innerText=arrJobAction[<%=session("language")%>]}catch(e){}
try{inner_ErrorJobs.innerText=arrErrorJobs[<%=session("language")%>]}catch(e){}
try{inner_EEEJobs.innerText=arrEEEJobs[<%=session("language")%>]}catch(e){}
try{inner_Shift.innerText=arrShift[<%=session("language")%>]}catch(e){}
try{inner_Labour.innerText=arrLabour[<%=session("language")%>]}catch(e){}
try{inner_UnCloseJob.innerText=arrUnCloseJob[<%=session("language")%>]}catch(e){}
try{inner_BatchHold.innerText=arrBatchHold[<%=session("language")%>]}catch(e){}
try{inner_BatchRelease.innerText=arrBatchRelease[<%=session("language")%>]}catch(e){}
try{inner_UnInstore.innerText=arrUnInstore[<%=session("language")%>]}catch(e){}
try{inner_JobPriority.innerText=arrJobPriority[<%=session("language")%>]}catch(e){}

try{inner_PiecePartAutoScrap.innerText=arrPiecePartAutoScrap[<%=session("language")%>]}catch(e){}

try{inner_PiecePartManualScrap.innerText=arrPiecePartManualScrap[<%=session("language")%>]}catch(e){}

try{inner_MisMatch.innerText=arrMisMatch[<%=session("language")%>]}catch(e){}
 




try{inner_SubLineJob.innerText=arrSubLineJob[<%=session("language")%>]}catch(e){}

try{inner_SubLineBOM.innerText=arrSubLineBOM[<%=session("language")%>]}catch(e){}

//--OBI 
try{inner_OBI.innerText=arrOBI[<%=session("language")%>]}catch(e){} 
try{inner_OBIJobClosing.innerText=arrOBIJobClosing[<%=session("language")%>]}catch(e){}
try{inner_OBIJobClosingChange.innerText=arrOBIJobClosingChange[<%=session("language")%>]}catch(e){}
try{inner_OBIJobClosingScrap.innerText=arrOBIJobClosingScrap[<%=session("language")%>]}catch(e){}
try{inner_OBIJobClosingScrapChange.innerText=arrOBIJobClosingScrapChange[<%=session("language")%>]}catch(e){}
try{inner_OBIJobSysClosing.innerText=arrOBIJobSysClosing[<%=session("language")%>]}catch(e){}
//--Job Cycle
try{inner_CustomizationReport.innerText=arrCustomizationReport[<%=session("language")%>]}catch(e){}
try{inner_JobCycle.innerText=arrJobCycle[<%=session("language")%>]}catch(e){}
try{inner_JobLead.innerText=arrJobLead[<%=session("language")%>]}catch(e){}
try{inner_FamilyCycle.innerText=arrFamilyCycle[<%=session("language")%>]}catch(e){}
//--Weekly Yield report
try{inner_DailyYieldReport.innerText=arrDailyYieldReport[<%=session("language")%>]}catch(e){}
try{inner_SeriesYield.innerText=arrSeriesYield[<%=session("language")%>]}catch(e){}
try{inner_JobYield.innerText=arrJobYield[<%=session("language")%>]}catch(e){}
try{inner_DailyFinanceYield.innerText=arrDailyFinanceYield[<%=session("language")%>]}catch(e){}
try{inner_DailyStationYield.innerText=arrDailyStationYield[<%=session("language")%>]}catch(e){}
try{inner_DailyOperatorYield.innerText=arrDailyOperatorYield[<%=session("language")%>]}catch(e){}
try{inner_DailyMachineYield.innerText=arrDailyMachineYield[<%=session("language")%>]}catch(e){}
try{inner_DailyLineYield.innerText=arrDailyLineYield[<%=session("language")%>]}catch(e){}
try{inner_DailyStockYield.innerText=arrDailyStockYield[<%=session("language")%>]}catch(e){}
try{inner_PartYield.innerText=arrPartYield[<%=session("language")%>]}catch(e){}
try{inner_MachineYield.innerText=arrMachineYield[<%=session("language")%>]}catch(e){}
try{inner_StationYield.innerText=arrStationYield[<%=session("language")%>]}catch(e){}
try{inner_LineYield.innerText=arrLineYield[<%=session("language")%>]}catch(e){}

//weekly yield
try{inner_WeeklyYieldReport.innerText=arrWeeklyYieldReport[<%=session("language")%>]}catch(e){}
try{inner_WeeklyFinanceYield.innerText=arrWeeklyFinanceYield[<%=session("language")%>]}catch(e){}
try{inner_FinalFamilyYield.innerText=arrFinalFamilyYield[<%=session("language")%>]}catch(e){}
try{inner_FinalSeriesYield.innerText=arrFinalSeriesYield[<%=session("language")%>]}catch(e){}
try{inner_FinalPartYield.innerText=arrFinalPartYield[<%=session("language")%>]}catch(e){}

 
//--Process
try{inner_DailyLineWIP.innerText=arrDailyLineWIP[<%=session("language")%>]}catch(e){}
try{inner_DailyFamilyWIP.innerText=arrDailyFamilyWIP[<%=session("language")%>]}catch(e){}
try{inner_JobInformation.innerText=arrJobInformation[<%=session("language")%>]}catch(e){}
try{inner_StoreDefectcode.innerText=arrStoreDefectcode[<%=session("language")%>]}catch(e){}
try{inner_ProcessReport.innerText=arrProcessReport[<%=session("language")%>]}catch(e){}
try{inner_FamilyLineLost.innerText=arrFamilyLineLost[<%=session("language")%>]}catch(e){}
try{inner_FamilyScrap.innerText=arrFamilyScrap[<%=session("language")%>]}catch(e){}
try{inner_FailureRatio.innerText=arrFailureRatio[<%=session("language")%>]}catch(e){}
try{inner_OperatorOutput.innerText=arrOperatorOutput[<%=session("language")%>]}catch(e){}
try{inner_PartWIP.innerText=arrPartWIP[<%=session("language")%>]}catch(e){}
try{inner_LineWIP.innerText=arrLineWIP[<%=session("language")%>]}catch(e){}
try{inner_Output.innerText=arrOutput[<%=session("language")%>]}catch(e){}
try{inner_StoreReport.innerText=arrStoreReport[<%=session("language")%>]}catch(e){}


try{inner_JobStore.innerText=arrJobStore[<%=session("language")%>]}catch(e){}
try{inner_StockRecords.innerText=arrStockRecords[<%=session("language")%>]}catch(e){}
try{inner_StoreChange.innerText=arrStoreChange[<%=session("language")%>]}catch(e){}
try{inner_StoreNoYield.innerText=arrStoreNoYield[<%=session("language")%>]}catch(e){}
try{inner_ScrapReport.innerText=arrScrapReport[<%=session("language")%>]}catch(e){}
try{inner_JobScrap.innerText=arrJobScrap[<%=session("language")%>]}catch(e){}
try{inner_ScrapChange.innerText=arrScrapChange[<%=session("language")%>]}catch(e){}
//--Chart Report
try{inner_ChartReport.innerText=arrChartReport[<%=session("language")%>]}catch(e){}
try{inner_ChartProduction.innerText=arrChartProduction[<%=session("language")%>]}catch(e){}
try{inner_ChartWIP.innerText=arrChartWIP[<%=session("language")%>]}catch(e){}





try{inner_ChartFinalFamilyYieldCompare_Daily.innerText=arrChartFinalFamilyYieldCompare_Daily[<%=session("language")%>]}catch(e){}
try{inner_FPYWeeklyChart.innerText=arr_FPYWeeklyChart[<%=session("language")%>]}catch(e){}
try{inner_FPYDailyChart.innerText=arr_FPYDailyChart[<%=session("language")%>]}catch(e){}
 
try{inner_ReworkSlappingYield.innerText=arr_ReworkSlappingYield[<%=session("language")%>]}catch(e){}
try{inner_WIPReport.innerText=arr_WIPReport[<%=session("language")%>]}catch(e){}
try{inner_NewJobYield.innerText=arr_NewJobYield[<%=session("language")%>]}catch(e){}


try{inner_StationGroupFinanceYield.innerText=arr_FinancialReport[<%=session("language")%>]}catch(e){}
try{inner_NewWIP.innerText=arr_NewWIP[<%=session("language")%>]}catch(e){}
try{inner_RetestQAHistory.innerText=arr_RetestQAHistory[<%=session("language")%>]}catch(e){}



try{inner_ChartLineLost.innerText=arrChartLineLost[<%=session("language")%>]}catch(e){}
try{inner_ChartQuantityScrap.innerText=arrChartQuantityScrap[<%=session("language")%>]}catch(e){}

try{inner_HoldReport.innerText=arrHoldReport[<%=session("language")%>]}catch(e){}
try{inner_JobCycleTime_New.innerText=arrJobCycleTime[<%=session("language")%>]}catch(e){}
try{inner_NONEFIFO.innerText=arrNONEFIFO[<%=session("language")%>]}catch(e){}
try{inner_JobOverdue.innerText=arrJobOverdue[<%=session("language")%>]}catch(e){}


try{inner_NewYieldReport.innerText=arrNewYieldReport[<%=session("language")%>]}catch(e){}


try{inner_WIPGroup.innerText=arr_WIPGroup[<%=session("language")%>]}catch(e){}

 



try{inner_NewFinalYieldReport.innerText=arrNewFinalYieldReport[<%=session("language")%>]}catch(e){}

try{inner_NewFinalYieldReport3.innerText=arrNewFinalYieldReport3[<%=session("language")%>]}catch(e){}

try{inner_NewDefectReport.innerText=arrNewDefectReport[<%=session("language")%>]}catch(e){}

try{inner_NewDefectReportWeekly.innerText=arrNewDefectReportWeekly[<%=session("language")%>]}catch(e){}
try{inner_NewDefectReportWeekly2.innerText=arrNewDefectReportWeekly2[<%=session("language")%>]}catch(e){}

 
 try{inner_FY_Daily_TrackingReport.innerText=arr_FY_Daily_TrackingReport[<%=session("language")%>]}catch(e){}

try{inner_GeYieldReport.innerText=arr_GEYieldReport[<%=session("language")%>]}catch(e){}
try{inner_FinancialReport.innerText=arr_FinancialReport[<%=session("language")%>]}catch(e){}

//--Tracking Report
try{inner_TrackingReport.innerText=arrTrackingReport[<%=session("language")%>]}catch(e){}
try{inner_JobSchedule.innerText=arrJobSchedule[<%=session("language")%>]}catch(e){}
try{inner_JobTracking.innerText=arrJobTracking[<%=session("language")%>]}catch(e){}
//--System Settings
try{inner_SystemSettings.innerText=arrSystemSettings[<%=session("language")%>]}catch(e){}
try{inner_Factories.innerText=arrFactories[<%=session("language")%>]}catch(e){}
try{inner_Sections.innerText=arrSections[<%=session("language")%>]}catch(e){}
try{inner_Lines.innerText=arrLines[<%=session("language")%>]}catch(e){}
try{inner_FinanceSeriesGroup.innerText=arrFinanceSeriesGroup[<%=session("language")%>]}catch(e){}
try{inner_FinanceSeries.innerText=arrFinanceSeries[<%=session("language")%>]}catch(e){}


try{inner_SubSeries_New.innerText=arrSubSeriesNew[<%=session("language")%>]}catch(e){}
try{inner_Series_New.innerText=arrSeriesNew[<%=session("language")%>]}catch(e){}
try{inner_Family_New.innerText=arrFamilyNew[<%=session("language")%>]}catch(e){}
try{inner_Rework.innerText=arr_Rework[<%=session("language")%>]}catch(e){}

try{inner_Models.innerText=arrModels[<%=session("language")%>]}catch(e){}
try{inner_GlueWIRE.innerText=arrGlueWIRE[<%=session("language")%>]}catch(e){}
try{inner_PackingTest.innerText=arrPackingTest[<%=session("language")%>]}catch(e){}
try{inner_PackingPlan.innerText=arrPackingPlan[<%=session("language")%>]}catch(e){}
try{inner_PackingPlanKIPC.innerText=arrPackingPlanKIPC[<%=session("language")%>]}catch(e){}
try{inner_Stations_New.innerText=arrStations1[<%=session("language")%>]}catch(e){}

try{inner_Stations_Group.innerText=arrStationGroups[<%=session("language")%>]}catch(e){}

try{inner_Stations_Group_Setting.innerText=arrStationGroupSettings[<%=session("language")%>]}catch(e){}



try{inner_Actions_New.innerText=arrActions1[<%=session("language")%>]}catch(e){}
try{inner_DefectCodes_New.innerText=arrDefectCodes1[<%=session("language")%>]}catch(e){}
try{inner_PartsRoutine_New.innerText=arrPartsRoutine1[<%=session("language")%>]}catch(e){}

try{inner_GeneralSetting.innerText=arrGeneralSetting[<%=session("language")%>]}catch(e){}


try{inner_DefectCodeGroups.innerText=arrDefectCodeGroups[<%=session("language")%>]}catch(e){}
try{inner_Machines.innerText=arrMachines[<%=session("language")%>]}catch(e){}
 
try{inner_Groups.innerText=arrGroups[<%=session("language")%>]
    inner_HoldGroups.innerText=arrHoldGroups[<%=session("language")%>]
}catch(e){}
try{inner_Authority.innerText=arrAuthority[<%=session("language")%>]}catch(e){}
 
try{inner_Dual.innerText=arrDual[<%=session("language")%>]}catch(e){}
try{inner_LCD.innerText=arrLCD[<%=session("language")%>]}catch(e){}
//System Administration
try{inner_SystemAdministation.innerText=arrSystemAdministation[<%=session("language")%>]}catch(e){}
try{inner_Operators.innerText=arrOperators[<%=session("language")%>]}catch(e){}
try{inner_UpdateOperators.innerText=arrUpdateOperators[<%=session("language")%>]}catch(e){}
try{inner_Engineers.innerText=arrEngineers[<%=session("language")%>]}catch(e){}
try{inner_UpdateEngineers.innerText=arrUpdateEngineers[<%=session("language")%>]}catch(e){}
try{inner_UserGroups.innerText=arrUserGroups[<%=session("language")%>]}catch(e){}
try{inner_Roles.innerText=arrRoles[<%=session("language")%>]}catch(e){}
try{inner_ApprovalRoles.innerText=arrApprovalRoles[<%=session("language")%>]}catch(e){}
try{inner_Events.innerText=arrEvents[<%=session("language")%>]}catch(e){}
try{inner_Tasks.innerText=arrTasks[<%=session("language")%>]}catch(e){}
try{inner_Forms.innerText=arrForms[<%=session("language")%>]}catch(e){}
try{inner_DatabaseJob.innerText=arrDatabaseJob[<%=session("language")%>]}catch(e){}
try{inner_JobRecycle.innerText=arrJobRecycle[<%=session("language")%>]}catch(e){}
try{inner_SubJobRecycle.innerText=arrSubJobRecycle[<%=session("language")%>]}catch(e){}
try{inner_Audit.innerText=arrAudit[<%=session("language")%>]}catch(e){}
try{inner_Help.innerText=arrHelp[<%=session("language")%>]}catch(e){}
try{inner_Profile.innerText=arrProfile[<%=session("language")%>]}catch(e){}
try{inner_Basic.innerText=arrBasic[<%=session("language")%>]}catch(e){}
try{inner_MyTask.innerText=arrMyTask[<%=session("language")%>]}catch(e){}
try{inner_MyForm.innerText=arrMyForm[<%=session("language")%>]}catch(e){}
try{inner_FormBoard.innerText=arrFormBoard[<%=session("language")%>]}catch(e){}
try{inner_LogOff.innerText=arrLogOff[<%=session("language")%>]}catch(e){}

try{inner_PrintPassword.innerText=arrPrintPassword[<%=session("language")%>]}catch(e){}

try{inner_Get_Unit.innerText=arrGetUnit[<%=session("language")%>]}catch(e){}
try{inner_Engineer_OQC.innerText=arrOQC[<%=session("language")%>]}catch(e){}

try{inner_ChangeModel.innerText=arrChangeModel[<%=session("language")%>]}catch(e){}

try{inner_OQCPPM.innerText=arrOQCPPM[<%=session("language")%>]}catch(e){}
try{inner_DailyStationYieldNew.innerText=arrDailyStationYieldNew[<%=session("language")%>]}catch(e){}
try{inner_WeeklyYield_OP.innerText=arrWeeklyYield_OP[<%=session("language")%>]}catch(e){}
try{inner_RetestYield.innerText=arrRetestYield[<%=session("language")%>]}catch(e){}
try{inner_SlappingRate.innerText=arrSlappingRate[<%=session("language")%>]}catch(e){}

try{inner_RetestIn.innerText=arrRetestIn[<%=session("language")%>]}catch(e){}
try{inner_JobStatus.innerText=arrJobStatus[<%=session("language")%>]}catch(e){}
try{inner_Solder.innerText=arrAutoSolder[<%=session("language")%>]}catch(e){}
try{inner_Priority.innerText=arrJobPriority[<%=session("language")%>]}catch(e){}

try{inner_JobType.innerText=arrJobType[<%=session("language")%>]}catch(e){}

try{inner_LABORRATE_OH.innerText=arrLABORRATE_OH[<%=session("language")%>]}catch(e){}

try{inner_2DCodeReport.innerText=arr2DCodeReport[<%=session("language")%>]}catch(e){}
try{inner_PackBoxReport.innerText=arrPackingBoxReport[<%=session("language")%>]}catch(e){}
try{inner_InventoryReport.innerText=arrInventoryReport[<%=session("language")%>]}catch(e){}

try{inner_MATERIAL.innerText=arrMATERIAL[<%=session("language")%>]}catch(e){}

try{inner_PrdReport.innerText=arrPrdReport[<%=session("language")%>]}catch(e){}
try{inner_CheckBoxList.innerText=arrCheckBoxList[<%=session("language")%>]}catch(e){}
try{inner_EMRList.innerText=arrEMRList[<%=session("language")%>]}catch(e){}
try{inner_PTCList.innerText=arrPTCList[<%=session("language")%>]}catch(e){}
try{inner_EMRManage.innerText=arrEMRManage[<%=session("language")%>]}catch(e){}
try{inner_Material_Record.innerText=arrMaterial_Record[<%=session("language")%>]}catch(e){}
try{inner_Supplier_Material.innerText=arrSupplier_Material[<%=session("language")%>]}catch(e){}
try{inner_Customer_Material.innerText=arrCustomer_Material[<%=session("language")%>]}catch(e){}






try{inner_TrayLink.innerText=arrTrayLink[<%=session("language")%>]}catch(e){}
try{inner_DefectCodeDistribution.innerText=arrDefectCodeDistribution[<%=session("language")%>]}catch(e){}
try{inner_ComputerSet.innerText=arrComputerSet[<%=session("language")%>]}catch(e){}
try{inner_WII.innerText=arrWII[<%=session("language")%>]}catch(e){}
try{inner_DeliveryReport.innerText=arrDeliveryReport[<%=session("language")%>]}catch(e){}
try{inner_DailyTarget.innerText=arrDailyTarget[<%=session("language")%>]}catch(e){}
try{inner_RMAReport.innerText=arrRMA[<%=session("language")%>]}catch(e){}
try{inner_WarehouseReport.innerText=arrWarehouseReport[<%=session("language")%>]}catch(e){}
}
language()
</script>