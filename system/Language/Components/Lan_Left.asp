<script language="javascript">
var strHome="Home|首页"
var arrHome=strHome.split("|")
//-- Job Monitor
var strMonitorJobs="Monitor Jobs|监控工单" 
var arrMonitorJobs=strMonitorJobs.split("|")
var strJobsList="Master Jobs List|主工单列表" 
var arrJobsList=strJobsList.split("|")
var strSubJobs="Sub Jobs List|子工单列表" 
var arrSubJobs=strSubJobs.split("|")

var strSubJobsHistory="Sub Jobs List History|历史子工单列表" 
var arrSubJobsHistory=strSubJobsHistory.split("|")

 
var strJobDayCycle="Job Cycle|子工单周转" 
var arrJobDayCycle=strJobDayCycle.split("|")
var strJobsProgress="Jobs Progress|工单进程" 
var arrJobsProgress=strJobsProgress.split("|")
var strStoreRecords="Store Records|入库记录" 
var arrStoreRecords=strStoreRecords.split("|")
var strPreStoreRecords="Pre Store Records|预入库记录" 
var arrPreStoreRecords=strPreStoreRecords.split("|")
var strStoreChangeRecords="Store Change Records|入库变更记录" 
var arrStoreChangeRecords=strStoreChangeRecords.split("|")
var strPreScrapRecords="Pre Scrap Records|预报废记录" 
var arrPreScrapRecords=strPreScrapRecords.split("|")
var strScrapRecords="Scrap Records|报废记录" 
var arrScrapRecords=strScrapRecords.split("|")
var strScrapChangeRecords="Scrap Change Records|报废变更记录" 
var arrScrapChangeRecords=strScrapChangeRecords.split("|")
var strJobMaterial="Material Search|材料查询" 
var arrJobMaterial=strJobMaterial.split("|")
var strJobMaterial2="Material Dispatch|发料查询" 
var arrJobMaterial2=strJobMaterial2.split("|")
var strJobAction="Action Track|步骤追踪" 
var arrJobAction=strJobAction.split("|")
var strErrorJobs="Error Jobs|错误工单" 
var arrErrorJobs=strErrorJobs.split("|")
var strEEEJobs="EEE Jobs|试验工单" 
var arrEEEJobs=strEEEJobs.split("|")
var strShift="Shift Manage|班次开关" 
var arrShift=strShift.split("|")
var strLabour="Labour Manage|劳动力管理" 
var arrLabour=strLabour.split("|")
var strCustomizationReport="Customization Report|定制报告" 
var arrCustomizationReport=strCustomizationReport.split("|")
var strJobCycle="Job Cycle|工单周转" 
var arrJobCycle=strJobCycle.split("|")
var strJobLead="Job Lead|工单完成" 
var arrJobLead=strJobLead.split("|")
var strFamilyCycle="Family Cycle|家族周转" 
var arrFamilyCycle=strFamilyCycle.split("|")
var strUnCloseJob="UnClose Job|未关工单"
var arrUnCloseJob=strUnCloseJob.split("|")
var strBatchHold="Hold Job|Hold工单" 
var arrBatchHold=strBatchHold.split("|")

var strBatchRelease="Release Job|Release工单" 
var arrBatchRelease=strBatchRelease.split("|") 
var strUnInstore="UnInstore Job|未入库工单" 
var arrUnInstore=strUnInstore.split("|") 

var strJobPriority="Job Priority|工单优先级" 
var arrJobPriority=strJobPriority.split("|") 

var strPiecePartAutoScrap="Piece Parts Auto Scrap|原材料自动报废" 
var arrPiecePartAutoScrap=strPiecePartAutoScrap.split("|") 


var strPiecePartManualScrap="Piece Parts Manual Scrap|原材料手动报废" 
var arrPiecePartManualScrap=strPiecePartManualScrap.split("|") 


var strMisMatch="Final Test Mismatch Reg|不匹配数量登记" 
var arrMisMatch=strMisMatch.split("|") 



var strSubLineJob="Sub Line Job|Sub线工单" 
var arrSubLineJob=strSubLineJob.split("|") 

var strSubLineBOM="Sub Line BOM|Sub线BOM" 
var arrSubLineBOM=strSubLineBOM.split("|") 


var strLABORRATE_OH="Labor Rate&OH Setting|Labor Rate和OH设定" 
var arrLABORRATE_OH=strLABORRATE_OH.split("|") 

 

//--OBI
var strOBI="ERP Integration|ERP集成" 
var arrOBI=strOBI.split("|")
var strOBIJobClosing="Job Auto (Store)|工单自动（入库）" 
var arrOBIJobClosing=strOBIJobClosing.split("|")
var strOBIJobClosingScrap="Job Auto (Scrap)|工单自动（报废）" 
var arrOBIJobClosingScrap=strOBIJobClosingScrap.split("|")
var strOBIJobClosingChange="Job Auto (Store Change)|工单自动（入库变更）" 
var arrOBIJobClosingChange=strOBIJobClosingChange.split("|")
var strOBIJobClosingScrapChange="Job Auto (Scrap Change)|工单自动（报废变更）" 
var arrOBIJobClosingScrapChange=strOBIJobClosingScrapChange.split("|")
var strOBIJobSysClosing="Job Auto (Error)|工单自动（出错）" 
var arrOBIJobSysClosing=strOBIJobSysClosing.split("|")
//--Yield Report
var strDailyYieldReport="Daily Yield Report|每日良率报告"
var arrDailyYieldReport=strDailyYieldReport.split("|")
var strSeriesYield="Series Yield|系列良率" 
var arrSeriesYield=strSeriesYield.split("|")
var strJobYield="Job Yield|工单良率" 
var arrJobYield=strJobYield.split("|")
var strDailyFinanceYield="Daily Finance Yield|每日财务良率" 
var arrDailyFinanceYield=strDailyFinanceYield.split("|")
var strDailyStationYield="Daily Station Yield|每日站点良率" 
var arrDailyStationYield=strDailyStationYield.split("|")
var strDailyOperatorYield="Daily Operator Yield|每日操作工良率" 
var arrDailyOperatorYield=strDailyOperatorYield.split("|")
var strDailyMachineYield="Daily Machine Yield|每日机器良率" 
var arrDailyMachineYield=strDailyMachineYield.split("|")
var strDailyLineYield="Daily Line Yield|每日产线良率" 
var arrDailyLineYield=strDailyLineYield.split("|")
var strDailyStockYield="Daily Stock Yield|每日入库良率" 
var arrDailyStockYield=strDailyStockYield.split("|")
var strPartYield="Part Yield|型号良率" 
var arrPartYield=strPartYield.split("|")
var strMachineYield="Machine Yield|机器良率" 
var arrMachineYield=strMachineYield.split("|")
var strStationYield="Station Yield|站点良率" 
var arrStationYield=strStationYield.split("|")
var strLineYield="Line Yield|线别良率" 
var arrLineYield=strLineYield.split("|")
//--weekly yiled
var strWeeklyYieldReport="Weekly Yield Report|每周良率报告" 
var arrWeeklyYieldReport=strWeeklyYieldReport.split("|")
var strWeeklyFinanceYield="Weekly Finance Yield|每周财务良率" 
var arrWeeklyFinanceYield=strWeeklyFinanceYield.split("|")
var strFinalFamilyYield="Final Family Yield|最终家族良率" 
var arrFinalFamilyYield=strFinalFamilyYield.split("|")
var strFinalSeriesYield="Final Series Yield|最终系列良率" 
var arrFinalSeriesYield=strFinalSeriesYield.split("|")
var strFinalPartYield="Final Part Yield|最终型号良率" 
var arrFinalPartYield=strFinalPartYield.split("|")
//--Process
var strDailyLineWIP="Daily Line WIP|每日线别WIP" 
var arrDailyLineWIP=strDailyLineWIP.split("|")
var strDailyFamilyWIP="Daily Family WIP|每日家族WIP" 
var arrDailyFamilyWIP=strDailyFamilyWIP.split("|")
var strJobInformation="Job Information|工单信息" 
var arrJobInformation=strJobInformation.split("|")
var strStoreDefectcode="Store Defectcode|入库缺陷" 
var arrStoreDefectcode=strStoreDefectcode.split("|")
var strProcessReport="Process Report|过程报告" 
var arrProcessReport=strProcessReport.split("|")
var strFamilyLineLost="Family Line Lost|家族产线损失" 
var arrFamilyLineLost=strFamilyLineLost.split("|")
var strFamilyScrap="Family Scrap|家族报废" 
var arrFamilyScrap=strFamilyScrap.split("|")
var strFailureRatio="Failure Ratio|缺陷比例" 
var arrFailureRatio=strFailureRatio.split("|")

var strHoldReport="Job Hold Report|工单Hold报告" 
var arrHoldReport=strHoldReport.split("|")
var strJobCycleTime="Job Cycle Time|工单周转报告" 
var arrJobCycleTime=strJobCycleTime.split("|")
var strNONEFIFO="NONE FIFO|NONE FIFO" 
var arrNONEFIFO=strNONEFIFO.split("|")
var strJobOverdue="Job Overdue Report|工单超时报告" 
var arrJobOverdue=strJobOverdue.split("|")


var strOperatorOutput="Operator Output|操作工产出" 
var arrOperatorOutput=strOperatorOutput.split("|")
var strPartWIP="Part WIP|型号WIP" 
var arrPartWIP=strPartWIP.split("|")
var strLineWIP="Line WIP|线别WIP" 
var arrLineWIP=strLineWIP.split("|")
var strOutput="Output|产能" 
var arrOutput=strOutput.split("|")
var strJobStore="Job Store|工单入库报告" 
var arrJobStore=strJobStore.split("|")
var strStoreRecords="Store Records|入库记录" 
var arrStoreRecords=strStoreRecords.split("|")
var strStockRecords="Store List|入库清单" 
var arrStockRecords=strStockRecords.split("|")
var strStoreChange="Store Change|入库变更记录" 
var arrStoreChange=strStoreChange.split("|")
var strStoreNoYield="Store No Yield|不计良率入库" 
var arrStoreNoYield=strStoreNoYield.split("|")
var strScrapReport="Scrap Report|报废报告" 
var arrScrapReport=strScrapReport.split("|")
var strJobScrap="Job Scrap|工单报废" 
var arrJobScrap=strJobScrap.split("|")
var strScrapChange="Scrap Change|报废变更记录" 
var arrScrapChange=strScrapChange.split("|")
//--Chart Report
var strChartReport="Yield Report|良率报告" 
var arrChartReport=strChartReport.split("|")
var strChartProduction="Prodution Chart|生产图表" 
var arrChartProduction=strChartProduction.split("|")
var strChartWIP="WIP Chart|在制品图表" 
var arrChartWIP=strChartWIP.split("|")












var strChartFinalFamilyYieldCompare_Daily="First past yield compare Chart_Daily|First Past Yield日对比图表" 
var arrChartFinalFamilyYieldCompare_Daily=strChartFinalFamilyYieldCompare_Daily.split("|")

var strChartLineLost="Line Lost Chart|在线损失图表" 
var arrChartLineLost=strChartLineLost.split("|")
var strChartQuantityScrap="Scrap Chart (Quantity)|报废图表 （数量）" 
var arrChartQuantityScrap=strChartQuantityScrap.split("|")


var strNewYieldReport="FPY Report|一次良率报告" 
var arrNewYieldReport=strNewYieldReport.split("|")


var strNewFinalYieldReport3="Final Yield Detail Report|最终良率详细报告" 
var arrNewFinalYieldReport3=strNewFinalYieldReport3.split("|")

var strNewFinalYieldReport="Final Yield  Report|最终良率报告" 
var arrNewFinalYieldReport=strNewFinalYieldReport.split("|")


var strNewDefectReport="FPY Report|一次良率报告" 
var arrNewDefectReport=strNewDefectReport.split("|")


var strNewDefectReportWeekly="Weekly FPY Report|一次良率周报告" 
var arrNewDefectReportWeekly=strNewDefectReportWeekly.split("|")

var strNewDefectReportWeekly2="Weekly FPY Report(Day Stack_up)|一次良率周报告(日累加)" 
var arrNewDefectReportWeekly2=strNewDefectReportWeekly2.split("|")



var str_FPYWeeklyChart="Weekly FPY Chart Report(Day Stack_up)|一次良率周报告图表(日累加)" 
var arr_FPYWeeklyChart=str_FPYWeeklyChart.split("|")

var str_FPYDailyChart="Daily FPY Chart Report|一次良率日报告图表" 
var arr_FPYDailyChart=str_FPYDailyChart.split("|")

var str_Rework="Rework Summary Report|Rework 分类报告" 
var arr_Rework=str_Rework.split("|")

var str_ReworkSlappingYield="ReworkSlapping Yield Report|ReworkSlapping良率" 
var arr_ReworkSlappingYield=str_ReworkSlappingYield.split("|")



var str_NewJobYield="Job Yield|工单良率" 
var arr_NewJobYield=str_NewJobYield.split("|")
 
 var str_FY_Daily_TrackingReport="FY Daily Tracking|日良率跟踪报告" 
var arr_FY_Daily_TrackingReport=str_FY_Daily_TrackingReport.split("|")

 var str_GEYieldReport="GE Yield Report|GE 良率报告" 
var arr_GEYieldReport=str_GEYieldReport.split("|")

var str_FinancialReport="Financial Report|财务良率" 
var arr_FinancialReport=str_FinancialReport.split("|")

var str_StationGroupFinanceYield="Station Group Financial Yield|站点组财务良率" 
var arr_StationGroupFinanceYield=str_StationGroupFinanceYield.split("|")

var str_WIPGroup="WIP & Movement Report|WIP和运转报告" 
var arr_WIPGroup=str_WIPGroup.split("|")

 
var str_WIPReport="Before Test WIP Report|测试前WIP 报告" 
var arr_WIPReport=str_WIPReport.split("|")

var str_NewWIP="After Test WIP Report|测试后WIP报告" 
var arr_NewWIP=str_NewWIP.split("|")


var str_RetestQAHistory="Retest QA History|Retest QA历史报告" 
var arr_RetestQAHistory=str_RetestQAHistory.split("|")



//--Tracking
var strTrackingReport="Tracking Report|追踪报告" 
var arrTrackingReport=strTrackingReport.split("|")
var strJobSchedule="Job Schedule|工单计划" 
var arrJobSchedule=strJobSchedule.split("|")
var strJobTracking="Job Tracking|工单追踪" 
var arrJobTracking=strJobTracking.split("|")
//System Settings
var strSystemSettings="System Settings|系统设置" 
var arrSystemSettings=strSystemSettings.split("|")
var strFactories="Factories|工厂" 
var arrFactories=strFactories.split("|")
var strSections="Sections|生产区段" 
var arrSections=strSections.split("|")
var strLines="Lines|线别" 
var arrLines=strLines.split("|")
var strFinanceSeriesGroup="Finance Series Group|财务系列组" 
var arrFinanceSeriesGroup=strFinanceSeriesGroup.split("|")
var strFinanceSeries="Finance Series|财务系列" 
var arrFinanceSeries=strFinanceSeries.split("|")





var strSubSeriesNew="SubSeries|子系列" 
var arrSubSeriesNew=strSubSeriesNew.split("|")
var strSeriesNew="Series|系列" 
var arrSeriesNew=strSeriesNew.split("|")
var strFamilyNew="Family|家族" 
var arrFamilyNew=strFamilyNew.split("|")




var strModels="Models|型号" 
var arrModels=strModels.split("|")

var strGlueWIRE="Glue&WIRE Station Config|胶水&铜线工位配置" 
var arrGlueWIRE=strGlueWIRE.split("|")

var strPackingTest="Packing Config|包装配置" 
var arrPackingTest=strPackingTest.split("|")

var strPackingPlan="Packing Plan|包装计划" 
var arrPackingPlan=strPackingPlan.split("|")

var strPackingPlanKIPC="KIPC Packing Plan|KIPC 包装计划" 
var arrPackingPlanKIPC=strPackingPlanKIPC.split("|")
 
var strPartsRoutine="Parts Routine|制程" 
var arrPartsRoutine=strPartsRoutine.split("|")
var strStations="Stations|站点" 
var arrStations=strStations.split("|")

var strStationGroups="Station Group|站点组" 
var arrStationGroups=strStationGroups.split("|")

var strStationGroupSettings="Station Group Setting|站点组设定" 
var arrStationGroupSettings=strStationGroupSettings.split("|")



var strActions="Actions|步骤" 
var arrActions=strActions.split("|")
var strDefectCodes="Defect Codes|缺陷代码" 
var arrDefectCodes=strDefectCodes.split("|")

var strStations1="Stations|站点" 
var arrStations1=strStations1.split("|")
var strActions1="Actions|步骤" 
var arrActions1=strActions1.split("|")
var strDefectCodes1="Defect Codes|缺陷代码" 
var arrDefectCodes1=strDefectCodes1.split("|")
var strPartsRoutine1="Parts Routing|制程" 
var arrPartsRoutine1=strPartsRoutine1.split("|")

var strGeneralSetting="General Setting|一般设置" 
var arrGeneralSetting=strGeneralSetting.split("|")




var strDefectCodeGroups="Defect Code Groups|缺陷代码组" 
var arrDefectCodeGroups=strDefectCodeGroups.split("|")
var strMachines="Machines|机器" 
var arrMachines=strMachines.split("|")
 
var strGroups="Groups|组" 
var arrGroups=strGroups.split("|")
var arrHoldGroups=Array("Hold Groups","Hold群组")
var strAuthority="Authority|权限" 
var arrAuthority=strAuthority.split("|")
 
var strDual="Finance Dual|胞设置" 
var arrDual=strDual.split("|")
var strLCD="LCD|LCD" 
var arrLCD=strLCD.split("|")
//system administration
var strSystemAdministation="System Administation|系统管理" 
var arrSystemAdministation=strSystemAdministation.split("|")
var strOperators="Operators|操作员" 
var arrOperators=strOperators.split("|")
var strUpdateOperators="Update Operators|更新操作员信息" 
var arrUpdateOperators=strUpdateOperators.split("|")
var strEngineers="Engineers|工程师" 
var arrEngineers=strEngineers.split("|")
var strUpdateEngineers="Update Engineers|更新工程师信息" 
var arrUpdateEngineers=strUpdateEngineers.split("|")
var strUserGroups="User Groups|用户组" 
var arrUserGroups=strUserGroups.split("|")
var strRoles="Roles|角色" 
var arrRoles=strRoles.split("|")
var strApprovalRoles="Approval Roles|审批角色" 
var arrApprovalRoles=strApprovalRoles.split("|")
var strEvents="Events|事件" 
var arrEvents=strEvents.split("|")
var strTasks="System Task|系统任务" 
var arrTasks=strTasks.split("|")
var strForms="System Form|系统表单" 
var arrForms=strForms.split("|")
var strDatabaseJob="Database Job|数据库任务" 
var arrDatabaseJob=strDatabaseJob.split("|")
var strJobRecycle="Master Job Recycle|主回收桶" 
var arrJobRecycle=strJobRecycle.split("|")
var strSubJobRecycle="Sub Job Recycle|子工单回收桶" 
var arrSubJobRecycle=strSubJobRecycle.split("|")
var strAudit="Audit|审核" 
var arrAudit=strAudit.split("|")
var strHelp="Help|帮助" 
var arrHelp=strHelp.split("|")
var strProfile="Profile|个人配置" 
var arrProfile=strProfile.split("|")
var strBasic="Basic|基本信息" 
var arrBasic=strBasic.split("|")
var strMyTask="My Task|我的自定义任务" 
var arrMyTask=strMyTask.split("|")
var strMyForm="My Form|我的表单" 
var arrMyForm=strMyForm.split("|")
var strFormBoard="Form Board|表单看板" 
var arrFormBoard=strFormBoard.split("|")
var strLogOff="Log Off|注销" 
var arrLogOff=strLogOff.split("|")

var strPrintPassword="Print Password|标签打印密码" 
var arrPrintPassword=strPrintPassword.split("|")

var strGetUnit="Engineer Sample|工程师领借" 
var arrGetUnit=strGetUnit.split("|")
var strOQC="Release to OQC|转到OQC" 
var arrOQC=strOQC.split("|")


var strChangeModel="Change Model History|工单转型号记录" 
var arrChangeModel=strChangeModel.split("|")

var strOQCPPM="OQC PPM|OQC PPM报告" 
var arrOQCPPM=strOQCPPM.split("|")


var strDailyStationYieldNew="Station Daily Yield|Station日良率" 
var arrDailyStationYieldNew=strDailyStationYieldNew.split("|")

var strWeeklyYield_OP="Operator Weekly OutPut|操作员周产出报告" 
var arrWeeklyYield_OP=strWeeklyYield_OP.split("|")

var strRetestYield="Retest Yield|Retest 良率报告" 
var arrRetestYield=strRetestYield.split("|")


var strSlappingRate="Slapping Rate|拍打率" 
var arrSlappingRate=strSlappingRate.split("|")

var strRetestIn="Retest In|Retest投入" 
var arrRetestIn=strRetestIn.split("|")


var strJobStatus="Job Status|工单状态" 
var arrJobStatus=strJobStatus.split("|")


var strAutoSolder="Auto Solder Program|自动Solder程序" 
var arrAutoSolder=strAutoSolder.split("|")

var strJobPriority="Job Priority|Job 优先级" 
var arrJobPriority=strJobPriority.split("|")

var strJobType="Job Type|Job 类型" 
var arrJobType=strJobType.split("|")
var arr2DCodeReport=["2D Code Report","二维码报表"]
var arrWII=["WII","WII"]
var arrPrdReport=["Interval Production Report","生产报告"]
var arrCheckBoxList=["Check Box List","检查箱号"]
var arrEMRList=["Pot Membrane List","Pot 膜物料查询"]
var arrPTCList=["PTC List","产品追溯和控制"]
var arrEMRManage=["EMR Dispatch Manage","物料分配管理"]
var arrMaterial_Record=["Material Record","物料记录查询"]
var arrSupplier_Material=["Supplier Material Manage","供应商物料配置"]
var arrCustomer_Material=["Customer Lable Manage","客户标签配置"]
var arrPackingBoxReport=["Packing Report","包装报表"]
var arrStoreReport=["Store Report","入库报表"]
var arrInventoryReport=["Inventory Report","库存报表"]
var strTrayLink="TrayLink|Tray关系" 
var arrTrayLink=strTrayLink.split("|")

var strDefectCodeDistribution="Defect Code Distribution|缺陷代码分布" 
var arrDefectCodeDistribution=strDefectCodeDistribution.split("|")
var arrComputerSet=["Client Register","客户端注册"]
var arrDeliveryReport=["Delivery Report","交货报表"]
var arrDailyTarget=["Daily Target","生产目标设置"]
var arrMATERIAL=["MATERIAL Yield","原材料良率关系表"]
var arrRMA=["RMA Report","客退报告"]
var arrWarehouseReport=["Warehouse Report","仓库报表"]
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