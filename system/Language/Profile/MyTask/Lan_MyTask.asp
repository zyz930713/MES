<script language="javascript">
var strSearch="Search Task|搜索任务"
var arrSearch=strSearch.split("|")
var strSearchUser="User|用户"
var arrSearchUser=strSearchUser.split("|")
var strSearchName="Name|任务名称"
var arrSearchName=strSearchName.split("|")
var strSearchButton="Search|搜索"
var arrSearchButton=strSearchButton.split("|")

var strBrowse="My Task List|我的自定义任务列表"
var arrBrowse=strBrowse.split("|")
var strBrowseUser="User|用户"
var arrBrowseUser=strBrowseUser.split("|")
var strAdd="Add a New Task|新增任务" 
var arrAdd=strAdd.split("|")
var strNO="No|编号" 
var arrNO=strNO.split("|")
var strAction="Action|动作" 
var arrAction=strAction.split("|")
var strUser="User|用户" 
var arrUser=strUser.split("|")
var strTaskName="Task Name|任务名称" 
var arrTaskName=strTaskName.split("|")
var strTaskType="Task Type|任务类型" 
var arrTaskType=strTaskType.split("|")
var strScheduleType="Schedule Type|计划类型" 
var arrScheduleType=strScheduleType.split("|")
var strStartDay="Start Day|起始日期" 
var arrStartDay=strStartDay.split("|")
var strRunTime="Run Time|运行时间" 
var arrRunTime=strRunTime.split("|")
var strSchedule="Scheduled Week Day|计划星期数" 
var arrSchedule=strSchedule.split("|")
var strNextRunTime="Next Run Time|下一次运行时间" 
var arrNextRunTime=strNextRunTime.split("|")
var strRecievers="Recievers|收件人" 
var arrRecievers=strRecievers.split("|")
function language()
{
try{inner_Search.innerText=arrSearch[<%=session("language")%>]}catch(e){} 
try{inner_SearchUser.innerText=arrSearchUser[<%=session("language")%>]}catch(e){} 
try{inner_SearchName.innerText=arrSearchName[<%=session("language")%>]}catch(e){} 
try{document.all.searchbutton.value=arrSearchButton[<%=session("language")%>]}catch(e){} 
try{inner_Browse.innerText=arrBrowse[<%=session("language")%>]}catch(e){} 
try{inner_BrowseUser.innerText=arrBrowseUser[<%=session("language")%>]}catch(e){} 
try{inner_Add.innerText=arrAdd[<%=session("language")%>]}catch(e){} 
try{inner_NO.innerText=arrNO[<%=session("language")%>]}catch(e){} 
try{inner_Action.innerText=arrAction[<%=session("language")%>]}catch(e){}
try{inner_User.innerText=arrUser[<%=session("language")%>]}catch(e){}
try{inner_TaskName.innerText=arrTaskName[<%=session("language")%>]}catch(e){}
try{inner_TaskType.innerText=arrTaskType[<%=session("language")%>]}catch(e){}
try{inner_ScheduleType.innerText=arrScheduleType[<%=session("language")%>]}catch(e){}
try{inner_StartDay.innerText=arrStartDay[<%=session("language")%>]}catch(e){}
try{inner_RunTime.innerText=arrRunTime[<%=session("language")%>]}catch(e){}
try{inner_Schedule.innerText=arrSchedule[<%=session("language")%>]}catch(e){}
try{inner_NextRunTime.innerText=arrNextRunTime[<%=session("language")%>]}catch(e){}
try{inner_Recievers.innerText=arrRecievers[<%=session("language")%>]}catch(e){}
}
</script>