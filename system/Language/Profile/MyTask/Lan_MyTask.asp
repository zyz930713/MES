<script language="javascript">
var strSearch="Search Task|��������"
var arrSearch=strSearch.split("|")
var strSearchUser="User|�û�"
var arrSearchUser=strSearchUser.split("|")
var strSearchName="Name|��������"
var arrSearchName=strSearchName.split("|")
var strSearchButton="Search|����"
var arrSearchButton=strSearchButton.split("|")

var strBrowse="My Task List|�ҵ��Զ��������б�"
var arrBrowse=strBrowse.split("|")
var strBrowseUser="User|�û�"
var arrBrowseUser=strBrowseUser.split("|")
var strAdd="Add a New Task|��������" 
var arrAdd=strAdd.split("|")
var strNO="No|���" 
var arrNO=strNO.split("|")
var strAction="Action|����" 
var arrAction=strAction.split("|")
var strUser="User|�û�" 
var arrUser=strUser.split("|")
var strTaskName="Task Name|��������" 
var arrTaskName=strTaskName.split("|")
var strTaskType="Task Type|��������" 
var arrTaskType=strTaskType.split("|")
var strScheduleType="Schedule Type|�ƻ�����" 
var arrScheduleType=strScheduleType.split("|")
var strStartDay="Start Day|��ʼ����" 
var arrStartDay=strStartDay.split("|")
var strRunTime="Run Time|����ʱ��" 
var arrRunTime=strRunTime.split("|")
var strSchedule="Scheduled Week Day|�ƻ�������" 
var arrSchedule=strSchedule.split("|")
var strNextRunTime="Next Run Time|��һ������ʱ��" 
var arrNextRunTime=strNextRunTime.split("|")
var strRecievers="Recievers|�ռ���" 
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