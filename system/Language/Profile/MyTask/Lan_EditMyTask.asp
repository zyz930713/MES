<script language="javascript">
var strBrowse="Customize Task|�Զ�������"
var arrBrowse=strBrowse.split("|")
var strTaskType="Task Type|��������" 
var arrTaskType=strTaskType.split("|")
var strTaskName="Task Name|��������" 
var arrTaskName=strTaskName.split("|")
var strMailRecievers="Mail Recievers|�ռ���" 
var arrMailRecievers=strMailRecievers.split("|")
var strAvailableRecievers="Available Recievers|��ѡ�ռ���" 
var arrAvailableRecievers=strAvailableRecievers.split("|")
var strSelectedRecievers="Selected Recievers|��ѡ�ռ���" 
var arrSelectedRecievers=strSelectedRecievers.split("|")
var strParam1="Parameters 1|���� 1" 
var arrParam1=strParam1.split("|")
var strParam2="Parameters 2|���� 2" 
var arrParam2=strParam2.split("|")
var strParam3="Parameters 3|���� 3" 
var arrParam3=strParam3.split("|")
var strParam4="Parameters 4|���� 4" 
var arrParam4=strParam4.split("|")
var strSchedulePeriod="Has Scheduled|�Ƿ�ƻ�" 
var arrSchedulePeriod=strSchedulePeriod.split("|")
var strHappenItem="Hppen Item|��������" 
var arrHappenItem=strHappenItem.split("|")
var strMonday="Monday|����һ" 
var arrMonday=strMonday.split("|")
var strTuesday="Tuesday|���ڶ�" 
var arrTuesday=strTuesday.split("|")
var strWendesday="Wendesday|������" 
var arrWendesday=strWendesday.split("|")
var strThursday="Thursday|������" 
var arrThursday=strThursday.split("|")
var strFriday="Friday|������" 
var arrFriday=strFriday.split("|")
var strSaturday="Saturday|������" 
var arrSaturday=strSaturday.split("|")
var strSunday="Sunday|������" 
var arrSunday=strSunday.split("|")
var strHappenTime="Hppen Time|����ʱ��" 
var arrHappenTime=strHappenTime.split("|")
function language()
{
try{inner_Browse.innerText=arrBrowse[<%=session("language")%>]}catch(e){} 
try{inner_TaskType.innerText=arrTaskType[<%=session("language")%>]}catch(e){} 
try{inner_TaskName.innerText=arrTaskName[<%=session("language")%>]}catch(e){}
try{inner_MailRecievers.innerText=arrMailRecievers[<%=session("language")%>]}catch(e){}
try{inner_AvailableRecievers.innerText=arrAvailableRecievers[<%=session("language")%>]}catch(e){}
try{inner_SelectedRecievers.innerText=arrSelectedRecievers[<%=session("language")%>]}catch(e){}
try{inner_Param1.innerText=arrParam1[<%=session("language")%>]}catch(e){}
try{inner_Param2.innerText=arrParam2[<%=session("language")%>]}catch(e){}
try{inner_Param3.innerText=arrParam3[<%=session("language")%>]}catch(e){}
try{inner_Param4.innerText=arrParam4[<%=session("language")%>]}catch(e){}
try{inner_SchedulePeriod.innerText=arrSchedulePeriod[<%=session("language")%>]}catch(e){}
try{inner_HappenItem.innerText=arrHappenItem[<%=session("language")%>]}catch(e){}
try{inner_Monday.innerText=arrMonday[<%=session("language")%>]}catch(e){}
try{inner_Tuesday.innerText=arrTuesday[<%=session("language")%>]}catch(e){}
try{inner_Wendesday.innerText=arrWendesday[<%=session("language")%>]}catch(e){}
try{inner_Thursday.innerText=arrThursday[<%=session("language")%>]}catch(e){}
try{inner_Friday.innerText=arrFriday[<%=session("language")%>]}catch(e){}
try{inner_Saturday.innerText=arrSaturday[<%=session("language")%>]}catch(e){}
try{inner_Sunday.innerText=arrSunday[<%=session("language")%>]}catch(e){}
try{inner_HappenTime.innerText=arrHappenTime[<%=session("language")%>]}catch(e){}
}
</script>