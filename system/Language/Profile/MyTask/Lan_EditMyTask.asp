<script language="javascript">
var strBrowse="Customize Task|自定义任务"
var arrBrowse=strBrowse.split("|")
var strTaskType="Task Type|任务类型" 
var arrTaskType=strTaskType.split("|")
var strTaskName="Task Name|任务名称" 
var arrTaskName=strTaskName.split("|")
var strMailRecievers="Mail Recievers|收件人" 
var arrMailRecievers=strMailRecievers.split("|")
var strAvailableRecievers="Available Recievers|可选收件人" 
var arrAvailableRecievers=strAvailableRecievers.split("|")
var strSelectedRecievers="Selected Recievers|已选收件人" 
var arrSelectedRecievers=strSelectedRecievers.split("|")
var strParam1="Parameters 1|参数 1" 
var arrParam1=strParam1.split("|")
var strParam2="Parameters 2|参数 2" 
var arrParam2=strParam2.split("|")
var strParam3="Parameters 3|参数 3" 
var arrParam3=strParam3.split("|")
var strParam4="Parameters 4|参数 4" 
var arrParam4=strParam4.split("|")
var strSchedulePeriod="Has Scheduled|是否计划" 
var arrSchedulePeriod=strSchedulePeriod.split("|")
var strHappenItem="Hppen Item|发生周期" 
var arrHappenItem=strHappenItem.split("|")
var strMonday="Monday|星期一" 
var arrMonday=strMonday.split("|")
var strTuesday="Tuesday|星期二" 
var arrTuesday=strTuesday.split("|")
var strWendesday="Wendesday|星期三" 
var arrWendesday=strWendesday.split("|")
var strThursday="Thursday|星期四" 
var arrThursday=strThursday.split("|")
var strFriday="Friday|星期五" 
var arrFriday=strFriday.split("|")
var strSaturday="Saturday|星期六" 
var arrSaturday=strSaturday.split("|")
var strSunday="Sunday|星期日" 
var arrSunday=strSunday.split("|")
var strHappenTime="Hppen Time|发生时间" 
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