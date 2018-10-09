<script language="javascript">
var strBrowse="Add a New DB Job|新建计划任务" 
var arrBrowse=strBrowse.split("|")
var strAdministrator="Person|创建人" 
var arrAdministrator=strAdministrator.split("|")
var strJobItem="Job Item|计划脚本" 
var arrJobItem=strJobItem.split("|")
var strScheduleDay="Schedule Day|计划日期" 
var arrScheduleDay=strScheduleDay.split("|")
var strScheduleTime="Schedule Time|计划时间" 
var arrScheduleTime=strScheduleTime.split("|")
var strHour="hour|点" 
var arrHour=strHour.split("|")
var strMinute="minutes|分" 
var arrMinute=strMinute.split("|")
var strSelect="-- Select --|-- 选择 --" 
var arrSelect=strSelect.split("|")
var strSave="Save|保存" 
var arrSave=strSave.split("|")
var strReset="Reset|重置"
var arrReset=strReset.split("|")

function language()
{
try{inner_Browse.innerText=arrBrowse[<%=session("language")%>]}catch(e){}
try{inner_Administrator.innerText=arrAdministrator[<%=session("language")%>]}catch(e){}
try{inner_JobItem.innerText=arrJobItem[<%=session("language")%>]}catch(e){}
try{inner_ScheduleDay.innerText=arrScheduleDay[<%=session("language")%>]}catch(e){}
try{inner_ScheduleTime.innerText=arrScheduleTime[<%=session("language")%>]}catch(e){}
try{inner_ScheduleTime.innerText=arrScheduleTime[<%=session("language")%>]}catch(e){}
try{inner_Hour.innerText=arrHour[<%=session("language")%>]}catch(e){}
try{inner_Minute.innerText=arrMinute[<%=session("language")%>]}catch(e){}
try{document.all.hour.options[0].text=arrSelect[<%=session("language")%>]}catch(e){}
try{document.all.minute.options[0].text=arrSelect[<%=session("language")%>]}catch(e){}
try{document.all.minute.options[0].text=arrSelect[<%=session("language")%>]}catch(e){}
try{document.all.Save.value=arrSave[<%=session("language")%>]}catch(e){}
try{document.all.Reset.value=arrReset[<%=session("language")%>]}catch(e){}
}
</script>
