<script language="javascript">
var strBrowse="Add a New DB Job|�½��ƻ�����" 
var arrBrowse=strBrowse.split("|")
var strAdministrator="Person|������" 
var arrAdministrator=strAdministrator.split("|")
var strJobItem="Job Item|�ƻ��ű�" 
var arrJobItem=strJobItem.split("|")
var strScheduleDay="Schedule Day|�ƻ�����" 
var arrScheduleDay=strScheduleDay.split("|")
var strScheduleTime="Schedule Time|�ƻ�ʱ��" 
var arrScheduleTime=strScheduleTime.split("|")
var strHour="hour|��" 
var arrHour=strHour.split("|")
var strMinute="minutes|��" 
var arrMinute=strMinute.split("|")
var strSelect="-- Select --|-- ѡ�� --" 
var arrSelect=strSelect.split("|")
var strSave="Save|����" 
var arrSave=strSave.split("|")
var strReset="Reset|����"
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
