<script language="javascript">
var strBrowse="Edit Schedule|�޸ļƻ�����" 
var arrBrowse=strBrowse.split("|")
var strAdministrator="Person|������" 
var arrAdministrator=strAdministrator.split("|")
var strJobItem="Job Item|�ƻ��Ų�" 
var arrJobItem=strJobItem.split("|")
var strScheduleDay="Schedule Day|�ƻ�����" 
var arrScheduleDay=strScheduleDay.split("|")
var strScheduleTime="Schedule Time|�ƻ�ʱ��" 
var arrScheduleTime=strScheduleTime.split("|")
var strSelect="-- Select --|-- ѡ�� --" 
var arrSelect=strSelect.split("|")
var strHour="hour|��" 
var arrHour=strHour.split("|")
var strMinute="minutes|��" 
var arrMinute=strMinute.split("|")
var strUpdate="Update|����" 
var arrUpdate=strUpdate.split("|")
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
try{document.all.Update.value=arrUpdate[<%=session("language")%>]}catch(e){}
try{document.all.Reset.value=arrReset[<%=session("language")%>]}catch(e){}
}
</script>
