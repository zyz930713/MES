<script language="javascript">
var strBrowse="Add a New Schedule for|�½��ƻ�����" 
var arrBrowse=strBrowse.split("|")
var strLineName="Line Name|�߱�����" 
var arrLineName=strLineName.split("|")
var strAdministrator="Person|������" 
var arrAdministrator=strAdministrator.split("|")
var strShiftType="Shift Type|�������" 
var arrShiftType=strShiftType.split("|")
var strShiftOut="Shift Out|ͣ��" 
var arrShiftOut=strShiftOut.split("|")
var strShiftIn="Shift In|����" 
var arrShiftIn=strShiftIn.split("|")
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
try{inner_LineName.innerText=arrLineName[<%=session("language")%>]}catch(e){}
try{inner_Administrator.innerText=arrAdministrator[<%=session("language")%>]}catch(e){}
try{inner_ShiftType.innerText=arrShiftType[<%=session("language")%>]}catch(e){}
try{inner_ShiftOut.innerText=arrShiftOut[<%=session("language")%>]}catch(e){}
try{inner_ShiftIn.innerText=arrShiftIn[<%=session("language")%>]}catch(e){}
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
