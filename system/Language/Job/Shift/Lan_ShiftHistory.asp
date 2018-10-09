<script language="javascript">
var strBrowse="Browse Shift History|浏览班次历史纪录" 
var arrBrowse=strBrowse.split("|")
var strUser="User|用户" 
var arrUser=strUser.split("|")
var strNO="NO|序列" 
var arrNO=strNO.split("|")
var strLineName="Line Name|线别名称" 
var arrLineName=strLineName.split("|")
var strShiftType="Shift Type|班次类型" 
var arrShiftType=strShiftType.split("|")
var strShiftPerson="Shift Person|班次管理人" 
var arrShiftPerson=strShiftPerson.split("|")
var strShiftTime="Shift Time|班次时间" 
var arrShiftTime=strShiftTime.split("|")
function language()
{
try{inner_Browse.innerText=arrBrowse[<%=session("language")%>]}catch(e){}
try{inner_User.innerText=arrUser[<%=session("language")%>]}catch(e){}
try{inner_NO.innerText=arrNO[<%=session("language")%>]}catch(e){}
try{inner_LineName.innerText=arrLineName[<%=session("language")%>]}catch(e){}
try{inner_ShiftType.innerText=arrShiftType[<%=session("language")%>]}catch(e){}
try{inner_ShiftPerson.innerText=arrShiftPerson[<%=session("language")%>]}catch(e){}
try{inner_ShiftTime.innerText=arrShiftTime[<%=session("language")%>]}catch(e){}
}
language()
</script>
