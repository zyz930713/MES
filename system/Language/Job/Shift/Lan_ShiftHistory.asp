<script language="javascript">
var strBrowse="Browse Shift History|��������ʷ��¼" 
var arrBrowse=strBrowse.split("|")
var strUser="User|�û�" 
var arrUser=strUser.split("|")
var strNO="NO|����" 
var arrNO=strNO.split("|")
var strLineName="Line Name|�߱�����" 
var arrLineName=strLineName.split("|")
var strShiftType="Shift Type|�������" 
var arrShiftType=strShiftType.split("|")
var strShiftPerson="Shift Person|��ι�����" 
var arrShiftPerson=strShiftPerson.split("|")
var strShiftTime="Shift Time|���ʱ��" 
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
