<script language="javascript">
var strSearch="Search Line|��ѯ�߱�"
var arrSearch=strSearch.split("|")
var strSearchLine="Line Name|�߱�" 
var arrSearchLine=strSearchLine.split("|")
var strBrowse="Browse Shift Status|������״̬" 
var arrBrowse=strBrowse.split("|")
var strUser="User|�û�" 
var arrUser=strUser.split("|")
var strbyPart="Browse by Part|���ͺ�" 
var arrbyPart=strbyPart.split("|")
var strNO="NO|����" 
var arrNO=strNO.split("|")
var strAction="Action|����" 
var arrAction=strAction.split("|")
var strSchedule="Schedule Job|���żƻ�" 
var arrSchedule=strSchedule.split("|")
var strRecords="Schedule Records|�ƻ���¼" 
var arrRecords=strRecords.split("|")
var strOpenLine="Click icon to open line|�����ť����" 
var arrOpenLine=strOpenLine.split("|")
var strStopLine="Click icon to stop line|�����ť����" 
var arrStopLine=strStopLine.split("|")
var strStatus="Status|��ǰ״̬" 
var arrStatus=strStatus.split("|")
var strLineName="Line Name|�߱�����" 
var arrLineName=strLineName.split("|")
var strFactory="Factory|����" 
var arrFactory=strFactory.split("|")
var strSection="Section|��������" 
var arrSection=strSection.split("|")
var strAdministrators="Administrators|����Ա" 
var arrAdministrators=strAdministrators.split("|")
function language()
{
try{inner_Search.innerText=arrSearch[<%=session("language")%>]}catch(e){} 
try{inner_SearchLine.innerText=arrSearchLine[<%=session("language")%>]}catch(e){}
try{inner_Browse.innerText=arrBrowse[<%=session("language")%>]}catch(e){}
try{inner_User.innerText=arrUser[<%=session("language")%>]}catch(e){}
try{inner_NO.innerText=arrNO[<%=session("language")%>]}catch(e){}
try{inner_Action.innerText=arrAction[<%=session("language")%>]}catch(e){}
try{inner_Schedule.innerText=arrSchedule[<%=session("language")%>]}catch(e){}
try{inner_Records.innerText=arrRecords[<%=session("language")%>]}catch(e){}
try{inner_openline.setAttribute("title",arrOpenLine[<%=session("language")%>])}catch(e){}
try{inner_stopline.setAttribute("title",arrStopLine[<%=session("language")%>])}catch(e){}
try{inner_Status.innerText=arrStatus[<%=session("language")%>]}catch(e){}
try{inner_LineName.innerText=arrLineName[<%=session("language")%>]}catch(e){}
try{document.all.factory.options[0].text=arrFactory[<%=session("language")%>]}catch(e){}
try{inner_Section.innerText=arrSection[<%=session("language")%>]}catch(e){}
try{inner_Administrators.innerText=arrAdministrators[<%=session("language")%>]}catch(e){}
}
</script>
