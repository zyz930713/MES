<script language="javascript">
var strSearch="Search Line|查询线别"
var arrSearch=strSearch.split("|")
var strSearchLine="Line Name|线别" 
var arrSearchLine=strSearchLine.split("|")
var strBrowse="Browse Shift Status|浏览班次状态" 
var arrBrowse=strBrowse.split("|")
var strUser="User|用户" 
var arrUser=strUser.split("|")
var strbyPart="Browse by Part|按型号" 
var arrbyPart=strbyPart.split("|")
var strNO="NO|序列" 
var arrNO=strNO.split("|")
var strAction="Action|动作" 
var arrAction=strAction.split("|")
var strSchedule="Schedule Job|安排计划" 
var arrSchedule=strSchedule.split("|")
var strRecords="Schedule Records|计划记录" 
var arrRecords=strRecords.split("|")
var strOpenLine="Click icon to open line|点击按钮开线" 
var arrOpenLine=strOpenLine.split("|")
var strStopLine="Click icon to stop line|点击按钮关线" 
var arrStopLine=strStopLine.split("|")
var strStatus="Status|当前状态" 
var arrStatus=strStatus.split("|")
var strLineName="Line Name|线别名称" 
var arrLineName=strLineName.split("|")
var strFactory="Factory|工厂" 
var arrFactory=strFactory.split("|")
var strSection="Section|生产区段" 
var arrSection=strSection.split("|")
var strAdministrators="Administrators|管理员" 
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
