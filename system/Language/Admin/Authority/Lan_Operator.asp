<script language="javascript">
var strSearch="Search Operator|查询操作工"
var arrSearch=strSearch.split("|")
var strSearchCode="Code|工号" 
var arrSearchCode=strSearchCode.split("|")
var arrUser=["User","用户"]
var strSearchForcedCode="Forced Select |借调选择 " 
var arrSearchForcedCode=strSearchForcedCode.split("|")
var strSearchFactorySelect="-- Select Factory --|-- 选择工厂 --" 
var arrSearchFactorySelect=strSearchFactorySelect.split("|")
var strSearchEnglishName="English Name|英文姓名" 
var arrSearchEnglishName=strSearchEnglishName.split("|")
var strSearchChineseName="Chinese Name|中文姓名" 
var arrSearchChineseName=strSearchChineseName.split("|")
var strSearchFind="Search|查询" 
var arrSearchFind=strSearchFind.split("|")
var strBrowse="Browse Operators' Authority (by Operator) |浏览操作工的权限（按操作工）" 
var arrBrowse=strBrowse.split("|")
var strbyStation="Browse by Station|按站点" 
var arrbyStation=strbyStation.split("|")
var strbyPart="Browse by Part|按型号" 
var arrbyPart=strbyPart.split("|")
var strAdd="Add a New Operator|新增操作工" 
var arrAdd=strAdd.split("|")
var strNO="NO|序列" 
var arrNO=strNO.split("|")
var strAction="Action|动作" 
var arrAction=strAction.split("|")
var strLocked="Locked|是否锁定" 
var arrLocked=strLocked.split("|")
var strPractised="Practised|是否实习" 
var arrPractised=strPractised.split("|")
var strPeriod="Period|实习期" 
var arrPeriod=strPeriod.split("|")
var strCode="Code|工号" 
var arrCode=strCode.split("|")
var strEnglishName="English Name|英文姓名" 
var arrEnglishName=strEnglishName.split("|")
var strChineseName="Chinese Name|中文姓名" 
var arrChineseName=strChineseName.split("|")
var strFactory="Factory|工厂" 
var arrFactory=strFactory.split("|")
var strLine="Line|线别" 
var arrLine=strLine.split("|")
var strStations="Authorized Stations|授权站点" 
var arrStations=strStations.split("|")
var strParts="Authorized Parts|授权型号" 
var arrParts=strParts.split("|")
var arrSearchFrom=["From","从"]
var arrSearchTo=["To","到"]
var arrEditData=["Edit Data","编辑数据"]
var arrAddData=["Add Data","新增数据"]
var arrBtnOK=[" OK ","确定"]
var arrBtnReset=["Reset","重置"]
function language()
{
try{inner_Search.innerText=arrSearch[<%=session("language")%>]}catch(e){} 
try{inner_EditData.innerText=arrEditData[<%=session("language")%>]}catch(e){}
try{inner_AddData.innerText=arrAddData[<%=session("language")%>]}catch(e){}
try{inner_SearchCode.innerText=arrSearchCode[<%=session("language")%>]}catch(e){}
try{inner_User.innerText=arrUser[<%=session("language")%>]}catch(e){}
try{inner_SearchForcedCode.innerText=arrSearchForcedCode[<%=session("language")%>]}catch(e){}
try{document.all.factory.options[0].text=arrSearchFactorySelect[<%=session("language")%>]}catch(e){}
try{inner_SearchEnglishName.innerText=arrSearchEnglishName[<%=session("language")%>]}catch(e){}
try{inner_SearchChineseName.innerText=arrSearchChineseName[<%=session("language")%>]}catch(e){}
try{document.all.Find.value=arrSearchFind[<%=session("language")%>]}catch(e){}
try{inner_Browse.innerText=arrBrowse[<%=session("language")%>]}catch(e){}
try{document.all.byStation.value=arrbyStation[<%=session("language")%>]}catch(e){}
try{document.all.byPart.value=arrbyPart[<%=session("language")%>]}catch(e){}
try{inner_Add.innerText=arrAdd[<%=session("language")%>]}catch(e){}
try{inner_NO.innerText=arrNO[<%=session("language")%>]}catch(e){}
try{inner_Action.innerText=arrAction[<%=session("language")%>]}catch(e){}
try{inner_Locked.innerText=arrLocked[<%=session("language")%>]}catch(e){}
try{inner_Practised.innerText=arrPractised[<%=session("language")%>]}catch(e){}
try{inner_Period.innerText=arrPeriod[<%=session("language")%>]}catch(e){}
try{inner_Code.innerText=arrCode[<%=session("language")%>]}catch(e){}
try{inner_EnglishName.innerText=arrEnglishName[<%=session("language")%>]}catch(e){}
try{inner_ChineseName.innerText=arrChineseName[<%=session("language")%>]}catch(e){}
try{document.all.factory.options[0].text=arrFactory[<%=session("language")%>]}catch(e){}
try{document.all.line.options[0].text=arrLine[<%=session("language")%>]}catch(e){}
try{inner_Stations.innerText=arrStations[<%=session("language")%>]}catch(e){}
try{inner_Parts.innerText=arrParts[<%=session("language")%>]}catch(e){}
try{inner_factory.innerText=arrFactory[<%=session("language")%>]}catch(e){}
try{inner_PeriodFrom.innerText=arrPeriod[<%=session("language")%>]+" "+arrSearchFrom[<%=session("language")%>]}catch(e){}
try{inner_PeriodTo.innerText=arrPeriod[<%=session("language")%>]+" "+arrSearchTo[<%=session("language")%>]}catch(e){}

try{document.all.btnOK.value=arrBtnOK[<%=session("language")%>]}catch(e){}
try{document.all.btnReset.value=arrBtnReset[<%=session("language")%>]}catch(e){}
}
language()
</script>
