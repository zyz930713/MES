<script language="javascript">
var strSearch="Search Operator|��ѯ������"
var arrSearch=strSearch.split("|")
var strSearchCode="Code|����" 
var arrSearchCode=strSearchCode.split("|")
var arrUser=["User","�û�"]
var strSearchForcedCode="Forced Select |���ѡ�� " 
var arrSearchForcedCode=strSearchForcedCode.split("|")
var strSearchFactorySelect="-- Select Factory --|-- ѡ�񹤳� --" 
var arrSearchFactorySelect=strSearchFactorySelect.split("|")
var strSearchEnglishName="English Name|Ӣ������" 
var arrSearchEnglishName=strSearchEnglishName.split("|")
var strSearchChineseName="Chinese Name|��������" 
var arrSearchChineseName=strSearchChineseName.split("|")
var strSearchFind="Search|��ѯ" 
var arrSearchFind=strSearchFind.split("|")
var strBrowse="Browse Operators' Authority (by Operator) |�����������Ȩ�ޣ�����������" 
var arrBrowse=strBrowse.split("|")
var strbyStation="Browse by Station|��վ��" 
var arrbyStation=strbyStation.split("|")
var strbyPart="Browse by Part|���ͺ�" 
var arrbyPart=strbyPart.split("|")
var strAdd="Add a New Operator|����������" 
var arrAdd=strAdd.split("|")
var strNO="NO|����" 
var arrNO=strNO.split("|")
var strAction="Action|����" 
var arrAction=strAction.split("|")
var strLocked="Locked|�Ƿ�����" 
var arrLocked=strLocked.split("|")
var strPractised="Practised|�Ƿ�ʵϰ" 
var arrPractised=strPractised.split("|")
var strPeriod="Period|ʵϰ��" 
var arrPeriod=strPeriod.split("|")
var strCode="Code|����" 
var arrCode=strCode.split("|")
var strEnglishName="English Name|Ӣ������" 
var arrEnglishName=strEnglishName.split("|")
var strChineseName="Chinese Name|��������" 
var arrChineseName=strChineseName.split("|")
var strFactory="Factory|����" 
var arrFactory=strFactory.split("|")
var strLine="Line|�߱�" 
var arrLine=strLine.split("|")
var strStations="Authorized Stations|��Ȩվ��" 
var arrStations=strStations.split("|")
var strParts="Authorized Parts|��Ȩ�ͺ�" 
var arrParts=strParts.split("|")
var arrSearchFrom=["From","��"]
var arrSearchTo=["To","��"]
var arrEditData=["Edit Data","�༭����"]
var arrAddData=["Add Data","��������"]
var arrBtnOK=[" OK ","ȷ��"]
var arrBtnReset=["Reset","����"]
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
