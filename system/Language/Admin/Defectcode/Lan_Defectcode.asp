<script language="javascript">
var strSearch="Search Defectcode|查询缺陷代码"
var arrSearch=strSearch.split("|")
var strSearchDefectID="Defect Code|缺陷代码" 
var arrSearchDefectID=strSearchDefectID.split("|")
var strSearchDefectCodeName="Defect Code Name|缺陷代码名称" 
var arrSearchDefectCodeName=strSearchDefectCodeName.split("|")
var strSearchChineseName="Chinese Name|中文名称" 
var arrSearchChineseName=strSearchChineseName.split("|")
var strSearchBelongedPart="Belonged Part|所属制程" 
var arrSearchBelongedPart=strSearchBelongedPart.split("|")
var strBrowse="Browse Defect Code List|查看缺陷代码列表" 
var arrBrowse=strBrowse.split("|")
var strUser="User|用户" 
var arrUser=strUser.split("|")
var strAdd="Add a New Defect Code|新增缺陷代码" 
var arrAdd=strAdd.split("|")
var strNO="NO|序列" 
var arrNO=strNO.split("|")
var strAction="Action|动作" 
var arrAction=strAction.split("|")
var strDefectCode="Defect Code|缺陷代码" 
var arrDefectCode=strDefectCode.split("|")
var strDefectName="Defect Name|名称" 
var arrDefectName=strDefectName.split("|")
var strChineseName="Chinese Name|中文名" 
var arrChineseName=strChineseName.split("|")
var strFactory="Factory|工厂" 
var arrFactory=strFactory.split("|")
var strTransaction="Transaction|线别" 
var arrTransaction=strTransaction.split("|")
var strAppliedMaterials="Applied Materials|应用材料" 
var arrAppliedMaterials=strAppliedMaterials.split("|")
var strBelongedPart="Belonged Part|所属制程" 
var arrBelongedPart=strBelongedPart.split("|")
var strBelongedStation="Belonged Station|所属站别" 
var arrBelongedStation=strBelongedStation.split("|")
function language()
{
try{inner_Search.innerText=arrSearch[<%=session("language")%>]}catch(e){} 
try{inner_SearchDefectID.innerText=arrSearchDefectID[<%=session("language")%>]}catch(e){}
try{inner_SearchDefectCodeName.innerText=arrSearchDefectCodeName[<%=session("language")%>]}catch(e){}
try{inner_SearchChineseName.innerText=arrSearchChineseName[<%=session("language")%>]}catch(e){}
try{inner_SearchBelongedPart.innerText=arrSearchBelongedPart[<%=session("language")%>]}catch(e){}
try{inner_Browse.innerText=arrBrowse[<%=session("language")%>]}catch(e){}
try{inner_User.innerText=arrUser[<%=session("language")%>]}catch(e){}
try{inner_Add.innerText=arrAdd[<%=session("language")%>]}catch(e){}
try{inner_NO.innerText=arrNO[<%=session("language")%>]}catch(e){}
try{inner_Action.innerText=arrAction[<%=session("language")%>]}catch(e){}
try{inner_DefectCode.innerText=arrDefectCode[<%=session("language")%>]}catch(e){}
try{inner_DefectName.innerText=arrDefectName[<%=session("language")%>]}catch(e){}
try{inner_ChineseName.innerText=arrChineseName[<%=session("language")%>]}catch(e){}
try{document.all.factory.options[0].text=arrFactory[<%=session("language")%>]}catch(e){}
try{inner_Transaction.innerText=arrTransaction[<%=session("language")%>]}catch(e){}
try{inner_AppliedMaterials.innerText=arrAppliedMaterials[<%=session("language")%>]}catch(e){}
try{document.all.part.options[0].text=arrBelongedPart[<%=session("language")%>]}catch(e){}
try{document.all.station.options[0].text=arrBelongedStation[<%=session("language")%>]}catch(e){}
}
</script>
