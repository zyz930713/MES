<script language="javascript">
var strSearch="Search Defectcode|��ѯȱ�ݴ���"
var arrSearch=strSearch.split("|")
var strSearchDefectID="Defect Code|ȱ�ݴ���" 
var arrSearchDefectID=strSearchDefectID.split("|")
var strSearchDefectCodeName="Defect Code Name|ȱ�ݴ�������" 
var arrSearchDefectCodeName=strSearchDefectCodeName.split("|")
var strSearchChineseName="Chinese Name|��������" 
var arrSearchChineseName=strSearchChineseName.split("|")
var strSearchBelongedPart="Belonged Part|�����Ƴ�" 
var arrSearchBelongedPart=strSearchBelongedPart.split("|")
var strBrowse="Browse Defect Code List|�鿴ȱ�ݴ����б�" 
var arrBrowse=strBrowse.split("|")
var strUser="User|�û�" 
var arrUser=strUser.split("|")
var strAdd="Add a New Defect Code|����ȱ�ݴ���" 
var arrAdd=strAdd.split("|")
var strNO="NO|����" 
var arrNO=strNO.split("|")
var strAction="Action|����" 
var arrAction=strAction.split("|")
var strDefectCode="Defect Code|ȱ�ݴ���" 
var arrDefectCode=strDefectCode.split("|")
var strDefectName="Defect Name|����" 
var arrDefectName=strDefectName.split("|")
var strChineseName="Chinese Name|������" 
var arrChineseName=strChineseName.split("|")
var strFactory="Factory|����" 
var arrFactory=strFactory.split("|")
var strTransaction="Transaction|�߱�" 
var arrTransaction=strTransaction.split("|")
var strAppliedMaterials="Applied Materials|Ӧ�ò���" 
var arrAppliedMaterials=strAppliedMaterials.split("|")
var strBelongedPart="Belonged Part|�����Ƴ�" 
var arrBelongedPart=strBelongedPart.split("|")
var strBelongedStation="Belonged Station|����վ��" 
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
