<script language="javascript">
var strSearch="Search Line|��ѯ�߱�"
var arrSearch=strSearch.split("|")
var strSearchLine="Line Name|�߱�" 
var arrSearchLine=strSearchLine.split("|")
var strBrowse="Browse Line List|����߱�" 
var arrBrowse=strBrowse.split("|")
var strUser="User|�û�" 
var arrUser=strUser.split("|")
var strbyPart="Browse by Part|���ͺ�" 
var arrbyPart=strbyPart.split("|")
var strAdd="Add a New Line|�����߱�" 
var arrAdd=strAdd.split("|")
var strNO="NO|����" 
var arrNO=strNO.split("|")
var strAction="Action|����" 
var arrAction=strAction.split("|")
var strStatus="Status|״̬" 
var arrStatus=strStatus.split("|")
var strLineName="Line Name|�߱�����" 
var arrLineName=strLineName.split("|")
var strFactory="Factory|����" 
var arrFactory=strFactory.split("|")
var strSection="Section|��������" 
var arrSection=strSection.split("|")
var strGroupLeader="Group Leader|���" 
var arrGroupLeader=strGroupLeader.split("|")
var strSupervisor="Supervisor|����" 
var arrSupervisor=strSupervisor.split("|")
var strMachineLabels="Machine Labels|��������" 
var arrMachineLabels=strMachineLabels.split("|")
var strAppliedParts="Applied Parts|Ӧ���ͺ�" 
var arrAppliedParts=strAppliedParts.split("|")
var arrEditData=["Edit Data","�༭����"]
var arrAddData=["Add Data","��������"]
var arrBtnOK=[" OK ","ȷ��"]
var arrBtnReset=["Reset","����"]
var arrFactoryCode=["Factory Code","��������"]
var arrCodeDate690=["690 Code Date","690��������"]
var arrCodeDate=["Code Date","�����������"]
var arrCodeLineName=["2DCode Line Name","��Ʒ�߱�"]
var arrCodeName=["Packing 2DCode Name","�����Ʒ����"]
var arr690CodeName=["690 2DCode Name","690��Ʒ����"]
var arrVersionNumber=["Version Number","�汾��"]
var arrPRODUCT=["Product Name","��Ŀ����"]





function language()
{
try{inner_EditData.innerText=arrEditData[<%=session("language")%>]}catch(e){}
try{inner_AddData.innerText=arrAddData[<%=session("language")%>]}catch(e){}
try{inner_Search.innerText=arrSearch[<%=session("language")%>]}catch(e){} 
try{inner_SearchLine.innerText=arrSearchLine[<%=session("language")%>]}catch(e){}
try{inner_Browse.innerText=arrBrowse[<%=session("language")%>]}catch(e){}
try{inner_User.innerText=arrUser[<%=session("language")%>]}catch(e){}
try{inner_Add.innerText=arrAdd[<%=session("language")%>]}catch(e){}
try{inner_NO.innerText=arrNO[<%=session("language")%>]}catch(e){}
try{inner_Action.innerText=arrAction[<%=session("language")%>]}catch(e){}
try{document.all.status.options[0].text=arrStatus[<%=session("language")%>]}catch(e){}
try{inner_LineName.innerText=arrLineName[<%=session("language")%>]}catch(e){}
try{inner_FactoryCode.innerText=arrFactoryCode[<%=session("language")%>]}catch(e){}




try{inner_CodeDate690.innerText=arrCodeDate690[<%=session("language")%>]}catch(e){}
try{inner_CodeDate.innerText=arrCodeDate[<%=session("language")%>]}catch(e){}
try{inner_CodeLineName.innerText=arrCodeLineName[<%=session("language")%>]}catch(e){}
try{inner_CodeName.innerText=arrCodeName[<%=session("language")%>]}catch(e){}
try{inner_690CodeName.innerText=arr690CodeName[<%=session("language")%>]}catch(e){}
try{inner_VersionNumber.innerText=arrVersionNumber[<%=session("language")%>]}catch(e){}
try{inner_Factory.innerText=arrFactory[<%=session("language")%>]}catch(e){}
try{inner_Section.innerText=arrSection[<%=session("language")%>]}catch(e){}
try{inner_GroupLeader.innerText=arrGroupLeader[<%=session("language")%>]}catch(e){}
try{inner_Supervisor.innerText=arrSupervisor[<%=session("language")%>]}catch(e){}
try{inner_MachineLabels.innerText=arrMachineLabels[<%=session("language")%>]}catch(e){}
try{inner_AppliedParts.innerText=arrAppliedParts[<%=session("language")%>]}catch(e){}
try{inner_Product.innerText=arrPRODUCT[<%=session("language")%>]}catch(e){}






try{document.all.btnOK.value=arrBtnOK[<%=session("language")%>]}catch(e){}
try{document.all.btnReset.value=arrBtnReset[<%=session("language")%>]}catch(e){}
}
</script>
