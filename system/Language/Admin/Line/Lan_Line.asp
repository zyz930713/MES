<script language="javascript">
var strSearch="Search Line|查询线别"
var arrSearch=strSearch.split("|")
var strSearchLine="Line Name|线别" 
var arrSearchLine=strSearchLine.split("|")
var strBrowse="Browse Line List|浏览线别" 
var arrBrowse=strBrowse.split("|")
var strUser="User|用户" 
var arrUser=strUser.split("|")
var strbyPart="Browse by Part|按型号" 
var arrbyPart=strbyPart.split("|")
var strAdd="Add a New Line|新增线别" 
var arrAdd=strAdd.split("|")
var strNO="NO|序列" 
var arrNO=strNO.split("|")
var strAction="Action|动作" 
var arrAction=strAction.split("|")
var strStatus="Status|状态" 
var arrStatus=strStatus.split("|")
var strLineName="Line Name|线别名称" 
var arrLineName=strLineName.split("|")
var strFactory="Factory|工厂" 
var arrFactory=strFactory.split("|")
var strSection="Section|生产区段" 
var arrSection=strSection.split("|")
var strGroupLeader="Group Leader|领班" 
var arrGroupLeader=strGroupLeader.split("|")
var strSupervisor="Supervisor|主管" 
var arrSupervisor=strSupervisor.split("|")
var strMachineLabels="Machine Labels|机器编码" 
var arrMachineLabels=strMachineLabels.split("|")
var strAppliedParts="Applied Parts|应用型号" 
var arrAppliedParts=strAppliedParts.split("|")
var arrEditData=["Edit Data","编辑数据"]
var arrAddData=["Add Data","新增数据"]
var arrBtnOK=[" OK ","确定"]
var arrBtnReset=["Reset","重置"]
var arrFactoryCode=["Factory Code","工厂代码"]
var arrCodeDate690=["690 Code Date","690生产日期"]
var arrCodeDate=["Code Date","打包生产日期"]
var arrCodeLineName=["2DCode Line Name","产品线别"]
var arrCodeName=["Packing 2DCode Name","打包产品名称"]
var arr690CodeName=["690 2DCode Name","690产品名称"]
var arrVersionNumber=["Version Number","版本号"]
var arrPRODUCT=["Product Name","项目名称"]





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
