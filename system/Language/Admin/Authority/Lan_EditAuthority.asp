<script language="javascript">
var strBrowse="Edit Operator's (|浏览操作工（" 
var arrBrowse=strBrowse.split("|")
var strBrowse2=") Authority|）的权限" 
var arrBrowse2=strBrowse2.split("|")
var strUser="User|用户" 
var arrUser=strUser.split("|")
var strNO="NO|序列" 
var arrNO=strNO.split("|")
var strSelect="Select|选择" 
var arrSelect=strSelect.split("|")
var strFatotory="Fatotory|工厂" 
var arrFatotory=strFatotory.split("|")
var strStaionName="Staion Name|站点名称" 
var arrStaionName=strStaionName.split("|")
var strPartName="Part Name|型号名称" 
var arrPartName=strPartName.split("|")
var strType="Type|类型" 
var arrType=strType.split("|")
var strCheckAll="Check All|全部选中" 
var arrCheckAll=strCheckAll.split("|")
var strUnCheckAll="Uncheck All|全部不选中" 
var arrUnCheckAll=strUnCheckAll.split("|")
var strUpdate=" OK |确定" 
var arrUpdate=strUpdate.split("|")
var strReset="Reset|重置" 
var arrReset=strReset.split("|")
function language()
{
try{inner_Browse.innerText=arrBrowse[<%=session("language")%>]}catch(e){}
try{inner_Browse2.innerText=arrBrowse2[<%=session("language")%>]}catch(e){}
try{inner_User.innerText=arrUser[<%=session("language")%>]}catch(e){}
try{inner_NO.innerText=arrNO[<%=session("language")%>]}catch(e){}
try{inner_Select.innerText=arrSelect[<%=session("language")%>]}catch(e){}
try{inner_Fatotory.innerText=arrFatotory[<%=session("language")%>]}catch(e){}
try{inner_StaionName.innerText=arrStaionName[<%=session("language")%>]}catch(e){}
try{inner_NO2.innerText=arrNO[<%=session("language")%>]}catch(e){}
try{inner_Select2.innerText=arrSelect[<%=session("language")%>]}catch(e){}
try{inner_Fatotory2.innerText=arrFatotory[<%=session("language")%>]}catch(e){}
try{inner_PartName.innerText=arrPartName[<%=session("language")%>]}catch(e){}
try{document.all.factory.options[0].text=arrFatotory[<%=session("language")%>]}catch(e){}
try{document.all.type.options[0].text=arrType[<%=session("language")%>]}catch(e){}
try{document.all.CheckAll.value=arrCheckAll[<%=session("language")%>]}catch(e){}
try{document.all.UnCheckAll.value=arrUnCheckAll[<%=session("language")%>]}catch(e){}
try{document.all.Update.value=arrUpdate[<%=session("language")%>]}catch(e){}
try{document.all.Reset.value=arrReset[<%=session("language")%>]}catch(e){}
}
language()
</script>
