<script language="javascript">
var strBrowse="Edit Operator's (|�����������" 
var arrBrowse=strBrowse.split("|")
var strBrowse2=") Authority|����Ȩ��" 
var arrBrowse2=strBrowse2.split("|")
var strUser="User|�û�" 
var arrUser=strUser.split("|")
var strNO="NO|����" 
var arrNO=strNO.split("|")
var strSelect="Select|ѡ��" 
var arrSelect=strSelect.split("|")
var strFatotory="Fatotory|����" 
var arrFatotory=strFatotory.split("|")
var strStaionName="Staion Name|վ������" 
var arrStaionName=strStaionName.split("|")
var strPartName="Part Name|�ͺ�����" 
var arrPartName=strPartName.split("|")
var strType="Type|����" 
var arrType=strType.split("|")
var strCheckAll="Check All|ȫ��ѡ��" 
var arrCheckAll=strCheckAll.split("|")
var strUnCheckAll="Uncheck All|ȫ����ѡ��" 
var arrUnCheckAll=strUnCheckAll.split("|")
var strUpdate=" OK |ȷ��" 
var arrUpdate=strUpdate.split("|")
var strReset="Reset|����" 
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
