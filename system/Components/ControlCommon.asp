<%
function getMainMenu(prefix,item_list,url_list,target_list)
if item_list<>"" then
array_mainmenu=split(item_list,"|")
array_mainurl=split(url_list,"|")
array_maintarget=split(target_list,"|")
menustring="<TABLE style=""BORDER-RIGHT: #ffffff 0px solid; BORDER-TOP: #ffffff 0px solid; BORDER-LEFT: #ffffff 0px solid; BORDER-BOTTOM: #ffffff 0px solid"" cellSpacing=""0"" cellPadding=""2"" width=""100%"" bgColor=""#ffffff""><TBODY><TR vAlign=""center""><TD align=""left""><DIV style=""FONT-WEIGHT: 700; FONT-SIZE: 12px; VISIBILITY: visible; FONT-FAMILY: Verdana; BACKGROUND-COLOR: #ffffff"" noWrap>"
	for mi=0 to ubound(array_mainmenu)
		array_submenu=split(array_mainmenu(mi),"$")
		array_suburl=split(array_mainurl(mi),"$")
		array_subtarget=split(array_maintarget(mi),"$")
		if ubound(array_submenu)>=1 then
			array_itemmenu=split(array_submenu(1),"!")
			array_itemurl=split(array_suburl(1),"!")
			array_itemtarget=split(array_subtarget(1),"!")
			'submenu div
			menustring=menustring&"<SPAN id=""menusub_div_"&prefix&mi&""" style=""VISIBILITY: hidden; POSITION: absolute; zorder: 1"">&nbsp;<BR><TABLE class=""Sub_Menu_Table"" onmouseover=""OpenMenu(menusub_div_"&prefix&mi&")"" onmouseout=""CloseMenu(menusub_div_"&prefix&mi&")"" cellSpacing=""2"" cellPadding=""2""><TBODY>"
				for ml=0 to ubound(array_itemmenu)
				menustring=menustring&"<TR><TD class=""Sub_Menu_TD"" id=""sub_td_"&prefix&mi&"_"&ml&""" onmouseover=""Sub_TD_Over(sub_td_"&prefix&mi&"_"&ml&",sub_class_"&prefix&mi&"_"&ml&")"" onmouseout=""Sub_TD_Out(sub_td_"&prefix&mi&"_"&ml&",sub_class_"&prefix&mi&"_"&ml&")""><SPAN NOWRAP><A class=""Sub_Menu_Link"" id=""sub_class_"&prefix&mi&"_"&ml&""" title="""" style=""COLOR: #ffffff"" href="""&array_itemurl(ml)&""" target="""&array_itemtarget(ml)&""">"&array_itemmenu(ml)&"</A></SPAN></TD></TR>"
				next
			menustring=menustring&"</TBODY></TABLE></SPAN>"
		end if
	
		'mainmenu item
		menustring=menustring&"<SPAN class=""Main_Menu_Span"" style=""POSITION: relative"""
		if ubound(array_submenu)>=1 then
		menustring=menustring&" onmouseover=""OpenMenu(menusub_div_"&prefix&mi&")"" onmouseout=""CloseMenu(menusub_div_"&prefix&mi&")"""
		end if
		menustring=menustring&">"
		if array_suburl(0)<>"#"	then
		menustring=menustring&"<A class=""Main_Menu_Link"" title="""" href="""&array_suburl(0)&""" target="""&array_subtarget(0)&""">&nbsp;"&array_submenu(0)&"</A>"
		else
		menustring=menustring&"&nbsp;"&array_submenu(0)
		end if
		if mi<ubound(array_mainmenu) then
		menustring=menustring&"</SPAN>&nbsp;<FONT color=""#6284a9"">|</FONT>"
		end if
	next
menustring=menustring&"</DIV></TD></TR></TBODY></TABLE>"
getMainMenu=menustring
else
getMainMenu=""
end if
end function

function Control_Tab_Select(tablist,default,selected_color,unselected_color)
if session("language")="0" then
basic_tab="&nbsp;Basic Settings&nbsp;"
content_tab="&nbsp;Content Detail&nbsp;"
authority_tab="&nbsp;Authority Settings&nbsp;"
schedule_tab="&nbsp;Schedule Settings&nbsp;"
upload_tab="&nbsp;Upload&nbsp;"
bunch_upload_tab="&nbsp;Bunch Upload&nbsp;"
authentication_tab="&nbsp;Authentication Settings&nbsp;"
finance_tab="&nbsp;Finance Settings&nbsp;"
customized_tab="&nbsp;Customized Settings&nbsp;"
scrap_tab="&nbsp;Scrap Settings&nbsp;"
explorer_tab="&nbsp;Explorer Tab&nbsp;"
list_tab="&nbsp;List Tab&nbsp;"
search_tab="&nbsp;Search Tab&nbsp;"
calendar_tab="&nbsp;Calendar Tab&nbsp;"
check_tab="&nbsp;Check Tab&nbsp;"
result_tab="&nbsp;Result Tab&nbsp;"
new_tab="&nbsp;Add Tab&nbsp;"
news_tab="&nbsp;News Query&nbsp;"
file_tab="&nbsp;File Query&nbsp;"
phone_tab="&nbsp;Phone Query&nbsp;"
software_tab="&nbsp;Software Query&nbsp;"
edit_tab="&nbsp;Edit Tab&nbsp;"
duty_tab="&nbsp;Duty Tab&nbsp;"
more_tab="&nbsp;Others Tab&nbsp;"
history_tab="&nbsp;History Tab&nbsp;"
report_tab="&nbsp;Report Tab&nbsp;"
change_tab="&nbsp;User Change Tab&nbsp;"
else
basic_tab="&nbsp;基本设置&nbsp;"
content_tab="&nbsp;具体内容&nbsp;"
authority_tab="&nbsp;权限设置&nbsp;"
schedule_tab="&nbsp;计划设置&nbsp;"
upload_tab="&nbsp;上传设置&nbsp;"
bunch_upload_tab="&nbsp;批量上传&nbsp;"
authentication_tab="&nbsp;认证设置&nbsp;"
finance_tab="&nbsp;财务设置&nbsp;"
customized_tab="&nbsp;自定义设置&nbsp;"
scrap_tab="&nbsp;报废设置&nbsp;"
explorer_tab="&nbsp;浏览框&nbsp;"
list_tab="&nbsp;列表框&nbsp;"
search_tab="&nbsp;查询框&nbsp;"
calendar_tab="&nbsp;日历框&nbsp;"
check_tab="&nbsp;检查框&nbsp;"
result_tab="&nbsp;结果框&nbsp;"
new_tab="&nbsp;新增框&nbsp;"
news_tab="&nbsp;新闻查询&nbsp;"
file_tab="&nbsp;文件查询&nbsp;"
phone_tab="&nbsp;电话查询&nbsp;"
software_tab="&nbsp;软件查询&nbsp;"
edit_tab="&nbsp;编辑框&nbsp;"
duty_tab="&nbsp;在职框&nbsp;"
more_tab="&nbsp;其他框&nbsp;"
history_tab="&nbsp;历史框&nbsp;"
report_tab="&nbsp;报表框&nbsp;"
change_tab="&nbsp;用户变更&nbsp;"
end if
if tablist<>"" then
array_tab=split(tablist,",")
Control_Tab_Select="<table border=""0"" cellspacing=""0"" cellpadding=""0"" class=""Border_Bark_Blue""><tr>"
  for j=0 to ubound(array_tab)
  	sub_array_tab=split(array_tab(j),"|")
  	select case sub_array_tab(0)
	case "TabBasic"
    Control_Tab_Select=Control_Tab_Select&"<td name=""TDTabBasic"" id=""TDTabBasic"" class="""&unselected_color&"""><div align=""center""><span style=""cursor:hand"" onClick=""ShowTab('"&tablist&"','TabBasic','"&selected_color&"','"&unselected_color&"')"">"&basic_tab&"</span></div></td>"
	case "TabContent"
    Control_Tab_Select=Control_Tab_Select&"<td name=""TDTabContent"" id=""TDTabContent"" class="""&unselected_color&"""><div align=""center""><span style=""cursor:hand"" onClick=""ShowTab('"&tablist&"','TabContent','"&selected_color&"','"&unselected_color&"')"">"&content_tab&"</span></div></td>"	
	case "TabAuthority"
	Control_Tab_Select=Control_Tab_Select&"<td name=""TDTabAuthority"" id=""TDTabAuthority"" class="""&unselected_color&"""><div align=""center""><span style=""cursor:hand"" onClick=""ShowTab('"&tablist&"','TabAuthority','"&selected_color&"','"&unselected_color&"')"">"&authority_tab&"</span></div></td>"
	case "TabSchedule"
	Control_Tab_Select=Control_Tab_Select&"<td name=""TDTabSchedule"" id=""TDTabSchedule"" class="""&unselected_color&"""><div align=""center""><span style=""cursor:hand"" onClick=""ShowTab('"&tablist&"','TabSchedule','"&selected_color&"','"&unselected_color&"')"">"&schedule_tab&"</span></div></td>"
	case "TabUpload"
	Control_Tab_Select=Control_Tab_Select&"<td name=""TDTabUpload"" id=""TDTabUpload"" class="""&unselected_color&"""><div align=""center""><span style=""cursor:hand"" onClick=""ShowTab('"&tablist&"','TabUpload','"&selected_color&"','"&unselected_color&"')"">"&upload_tab&"</span></div></td>"
	case "TabBunchUpload"
	Control_Tab_Select=Control_Tab_Select&"<td name=""TDTabBunchUpload"" id=""TDTabBunchUpload"" class="""&unselected_color&"""><div align=""center""><span style=""cursor:hand"" onClick=""ShowTab('"&tablist&"','TabBunchUpload','"&selected_color&"','"&unselected_color&"')"">"&bunch_upload_tab&"</span></div></td>"
	case "TabAuthentication"
	Control_Tab_Select=Control_Tab_Select&"<td name=""TDTabAuthentication"" id=""TDTabAuthentication"" class="""&unselected_color&"""><div align=""center""><span style=""cursor:hand"" onClick=""ShowTab('"&tablist&"','TabAuthentication','"&selected_color&"','"&unselected_color&"')"">"&authentication_tab&"</span></div></td>"
	case "TabFinance"
	Control_Tab_Select=Control_Tab_Select&"<td name=""TDTabFinance"" id=""TDTabFinance"" class="""&unselected_color&"""><div align=""center""><span style=""cursor:hand"" onClick=""ShowTab('"&tablist&"','TabFinance','"&selected_color&"','"&unselected_color&"')"">"&finance_tab&"</span></div></td>"
	case "TabCustomized"
	Control_Tab_Select=Control_Tab_Select&"<td name=""TDTabCustomized"" id=""TDTabCustomized"" class="""&unselected_color&"""><div align=""center""><span style=""cursor:hand"" onClick=""ShowTab('"&tablist&"','TabCustomized','"&selected_color&"','"&unselected_color&"')"">"&customized_tab&"</span></div></td>"
	case "TabScrap"
	Control_Tab_Select=Control_Tab_Select&"<td name=""TDTabScrap"" id=""TDTabScrap"" class="""&unselected_color&"""><div align=""center""><span style=""cursor:hand"" onClick=""ShowTab('"&tablist&"','TabScrap','"&selected_color&"','"&unselected_color&"')"">"&customized_tab&"</span></div></td>"
	case "TabExplorer"
	Control_Tab_Select=Control_Tab_Select&"<td name=""TDTabExplorer"" id=""TDTabExplorer"" class="""&unselected_color&"""><div align=""center""><span style=""cursor:hand"" onClick=""ShowTab('"&tablist&"','TabExplorer','"&selected_color&"','"&unselected_color&"')"">"&explorer_tab&"</span></div></td>"
	case "TabList"
	Control_Tab_Select=Control_Tab_Select&"<td name=""TDTabList"" id=""TDTabList"" class="""&unselected_color&"""><div align=""center""><span style=""cursor:hand"" onClick=""ShowTab('"&tablist&"','TabList','"&selected_color&"','"&unselected_color&"')"">"&list_tab&"</span></div></td>"
	case "TabSearch"
	Control_Tab_Select=Control_Tab_Select&"<td name=""TDTabSearch"" id=""TDTabSearch"" class="""&unselected_color&"""><div align=""center""><span style=""cursor:hand"" onClick=""ShowTab('"&tablist&"','TabSearch','"&selected_color&"','"&unselected_color&"')"">"&search_tab&"</span></div></td>"
	case "TabCalendar"
	Control_Tab_Select=Control_Tab_Select&"<td name=""TDTabCalendar"" id=""TDTabCalendar"" class="""&unselected_color&"""><div align=""center""><span style=""cursor:hand"" onClick=""ShowTab('"&tablist&"','TabCalendar','"&selected_color&"','"&unselected_color&"')"">"&calendar_tab&"</span></div></td>"
	case "TabNew"
	Control_Tab_Select=Control_Tab_Select&"<td name=""TDTabNew"" id=""TDTabNew"" class="""&unselected_color&"""><div align=""center""><span style=""cursor:hand"" onClick=""ShowTab('"&tablist&"','TabNew','"&selected_color&"','"&unselected_color&"')"">"&new_tab&"</span></div></td>"
	case "TabCheck"
	Control_Tab_Select=Control_Tab_Select&"<td name=""TDTabCheck"" id=""TDTabCheck"" class="""&unselected_color&"""><div align=""center""><span style=""cursor:hand"" onClick=""ShowTab('"&tablist&"','TabCheck','"&selected_color&"','"&unselected_color&"')"">"&check_tab&"</span></div></td>"
	case "TabResult"
	Control_Tab_Select=Control_Tab_Select&"<td name=""TDTabResult"" id=""TDTabResult"" class="""&unselected_color&"""><div align=""center""><span style=""cursor:hand"" onClick=""ShowTab('"&tablist&"','TabResult','"&selected_color&"','"&unselected_color&"')"">"&result_tab&"</span></div></td>"
	case "TabNews"
	Control_Tab_Select=Control_Tab_Select&"<td name=""TDTabNews"" id=""TDTabNews"" class="""&unselected_color&"""><div align=""center""><span style=""cursor:hand"" onClick=""ShowTab('"&tablist&"','TabNews','"&selected_color&"','"&unselected_color&"')"">"&news_tab&"</span></div></td>"
	case "TabPhone"
	Control_Tab_Select=Control_Tab_Select&"<td name=""TDTabPhone"" id=""TDTabPhone"" class="""&unselected_color&"""><div align=""center""><span style=""cursor:hand"" onClick=""ShowTab('"&tablist&"','TabPhone','"&selected_color&"','"&unselected_color&"')"">"&phone_tab&"</span></div></td>"
	case "TabFile"
	Control_Tab_Select=Control_Tab_Select&"<td name=""TDTabFile"" id=""TDTabFile"" class="""&unselected_color&"""><div align=""center""><span style=""cursor:hand"" onClick=""ShowTab('"&tablist&"','TabFile','"&selected_color&"','"&unselected_color&"')"">"&file_tab&"</span></div></td>"
	case "TabSoftware"
	Control_Tab_Select=Control_Tab_Select&"<td name=""TDTabSoftware"" id=""TDTabSoftware"" class="""&unselected_color&"""><div align=""center""><span style=""cursor:hand"" onClick=""ShowTab('"&tablist&"','TabSoftware','"&selected_color&"','"&unselected_color&"')"">"&software_tab&"</span></div></td>"
	case "TabEdit"
	Control_Tab_Select=Control_Tab_Select&"<td name=""TDTabEdit"" id=""TDTabEdit"" class="""&unselected_color&"""><div align=""center""><span style=""cursor:hand"" onClick=""ShowTab('"&tablist&"','TabEdit','"&selected_color&"','"&unselected_color&"')"">"&edit_tab&"</span></div></td>"
	case "TabDuty"
	Control_Tab_Select=Control_Tab_Select&"<td name=""TDTabDuty"" id=""TDTabDuty"" class="""&unselected_color&"""><div align=""center""><span style=""cursor:hand"" onClick=""ShowTab('"&tablist&"','TabDuty','"&selected_color&"','"&unselected_color&"')"">"&duty_tab&"</span></div></td>"
	case "TabMore"
	Control_Tab_Select=Control_Tab_Select&"<td name=""TDTabMore"" id=""TDTabMore"" class="""&unselected_color&"""><div align=""center""><span style=""cursor:hand"" onClick=""ShowTab('"&tablist&"','TabMore','"&selected_color&"','"&unselected_color&"')"">"&more_tab&"</span></div></td>"
	case "TabHistory"
	Control_Tab_Select=Control_Tab_Select&"<td name=""TDTabHistory"" id=""TDTabHistory"" class="""&unselected_color&"""><div align=""center""><span style=""cursor:hand"" onClick=""ShowTab('"&tablist&"','TabHistory','"&selected_color&"','"&unselected_color&"')"">"&history_tab&"</span></div></td>"
	case "TabReport"
	Control_Tab_Select=Control_Tab_Select&"<td name=""TDTabReport"" id=""TDTabReport"" class="""&unselected_color&"""><div align=""center""><span style=""cursor:hand"" onClick=""ShowTab('"&tablist&"','TabReport','"&selected_color&"','"&unselected_color&"')"">"&report_tab&"</span></div></td>"
	case "TabChange"
	Control_Tab_Select=Control_Tab_Select&"<td name=""TDTabChange"" id=""TDTabChange"" class="""&unselected_color&"""><div align=""center""><span style=""cursor:hand"" onClick=""ShowTab('"&tablist&"','TabChange','"&selected_color&"','"&unselected_color&"')"">"&Change_tab&"</span></div></td>"
	end select
	next
  Control_Tab_Select=Control_Tab_Select&"</tr></table>"
end if
end function

function Control_From_To(fromid,toid,fromoptions,tooptions,onchange)
if session("language")="0" then
available_options="Available Options"
selected_options="Selected Options"
add="Add"
move_up="Move Up"
remove="Remove"
move_down="Move Down"
add_all="Add All"
to_top="To Top"
remove_all="Remove All"
to_bottom="To Bottom"
else
available_options="可选项"
selected_options="已选项"
add="加入"
move_up="上移"
remove="移去"
move_down="下移"
add_all="加入所有"
to_top="置顶"
remove_all="移去所有"
to_bottom="置底"
end if
Control_From_To="<table border=""1"" align=""left"" cellpadding=""0"" cellspacing=""0""><tr><td class=""TD_Dark_Silver""><div align=""center"">"&available_options&"<span id="""&fromid&"_count""></span></div></td><td><div align=""center"">&nbsp;</div></td><td class=""TD_Dark_Silver""><div align=""center"">"&selected_options&"<span id="""&toid&"_count""></span></div></td><td><div align=""center"">&nbsp;</div></td></tr><tr><td rowspan=""4""><div align=""center""><table border=""0"" cellspacing=""0"" cellpadding=""0"" class=""Border_None""><tr><td rowspan=""3""><select name="""&fromid&""" size=""6"" multiple id="""&fromid&""">"&fromoptions&"</select></td><td valign=""top""><img src=""/Images/SpaceUp.gif"" width=""14"" height=""8"" style=""cursor:hand"" onClick=""javascript:if(document.all."&fromid&".size>3){document.all."&fromid&".size=document.all."&fromid&".size-3;document.all."&toid&".size=document.all."&toid&".size-3}""/></td></tr><tr><td>&nbsp;</td></tr><tr><td valign=""bottom""><img src=""/Images/SpaceDown.gif"" width=""14"" height=""8"" style=""cursor:hand"" onClick=""javascript:document.all."&fromid&".size=document.all."&fromid&".size+3;document.all."&toid&".size=document.all."&toid&".size+3""/></td></tr><tr><td>"&Control_Input(fromid&"_focus",20,null,search_default,false,"select_focus(document.all."&fromid&",'"&selected&"',this.value)",false)&"</td><td>&nbsp;</td></tr></table></div></td><td><div align=""center""><input name=""FromToAdd"" type=""button"" value="""&add&""" onClick=""item_move(document.all."&fromid&",document.all."&toid&");options_count(document.all."&fromid&",document.all."&toid&",document.all."&fromid&"_count,document.all."&toid&"_count);"&onchange&"""></div></td><td rowspan=""4""><div align=""center""><table border=""0"" cellspacing=""0"" cellpadding=""0"" class=""Border_None""><tr><td rowspan=""3""><select name="""&toid&""" size=""6"" multiple id="""&toid&""">"&tooptions&"</select></td><td valign=""top""><img src=""/Images/SpaceUp.gif"" width=""14"" height=""8"" style=""cursor:hand"" onClick=""javascript:if(document.all."&toid&".size>3){document.all."&fromid&".size=document.all."&fromid&".size-3;document.all."&toid&".size=document.all."&toid&".size-3}""/></td></tr><tr><td>&nbsp;</td></tr><tr><td valign=""bottom""><img src=""/Images/SpaceDown.gif"" width=""14"" height=""8"" style=""cursor:hand"" onClick=""javascript:document.all."&fromid&".size=document.all."&fromid&".size+3;document.all."&toid&".size=document.all."&toid&".size+3""/></td></tr><tr><td>"&Control_Input(toid&"_focus",20,null,search_default,false,"select_focus(document.all."&toid&",'"&selected&"',this.value)",false)&"</td><td>&nbsp;</td></tr></table></div></td><td><div align=""center""><input name=""FromToUp"" type=""button"" value="""&move_up&""" onClick=""item_up(document.all."&toid&")""></div></td></tr><tr><td><div align=""center""><input name=""FromToRemove"" type=""button"" value="""&remove&""" onClick=""item_move(document.all."&toid&",document.all."&fromid&");options_count(document.all."&fromid&",document.all."&toid&",document.all."&fromid&"_count,document.all."&toid&"_count);"&onchange&"""></div></td><td><div align=""center""><input name=""FromToDown"" type=""button"" value="""&move_down&""" onClick=""item_down(document.all."&toid&")""> </div></td></tr><tr><td><div align=""center""> <input name=""FromToAddAll"" type=""button"" value="""&add_all&""" onClick=""item_all(document.all."&fromid&",document.all."&toid&");options_count(document.all."&fromid&",document.all."&toid&",document.all."&fromid&"_count,document.all."&toid&"_count);"&onchange&"""></div></td><td><div align=""center""><input name=""FromToToTop"" type=""button"" value="""&to_top&"""  onClick=""item_top(document.all."&toid&")""></div></td></tr><tr><td><div align=""center""><input name=""FromToRemoveAll"" type=""button"" value="""&remove_all&""" onClick=""item_all(document.all."&toid&",document.all."&fromid&");options_count(document.all."&fromid&",document.all."&toid&",document.all."&fromid&"_count,document.all."&toid&"_count);"&onchange&"""></div></td><td><div align=""center""><input name=""FromToToButton"" type=""button"" value="""&to_bottom&"""  onClick=""item_bottom(document.all."&toid&")""></div></td></tr></table>"
end function

function Control_Input(id,length,default,title,disabled,onchange,capitalized)
if session("language")="0" then
capitalize_title="Capitalize words"
uncapitalize_title="Initialize words"
else
capitalize_title="字母大写"
uncapitalize_title="字母小写"
end if
Control_Input="<input name="""&id&""" id="""&id&""" type=""text"" size="""&length&""" value="""&default&"""  "
if disabled=true then
Control_Input=Control_Input&" disabled "
end if
Control_Input=Control_Input&" title="""&title&""" onChange="""&onchange&""""
if capitalized=true then
Control_Input=Control_Input&" onClick=""show_capitalize('"&id&"_capitalize_div')"""
end if
Control_Input=Control_Input&" onfocus=""focushandler(this)"" onblur=""blurhandler(this)"" />"
if capitalized=true then
Control_Input=Control_Input&"<div id="""&id&"_capitalize_div"" style=""visibility:hidden;position:absolute"">"&Control_Image(id&"_capitalize","/Capitalize.gif",18,18,capitalize_title,"absmiddle","capitalize('"&id&"')")&Control_Image(id&"_uncapitalize","/UnCapitalize.gif",18,18,uncapitalize_title,"absmiddle","uncapitalize('"&id&"')")&"</div>"
end if
end function

function Control_Input_NoCapitalize(id,length,default,title,disabled,onchange)
Control_Input_NoCapitalize="<input name="""&id&""" id="""&id&""" type=""text"" size="""&length&""" value="""&default&"""  "
if disabled=true then
Control_Input_NoCapitalize=Control_Input_NoCapitalize&"disabled "
end if
Control_Input_NoCapitalize=Control_Input_NoCapitalize&" title="""&title&""" onChange="""&onchange&""""
end function

function Control_Input_Hidden(id,length,default,disabled,onchange)
Control_Input_Hidden="<input name="""&id&""" id="""&id&""" type=""hidden"" size="""&length&""" value="""&default&"""  "
if disabled=true then
Control_Input_Hidden=Control_Input_Hidden&"disabled "
end if
Control_Input_Hidden=Control_Input_Hidden&"onChange="""&onchange&""" onfocus=""focushandler(this)"" onblur=""blurhandler(this)"" />"
end function

function Control_Password(id,length,default,disabled)
Control_Password="<input name="""&id&""" id="""&id&""" type=""password"" size="""&length&""" value="""&default&"""  "
if disabled=true then
Control_Password=Control_Password&"disabled "
end if
Control_Password=Control_Password&"onfocus=""focushandler(this)"" onblur=""blurhandler(this)"" />"
end function

function Control_Button(id,text,disabled,onclick)
Control_Button="<input name="""&id&""" type=""button"" "
if disabled=true then
Control_Button=Control_Button&"disabled "
end if
Control_Button=Control_Button&"value="""&text&""" onfocus=""focushandler(this)"" onblur=""blurhandler(this)"" onClick="""&onclick&""">"
end function

function Control_SubmitButton(id,text,disabled,action)
Control_SubmitButton="<input name="""&id&""" type=""submit"" "
if disabled=true then
Control_SubmitButton=Control_SubmitButton&"disabled "
end if
Control_SubmitButton=Control_SubmitButton&"value="""&text&""" onfocus=""focushandler(this)"" onblur=""blurhandler(this)"" onClick="""&action&""">"
end function

function Control_ResetButton(id,text,disabled,action)
Control_ResetButton="<input name="""&id&""" type=""reset"" "
if disabled=true then
Control_ResetButton=Control_ResetButton&"disabled "
end if
Control_ResetButton=Control_ResetButton&"value="""&text&""" onfocus=""focushandler(this)"" onblur=""blurhandler(this)"" onClick="""&action&""">"
end function

function Control_Simple_Select(id,default,options,selected,onChange,disabled)
if session("language")="0" then
search_default="Quick Search"
else
search_default="快速查询"
end if
Control_Simple_Select="<select name="""&id&"""" 
if disabled=true then
Control_Simple_Select=Control_Simple_Select&" disabled "
end if
Control_Simple_Select=Control_Simple_Select&" onfocus=""focushandler(this)"" onblur=""blurhandler(this)"" onChange="""&onChange&"""><option value="""">"&default&"</option>"&options&"</select>"
end function

function Control_Select(id,default,options,selected,onChange,disabled)
if session("language")="0" then
search_default="Quick Search"
else
search_default="快速查询"
end if
Control_Select="<select name="""&id&"""" 
if disabled=true then
Control_Select=Control_Select&" disabled "
end if
Control_Select=Control_Select&" onfocus=""focushandler(this)"" onblur=""blurhandler(this)"" onChange="""&onChange&"""><option value="""">"&default&"</option>"&options&"</select>"&Control_Input(id&"_focus",6,null,search_default,false,"select_focus(document.all."&id&",document.all."&id&"_backup,'"&selected&"',this.value);"&onChange,false)&"<span style=""visibility:hidden;position:absolute""><select name="""&id&"_backup"">"&options&"</select></span>"
end function

function Control_TimeUnit_Select(timeid,unitid,prefix_type,time_number,selected_unit,time_disabled,unit_disabled)
if session("language")="0" then
remind_before_text="Remind Before"
timeunit_default="-- Select Time Unit --"
array_timeunit=split("Minutes,Hours,Days",",")
else
before_text="在之前提醒"
timeunit_default="-- 时间单位 --"
array_timeunit=split("分钟,小时,日",",")
end if
array_timeunit_value=split("MM,HH,DD",",")
select case prefix_type
case "REMIND"
prefix=remind_before_text
case null
prefix=""
end select

for j=0 to ubound(array_timeunit_value)
uint_options=uint_options&"<option value="""&array_timeunit_value(j)&""" "
	if selected_unit=array_timeunit_value(j) then
	uint_options=uint_options&"selected "
	end if
uint_options=uint_options&">"&array_timeunit(j)&"</option>"
next

Control_TimeUnit_Select=prefix&Control_Input(timeid,4,time_number,null,time_disabled,"numbercheck(this)",false)&Control_Select(unitid,timeunit_default,uint_options,selected_unit,onChange,unit_disabled)
end function

function Control_Radio(id,check,checkvalue,text,onClick,disabled)
Control_Radio="<input type=""radio"" name="""&id&""" value="""&checkvalue&""" "
if check=true then
Control_Radio=Control_Radio&"checked "
end if
if disabled=true then
Control_Radio=Control_Radio&"disabled "
end if
Control_Radio=Control_Radio&"onclick="""&onClick&""" onfocus=""focushandler(this)"" onblur=""blurhandler(this)"" />"&text
end function

function Control_Check(id,disabled,check,checkvalue,text,onClick)
Control_Check="<input type=""checkbox"" name="""&id&""" value="""&checkvalue&""" "
if check=true then
Control_Check=Control_Check&"checked "
end if
if disabled=true then
Control_Check=Control_Check&"disabled "
end if
Control_Check=Control_Check&" onclick="""&onClick&""" onfocus=""focushandler(this)"" onblur=""blurhandler(this)"" />"&text
end function

function Control_Textarea(id,length,high,text)
Control_Textarea="<textarea name="""&id&""" id="""&id&""" cols="""&length&""" rows="""&high&""" onfocus=""focushandler(this)"" onblur=""blurhandler(this)"">"&text&"</textarea>"
end function

function Control_File(id,length)
Control_File="<input name="""&id&""" type=""file"" size="""&length&"""  onfocus=""focushandler(this)"" onblur=""blurhandler(this)"" />"
end function

function Control_YearPicker(id,selected_year,fromyear,toyear,onChange)
if session("language")="0" then
year_default="-- Select Year --"
year_text="&nbsp;Year&nbsp;"
else
year_default="-- 选择年份 --"
year_text="&nbsp;年&nbsp;"
end if
Control_YearPicker="<select name="""&id&""" id="""&id&""" onFocus=""focushandler(this)"" onBlur=""blurhandler(this)"" onChange="""&onChange&"""><option value="""">"&year_default&"</option>"
for ci=fromyear to toyear
	Control_YearPicker=Control_YearPicker&"<option value="""&ci&""""
	if ci=selected_year then
	Control_YearPicker=Control_YearPicker&"selected "
	end if
	Control_YearPicker=Control_YearPicker&">"&ci&"</option>"
next
Control_YearPicker=Control_YearPicker&"</select>"&year_text
end function

function Control_MonthPicker(id,selected_month,onChange)
if session("language")="0" then
month_default="-- Select Month --"
month_text="&nbsp;Month&nbsp;"
else
month_default="-- 选择月份 --"
month_text="&nbsp;月&nbsp;"
end if
Control_MonthPicker="<select name="""&id&""" id="""&id&""" onFocus=""focushandler(this)"" onBlur=""blurhandler(this)"" onChange="""&onChange&"""><option value="""">"&month_default&"</option>"
for ci=1 to 12
	Control_MonthPicker=Control_MonthPicker&"<option value="""&ci&""""
	if ci=selected_month then
	Control_MonthPicker=Control_MonthPicker&"selected "
	end if
	Control_MonthPicker=Control_MonthPicker&">"&monthconvert(ci)&"</option>"
next
Control_MonthPicker=Control_MonthPicker&"</select>"&month_text
end function

function Control_DatePicker(id,length,default_date,readonly)
Control_DatePicker="<input name="""&id&""" type=""text"" id="""&id&""" onFocus=""focushandler(this)"" onBlur=""blurhandler(this)"" value="""&default_date&""" size="""&length&""" readonly="""&readonly&"""><script language=""JavaScript"" type=""text/javascript"">function calendar_"&id&"_Callback(date, month, year){document.all."&id&".value=year + '-' + month + '-' + date;}calendar_"&id&"=new dynCalendar('calendar_"&id&"', 'calendar_"&id&"_Callback');</script>"
end function

function Control_TimePicker(hourid,minuteid,hour_default,minute_default,hour_disabled,minute_disabled,hour_onchange,minute_onchange)
if session("language")="0" then
text_hour=" hour "
text_minute=" minute "
else
text_hour=" 时 "
text_minute=" 分 "
end if
Control_TimePicker=Control_Input(hourid,"2",hour_default,null,hour_disabled,hour_onchange,false)&text_hour&Control_Input(minuteid,"2",minute_default,null,minute_disabled,minute_onchange,false)&text_minute
end function

function Control_Currency_select(id,default)
currency_string="RMB,USD,HKD"
if session("language")="0" then
currency_default="-- Select Currency --"
else
currency_default="-- 选择货币 --"
end if
a_currency_string=split(currency_string,",")
Control_Currency_select="<select name="""&id&""" id="""&id&""" onFocus=""focushandler(this)"" onBlur=""blurhandler(this)""><option value="""">"&currency_default&"</option>"
	for ci=0 to ubound(a_currency_string)
		Control_Currency_select=Control_Currency_select&"<option value="""&a_currency_string(ci)&""""
		if a_currency_string(ci)=default then
		Control_Currency_select=Control_Currency_select&" selected"
		end if
		Control_Currency_select=Control_Currency_select&">"&a_currency_string(ci)&"</option>"
	next
	Control_Currency_select=Control_Currency_select&"</select>"
end function

function richtext(id,body)
richtext="<script language=""JavaScript"" src=""/Components/RichText/js/richtext.js"" type=""text/javascript""></script><script language=""JavaScript"" src=""/Components/RichText/js/config.js"" type=""text/javascript""></script><script>initRTE( '"&id&"','"&body&"');</script>"
end function

function Control_Image(id,src,iwidth,iheight,title,align,onClick)
Control_Image="<img src=""/Images/"&src&""" width="""&iwidth&""" height="""&iheight&""" align="""&align&""" title="""&title&""" onClick="""&onClick&""" border=0 style=""cursor:hand""/>"
end function

Function Control_IP_Address(id,default)
if default<>"" then
IP_default=split(default,".")
else
IP_default=split("10.12.100.0",".")
end if
Control_IP_Address="<table border=""0"" cellspacing=""0"" cellpadding=""0""><tr><td><div align=""center"">"&Control_Input(id&"1",3,IP_default(0),null,false,"ip_check(this,"&id&"2,'"&IP_default(0)&"')",false)&"."&Control_Input(id&"2",3,IP_default(1),null,false,"ip_check(this,"&id&"3,'"&IP_default(1)&"')",false)&"."&Control_Input(id&"3",3,IP_default(2),null,false,"ip_check(this,"&id&"4,'"&IP_default(2)&"')",false)&"."&Control_Input(id&"4",3,IP_default(3),null,false,"ip_check(this,'','"&IP_default(3)&"')",false)&"</div></td></tr></table>"
end function

function Control_iFrame(id,iwidth,iheight,src,border)
Control_iFrame="<iframe name="""&id&""" id="""&id&""" src="""&src&""" width="""&iwidth&""" height="""&iheight&""" scrolling=""auto"" frameborder="""&border&"""></iframe>"
end function

function Control_Report_Interface(report_type,available_column_options,select_column_options)
if session("language")="0" then
record_type="Records Type"
recode_all="All Records"
recode_current="Current Searched Records"
report_column="Report Columns"
report_format="Report Format"
report_type_default="-- Select Report Format Type --"
report_button="Generate Report"
else
record_type="记录类型"
recode_all="所有记录"
recode_current="当前查询的记录"
report_column="报告内容"
report_format="报告格式"
report_type_default="-- 选择报告格式类型 --"
report_button="产生报告"
end if

if isnull(available_column_options)=false and available_column_options<>"" then
array_option=split(available_column_options,",")
	for rop=0 to ubound(array_option)
		array_option_string=split(array_option(rop),"|")
		option_value=array_option_string(0)
		array_option_text=split(array_option_string(1),"*")
		if session("language")="0" then
		option_text=array_option_text(0)
		else
		option_text=array_option_text(1)
		end if
		this_available_column_options=this_available_column_options&"<option value="""&option_value&""">"&option_text&"</option>"
	next
end if

type_string="Excel,Acrobat,XML"
type_value="0,1,2"
array_type_string=split(type_string,",")
array_type_value=split(type_value,",")
for rt=0 to ubound(array_type_string)
type_options=type_options&"<option value="""&array_type_value(rt)&""">"&array_type_string(rt)&"</option>"
next

Control_Report_Interface="<form action=""/Functions/RecordsExportFile.asp"" method=""post"" name=""form_report"" target=""_blank""><table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0""><tr><td>"&record_type&"</td><td>"&Control_Radio("record_type",true,"0",recode_all,null,false)&" "&Control_Radio("record_type",false,"1",recode_current,null,false)&"</td></tr><tr><td>"&report_column&"</td><td>"&Control_From_To("available_column","selected_column",this_available_column_options,this_select_column_options,null)&"</td></tr><tr><td>"&report_format&"</td><td>"&Control_Select("report_type",report_type_default,type_options,report_type,null,false)&"</td></tr><tr><td colspan=""2""><div align=""center"">"&Control_Button("report",report_button,disabled,"if (report_condition_check(document.form_report.selected_column)){document.form_report.submit()}")&"</div></td></tr></table></form>"
end function

if session("language")="0" then
my_hotlinks="My Hotlinks"
add_my_hotlink="Add to My Hotlinks"
remove_my_hotlink="Remove from My Hotlinks"
china_time_default="Beijing Time:"
world_city_default="World Time"
hot_links_default="-- Useful Links --"
factory_default="-- Select Factory --"
building_default="-- Select Building --"
department_default="-- Select Department --"
category_default="-- Select Category --"
device_default="-- Select Device --"
device_type_default="-- Select Device Type --"
device_group_default="-- Select Device Group --"
service_agent_default="-- Select Service Agent --"
depreciation_default="-- Select Depreciation Type --"
fixed_control_default="-- Select Control Type --"
fatherfolder_default="-- Select Father Folder --"
folder_default="-- Select Folder --"
log_event_default="-- Select Log Event --"
log_status_default="-- Select Log Status --"
authority_template_default="-- Select Authority Template --"
template_default="-- Select Template --"
exceedtime_default="-- Select Minutes --"
person_default="-- Select Person --"
buyer_default="-- Select Buyer --"
transactor_default="-- Select Transactor --"
transactreason_default="-- Select Transact Reason--"
meeting_room_default="-- Select Meeting Room --"
schedule_type_default="-- Select Schedule Type --"
order_status_default="-- Select Order Status --"
poll_template_default="-- Select Poll Template --"
currency_default="-- Select Currency --"
add_new_reference_folder="Add New Reference Folder"
edit_reference_folder="Edit Reference Folder"
delete_reference_folder="Delete Reference Folder"
add_new_reference="Add New Reference"
else
my_hotlinks="我的快捷方式"
add_my_hotlink="添加到我的快捷方式"
remove_my_hotlink="从我的快捷方式中删除"
china_time_default="北京时间："
world_city_default="世界时间"
hot_links_default="-- 常用链接 --"
factory_default="-- 选择工厂 --"
building_default="-- 选择厂房 --"
department_default="-- 选择部门 --"
category_default="-- 选择类别 --"
device_default="-- 选择设备 --"
device_type_default="-- 选择设备类型 --"
device_group_default="-- 选择设备组 --"
service_agent_default="-- 选择服务商 --"
depreciation_default="-- 选择报废类型 --"
fixed_control_default="-- 选择控制类型 --"
fatherfolder_default="-- 选择父文件夹 --"
folder_default="-- 选择文件夹 --"
log_event_default="-- 选择日志事件 --"
log_status_default="-- 选择日志状态 --"
authority_template_default="-- 选择权限模版 --"
template_default="-- 选择模版 --"
exceedtime_default="-- 选择分钟--"
person_default="-- 选择员工 --"
buyer_default="-- 选择采购人 --"
transactor_default="-- 选择处理人 --"
transactreason_default="-- 选择处理原因--"
meeting_room_default="-- 选择会议室 --"
schedule_type_default="-- 选择计划类型 --"
order_status_default="-- 选择单子类型 --"
poll_template_default="-- 选择投票模版 --"
currency_default="-- 选择货币 --"
add_new_reference_folder="添加新的资料夹"
edit_reference_folder="修改资料夹"
delete_reference_folder="删除资料夹"
add_new_reference="添加新的资料"
end if
%>
