<%
function getCategory(showstyle,categoryid,where,order)
	'isshowname=whether to output name of station;
	'showstyle=what's the ouput style:
	'	OPTION=select component
	'	TEXT=text words
	'where=where phrase of SQL query
	'splitchar=if ouput need to be split, what's character displayed as split
output=""
set rsS=server.CreateObject("adodb.recordset")
z=0
SQLS="SELECT CATEGORY_ID, CATEGORY_NAME, INVENTORY_NAME, KITTING, ISSUE FROM MATERIAL_CATEGORY  "&where& order
 
rsS.open SQLS,conn,1,3

if not rsS.eof then
			while not rsS.eof
					select case showstyle
					case "OPTION"
						output=output&"<option value='"&rsS("CATEGORY_ID")&"'"
							if rsS("CATEGORY_ID")=categoryid then
							output=output&"selected"
							end if
						output=output&">"&rsS("CATEGORY_NAME")&"</option>"
					end select
				rsS.movenext
			wend
			
			
end if

getCategory=output
rsS.close
set rsS=nothing
end function
 
%>