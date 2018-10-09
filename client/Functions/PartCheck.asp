<%
function PartCheck(validvalue,inputvalue)
	PartCheck=false
	valid=0
	if validvalue<>"" then
		if len(validvalue)=len(inputvalue) then
			for l=1 to len(validvalue)
				this_validchar=mid(validvalue,l,1)
				this_putchar=mid(inputvalue,l,1)
					if this_validchar<>"*" then
						if this_validchar=this_putchar then
							valid=valid+1
						end if
					else
					valid=valid+1
					end if
			next
		end if
	end if
	if valid=len(validvalue) then
		PartCheck=true
	else
		PartCheck=false
	end if
end function
%>