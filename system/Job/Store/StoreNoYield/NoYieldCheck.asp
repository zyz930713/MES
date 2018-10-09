<script language="javascript">
function NoYieldcheck(ob,thisvalue)
{
	if (ob.checked)
	{
	location.href="NoYield_Check.asp?nid="+thisvalue+"&path=<%=path%>&query=<%=query%>"
	}
	else
	{
	location.href="NoYield_Uncheck.asp?nid="+thisvalue+"&path=<%=path%>&query=<%=query%>"
	}
}
</script>
