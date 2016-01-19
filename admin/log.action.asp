<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="cms.header.asp" -->
<%
dim action : action = e.get("action")

select case lcase(action)
	
	case "del"
		call log.del()

	case "clear"
		call log.clear()

	case "ajax"
		call log.ajax_content()

end select

set log = nothing
redirect("log.listview.asp")

%>
<!--#include file="cms.footer.asp" -->
