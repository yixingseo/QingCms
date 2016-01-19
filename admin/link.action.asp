<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="cms.header.asp" -->
<!--#include file="../include/feedback.class.asp" -->
<!--#include file="../include/link.class.asp" -->

<%
dim action : action = e.get("action")
select case lcase(action)
	case "create"
		call link.create()
	case "update"
		call link.update()
	case "delete"
		call link.delete()
end select
set link = nothing
redirect "link.listview.asp"
%>

<!--#include file="cms.footer.asp" -->
